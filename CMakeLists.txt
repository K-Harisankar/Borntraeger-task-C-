cmake_minimum_required(VERSION 3.21)
project(CsvBakeryImporter LANGUAGES CXX)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Create executable
add_executable(${PROJECT_NAME} WIN32 src/main.cpp)

# Try to find SQLite3 via vcpkg first
find_package(unofficial-sqlite3 CONFIG QUIET)
if(unofficial-sqlite3_FOUND)
    target_link_libraries(${PROJECT_NAME} PRIVATE unofficial::sqlite3::sqlite3)
    message(STATUS "Found SQLite3 via vcpkg")
else()
    # Fallback: try system SQLite3 (prefer static)
    find_library(SQLITE3_LIBRARY NAMES libsqlite3.a sqlite3 NAMES_PER_DIR)
    find_path(SQLITE3_INCLUDE_DIR sqlite3.h)
    
    if(SQLITE3_LIBRARY AND SQLITE3_INCLUDE_DIR)
        target_link_libraries(${PROJECT_NAME} PRIVATE ${SQLITE3_LIBRARY})
        target_include_directories(${PROJECT_NAME} PRIVATE ${SQLITE3_INCLUDE_DIR})
        message(STATUS "Found system SQLite3: ${SQLITE3_LIBRARY}")
    else()
        # Download SQLite3 amalgamation if not found
        message(STATUS "SQLite3 not found, downloading amalgamation...")
        
        set(SQLITE_URL "https://www.sqlite.org/2023/sqlite-amalgamation-3430200.zip")
        set(SQLITE_ZIP "${CMAKE_BINARY_DIR}/sqlite.zip")
        set(SQLITE_DIR "${CMAKE_BINARY_DIR}/sqlite-amalgamation-3430200")
        
        if(NOT EXISTS ${SQLITE_DIR})
            file(DOWNLOAD ${SQLITE_URL} ${SQLITE_ZIP}
                SHOW_PROGRESS
                STATUS DOWNLOAD_STATUS)
            
            list(GET DOWNLOAD_STATUS 0 STATUS_CODE)
            if(NOT STATUS_CODE EQUAL 0)
                message(FATAL_ERROR "Failed to download SQLite3")
            endif()
            
            execute_process(
                COMMAND ${CMAKE_COMMAND} -E tar xf ${SQLITE_ZIP}
                WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
                RESULT_VARIABLE EXTRACT_RESULT
            )
            
            if(NOT EXTRACT_RESULT EQUAL 0)
                message(FATAL_ERROR "Failed to extract SQLite3")
            endif()
        endif()
        
        # Add SQLite3 source to our executable
        target_sources(${PROJECT_NAME} PRIVATE ${SQLITE_DIR}/sqlite3.c)
        target_include_directories(${PROJECT_NAME} PRIVATE ${SQLITE_DIR})
        target_compile_definitions(${PROJECT_NAME} PRIVATE 
            SQLITE_ENABLE_FTS4 
            SQLITE_ENABLE_RTREE
        )
        message(STATUS "Using SQLite3 amalgamation")
    endif()
endif()

# Windows-specific libraries
if(WIN32)
    target_link_libraries(${PROJECT_NAME} PRIVATE 
        comctl32
        comdlg32
        shell32
        ole32
    )
endif()

# Set working directory for debugging
set_property(TARGET ${PROJECT_NAME} PROPERTY VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_SOURCE_DIR}")

# For MinGW, link filesystem library and static runtime
if(MINGW)
    target_link_libraries(${PROJECT_NAME} PRIVATE stdc++fs)
    # Static linking to avoid DLL dependencies
    target_link_options(${PROJECT_NAME} PRIVATE 
        -static-libgcc 
        -static-libstdc++ 
        -static
        -Wl,-Bstatic
    )
endif()

message(STATUS "Build configured for Win32 GUI application")