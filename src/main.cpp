// Emergency Win32 GUI - Bakery CSV Import Tool
// Replace your src/main.cpp with this code

#include <windows.h>
#include <commdlg.h>
#include <commctrl.h>
#include <shlobj.h>
#include <sqlite3.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <string>
#include <filesystem>
#include <thread>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")

namespace fs = std::filesystem;

// Window controls IDs
#define ID_CSV_PATH_EDIT 1001
#define ID_DB_PATH_EDIT 1002
#define ID_CSV_BROWSE_BTN 1003
#define ID_DB_BROWSE_BTN 1004
#define ID_IMPORT_BTN 1005
#define ID_PROGRESS_BAR 1006
#define ID_LOG_EDIT 1007
#define ID_EXIT_BTN 1008

// Global variables
HWND g_hMainWindow = nullptr;
HWND g_hCsvPathEdit = nullptr;
HWND g_hDbPathEdit = nullptr;
HWND g_hProgressBar = nullptr;
HWND g_hLogEdit = nullptr;
HWND g_hImportBtn = nullptr;
bool g_importInProgress = false;

// Helper function to add log messages
void AddLogMessage(const std::string &message)
{
    std::string timestamp = std::to_string(GetTickCount() / 1000);
    std::string logEntry = "[" + timestamp + "s] " + message + "\r\n";

    int length = GetWindowTextLengthA(g_hLogEdit);
    SendMessageA(g_hLogEdit, EM_SETSEL, length, length);
    SendMessageA(g_hLogEdit, EM_REPLACESEL, FALSE, (LPARAM)logEntry.c_str());
    SendMessageA(g_hLogEdit, EM_SCROLLCARET, 0, 0);
}

// Helper function to convert ISO-8859-1 to UTF-8
std::string ConvertISO88591ToUTF8(const std::string &iso_string)
{
    std::string utf8_string;
    utf8_string.reserve(iso_string.size() * 2);

    for (unsigned char c : iso_string)
    {
        if (c < 0x80)
        {
            utf8_string += static_cast<char>(c);
        }
        else
        {
            utf8_string += static_cast<char>(0xC0 | (c >> 6));
            utf8_string += static_cast<char>(0x80 | (c & 0x3F));
        }
    }
    return utf8_string;
}

// Helper function to process decimal values
std::string ProcessDecimalValue(const std::string &value)
{
    if (value.empty())
        return value;

    std::string processed = value;
    size_t commaPos = processed.find(',');
    while (commaPos != std::string::npos)
    {
        processed[commaPos] = '.';
        commaPos = processed.find(',', commaPos + 1);
    }
    return processed;
}

// CSV Reader with ISO-8859-1 encoding support
std::vector<std::vector<std::string>> ReadCSVWithEncoding(const std::string &filename)
{
    std::vector<std::vector<std::string>> data;
    std::ifstream file(filename, std::ios::binary);

    if (!file.is_open())
    {
        AddLogMessage("ERROR: Cannot open file: " + filename);
        return data;
    }

    std::string line;
    while (std::getline(file, line))
    {
        std::string utf8Line = ConvertISO88591ToUTF8(line);

        std::stringstream ss(utf8Line);
        std::string cell;
        std::vector<std::string> row;

        while (std::getline(ss, cell, ';'))
        {
            cell.erase(0, cell.find_first_not_of(" \t\r\n"));
            cell.erase(cell.find_last_not_of(" \t\r\n") + 1);

            std::string processedCell = ProcessDecimalValue(cell);
            row.push_back(processedCell);
        }

        if (!row.empty())
            data.push_back(row);
    }

    file.close();
    AddLogMessage("SUCCESS: Read " + std::to_string(data.size()) + " rows from " + fs::path(filename).filename().string());
    return data;
}

// SQL Value formatting helper
std::string SqlValue(const std::string &val, bool isText, const std::string &defaultVal)
{
    if (val.empty())
        return defaultVal;
    if (isText)
    {
        std::string escaped = val;
        size_t pos = 0;
        while ((pos = escaped.find("'", pos)) != std::string::npos)
        {
            escaped.replace(pos, 1, "''");
            pos += 2;
        }
        return "'" + escaped + "'";
    }
    return val;
}

// Update progress bar
void UpdateProgress(int percentage)
{
    SendMessage(g_hProgressBar, PBM_SETPOS, percentage, 0);
}

// Create database tables
bool CreateTables(sqlite3 *db)
{
    const char *createSQL = R"SQL(
CREATE TABLE IF NOT EXISTS Matlist (
    MatItemNr      TEXT(6) PRIMARY KEY,
    Name           TEXT(32) NOT NULL,
    SetPlusTol     REAL DEFAULT 0.01,
    SetMinusTol    REAL DEFAULT 0.01,
    ThermcapIx     INTEGER DEFAULT -1,
    TA             INTEGER DEFAULT -1,
    TATyp          INTEGER DEFAULT -1,
    PriceKG        REAL DEFAULT 0.0,
    StSizeMin      REAL DEFAULT 0.0,
    StSizeAlarm    REAL DEFAULT 0.0,
    ReplaceMatNr   TEXT(6) DEFAULT '',
    CompTyp        INTEGER DEFAULT 0,
    Variante       INTEGER DEFAULT 1,
    MatTyp         INTEGER DEFAULT 0,
    InformUser     INTEGER DEFAULT 0,
    Decremt        INTEGER DEFAULT 1,
    Allergene      INTEGER DEFAULT 0,
    Barcode        TEXT(30) DEFAULT '',
    Gebindem       REAL DEFAULT 0.00
);

CREATE TABLE IF NOT EXISTS RecipeHead (
    Nr TEXT PRIMARY KEY,
    Name TEXT NOT NULL,
    LongName TEXT,
    PieceWeight REAL DEFAULT 1.0,
    RcpWeight REAL,
    RunTime TEXT DEFAULT '00:00:00',
    NameVariantB TEXT,
    NameVariantC TEXT,
    NameVariantD TEXT,
    NameVariantE TEXT,
    NameVariantF TEXT,
    NameVariantG TEXT,
    NameVariantH TEXT,
    NameVariantI TEXT,
    NameVariantJ TEXT,
    NameVariantK TEXT,
    TA REAL DEFAULT 0.0,
    TA_B REAL DEFAULT 0.0,
    TA_C REAL DEFAULT 0.0,
    TA_D REAL DEFAULT 0.0,
    TA_E REAL DEFAULT 0.0,
    TA_F REAL DEFAULT 0.0,
    TA_G REAL DEFAULT 0.0,
    TA_H REAL DEFAULT 0.0,
    TA_I REAL DEFAULT 0.0,
    TA_J REAL DEFAULT 0.0,
    TA_K REAL DEFAULT 0.0,
    PasteStill REAL DEFAULT 0.0,
    PasteStill_B REAL DEFAULT 0.0,
    PasteStill_C REAL DEFAULT 0.0,
    PasteStill_D REAL DEFAULT 0.0,
    PasteStill_E REAL DEFAULT 0.0,
    PasteStill_F REAL DEFAULT 0.0,
    PasteStill_G REAL DEFAULT 0.0,
    PasteStill_H REAL DEFAULT 0.0,
    PasteStill_I REAL DEFAULT 0.0,
    PasteStill_J REAL DEFAULT 0.0,
    PasteStill_K REAL DEFAULT 0.0,
    PasteTemp REAL DEFAULT 0.0,
    PasteTemp_B REAL DEFAULT 0.0,
    PasteTemp_C REAL DEFAULT 0.0,
    PasteTemp_D REAL DEFAULT 0.0,
    PasteTemp_E REAL DEFAULT 0.0,
    PasteTemp_F REAL DEFAULT 0.0,
    PasteTemp_G REAL DEFAULT 0.0,
    PasteTemp_H REAL DEFAULT 0.0,
    PasteTemp_I REAL DEFAULT 0.0,
    PasteTemp_J REAL DEFAULT 0.0,
    PasteTemp_K REAL DEFAULT 0.0,
    WaterKorr INTEGER DEFAULT 0,
    MixerGroup INTEGER DEFAULT 0,
    KnetRecipe INTEGER DEFAULT 0,
    TargetLine INTEGER DEFAULT 1,
    LineWeight REAL DEFAULT 0.0,
    Article1 TEXT,
    Article2 TEXT,
    Article3 TEXT,
    ShowBacktip INTEGER DEFAULT 0,
    Public INTEGER DEFAULT 1,
    RecipeGroup INTEGER DEFAULT 0,
    MinBatchWeight REAL DEFAULT 0.0,
    OptiBatchWeight REAL DEFAULT 0.0,
    MaxBatchWeight REAL DEFAULT 0.0
);

CREATE TABLE IF NOT EXISTS RecipeLine (
    RcpNr          TEXT NOT NULL,         
    RcpLine        INTEGER NOT NULL,      
    Variante       INTEGER DEFAULT 1,     
    MatItemNr      TEXT NOT NULL,         
    Dostyp         INTEGER DEFAULT 0,     
    ScaleNr        INTEGER DEFAULT 0,     
    SetWeight      REAL NOT NULL,         
    SetPlusTol     REAL DEFAULT 0.01,     
    SetMinusTol    REAL DEFAULT 0.01,     
    Discharge      INTEGER DEFAULT 0,     
    WaterTemp      REAL DEFAULT 0.0,      
    Mixtime        REAL DEFAULT 0.0,      
    Mixtyp         INTEGER DEFAULT 0,     
    KnetInc1       REAL DEFAULT 0.0,      
    KnetInc2       REAL DEFAULT 0.0,      
    KnetInc3       REAL DEFAULT 0.0,      
    KompInc1       REAL DEFAULT 0.0,      
    KompInc2       REAL DEFAULT 0.0,      
    KompInc3       REAL DEFAULT 0.0,      
    ReplaceMatNr   TEXT,                  
    CompTyp        INTEGER DEFAULT 0,     
    CompVariante   INTEGER DEFAULT 1,     
    EnergyEntry    INTEGER DEFAULT 0,     
    RecipeLineId   TEXT NOT NULL,         
    TransferToSPS  INTEGER DEFAULT 1,     
    KneadToolParam INTEGER,               
    KneadBowlParam INTEGER,               
    Gebinde        INTEGER DEFAULT 0,     
    RecipeTip      TEXT,
    FOREIGN KEY (RcpNr) REFERENCES RecipeHead(Nr),
    FOREIGN KEY (MatItemNr) REFERENCES Matlist(MatItemNr),
    PRIMARY KEY (RcpNr, RcpLine)
);
)SQL";

    char *errMsg = nullptr;
    int rc = sqlite3_exec(db, createSQL, nullptr, nullptr, &errMsg);

    if (rc != SQLITE_OK)
    {
        AddLogMessage("ERROR: Failed to create tables: " + std::string(errMsg ? errMsg : "Unknown"));
        if (errMsg)
            sqlite3_free(errMsg);
        return false;
    }

    AddLogMessage("SUCCESS: Database tables created");
    return true;
}

// Insert functions
bool InsertMatlist(sqlite3 *db, const std::vector<std::vector<std::string>> &data)
{
    std::vector<std::string> defaults = {
        "''", "''", "0.01", "0.01", "-1", "-1", "-1", "0.0", "0.0", "0.0",
        "''", "0", "1", "0", "0", "1", "0", "''", "0.00"};

    size_t colCount = defaults.size();
    char *errMsg = nullptr;
    sqlite3_exec(db, "BEGIN TRANSACTION", nullptr, nullptr, nullptr);

    for (size_t rowNum = 0; rowNum < data.size(); rowNum++)
    {
        UpdateProgress(10 + (int)((30 * rowNum) / data.size()));

        auto rowRaw = data[rowNum];
        std::vector<std::string> row = rowRaw;

        if (row.size() < colCount)
            row.resize(colCount, "");
        else if (row.size() > colCount)
            row.resize(colCount);

        std::string sql = "INSERT OR REPLACE INTO Matlist VALUES (";
        for (size_t i = 0; i < colCount; i++)
        {
            bool isText = (i == 0 || i == 1 || i == 10 || i == 17);
            std::string value = SqlValue(row[i], isText, defaults[i]);
            sql += value;
            if (i < colCount - 1)
                sql += ",";
        }
        sql += ");";

        int rc = sqlite3_exec(db, sql.c_str(), nullptr, nullptr, &errMsg);

        if (rc != SQLITE_OK)
        {
            sqlite3_exec(db, "ROLLBACK", nullptr, nullptr, nullptr);
            AddLogMessage("ERROR: Failed to insert Matlist row " + std::to_string(rowNum));
            if (errMsg)
                sqlite3_free(errMsg);
            return false;
        }
    }

    sqlite3_exec(db, "COMMIT", nullptr, nullptr, nullptr);
    AddLogMessage("SUCCESS: Imported " + std::to_string(data.size()) + " materials");
    return true;
}

bool InsertRecipeHead(sqlite3 *db, const std::vector<std::vector<std::string>> &data)
{
    std::vector<std::string> defaults = {
        "''", "''", "''", "1.0", "0.0", "'00:00:00'", "''", "''", "''", "''", "''", "''", "''", "''", "''", "''",
        "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0",
        "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0",
        "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0",
        "0", "0", "0", "1", "0.0", "''", "''", "''", "0", "1", "0", "0.0", "0.0", "0.0"};

    size_t colCount = defaults.size();
    char *errMsg = nullptr;
    sqlite3_exec(db, "BEGIN TRANSACTION", nullptr, nullptr, nullptr);

    for (size_t rowNum = 0; rowNum < data.size(); rowNum++)
    {
        UpdateProgress(40 + (int)((30 * rowNum) / data.size()));

        auto rowRaw = data[rowNum];
        std::vector<std::string> row = rowRaw;

        if (row.size() < colCount)
            row.resize(colCount, "");
        else if (row.size() > colCount)
            row.resize(colCount);

        std::string sql = "INSERT OR REPLACE INTO RecipeHead VALUES (";
        for (size_t i = 0; i < colCount; i++)
        {
            bool isText = (i == 0 || i == 1 || i == 2 || i == 5 ||
                           (i >= 6 && i <= 15) || (i >= 55 && i <= 57));
            std::string value = SqlValue(row[i], isText, defaults[i]);
            sql += value;
            if (i < colCount - 1)
                sql += ",";
        }
        sql += ");";

        int rc = sqlite3_exec(db, sql.c_str(), nullptr, nullptr, &errMsg);

        if (rc != SQLITE_OK)
        {
            sqlite3_exec(db, "ROLLBACK", nullptr, nullptr, nullptr);
            AddLogMessage("ERROR: Failed to insert RecipeHead row " + std::to_string(rowNum));
            if (errMsg)
                sqlite3_free(errMsg);
            return false;
        }
    }

    sqlite3_exec(db, "COMMIT", nullptr, nullptr, nullptr);
    AddLogMessage("SUCCESS: Imported " + std::to_string(data.size()) + " recipes");
    return true;
}

bool InsertRecipeLine(sqlite3 *db, const std::vector<std::vector<std::string>> &data)
{
    std::vector<std::string> defaults = {
        "''", "0", "1", "''", "0", "0", "0.0", "0.01", "0.01", "0",
        "0.0", "0.0", "0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "''",
        "0", "1", "0", "''", "1", "0", "0", "0", "''"};

    size_t colCount = defaults.size();
    char *errMsg = nullptr;
    sqlite3_exec(db, "BEGIN TRANSACTION", nullptr, nullptr, nullptr);

    for (size_t rowNum = 0; rowNum < data.size(); rowNum++)
    {
        UpdateProgress(70 + (int)((25 * rowNum) / data.size()));

        auto rowRaw = data[rowNum];
        std::vector<std::string> row = rowRaw;

        if (row.size() < colCount)
            row.resize(colCount, "");
        else if (row.size() > colCount)
            row.resize(colCount);

        std::string sql = "INSERT OR REPLACE INTO RecipeLine VALUES (";
        for (size_t i = 0; i < colCount; i++)
        {
            bool isText = (i == 0 || i == 3 || i == 19 || i == 23 || i == 28);
            std::string value = SqlValue(row[i], isText, defaults[i]);
            sql += value;
            if (i < colCount - 1)
                sql += ",";
        }
        sql += ");";

        int rc = sqlite3_exec(db, sql.c_str(), nullptr, nullptr, &errMsg);

        if (rc != SQLITE_OK)
        {
            sqlite3_exec(db, "ROLLBACK", nullptr, nullptr, nullptr);
            AddLogMessage("ERROR: Failed to insert RecipeLine row " + std::to_string(rowNum));
            if (errMsg)
                sqlite3_free(errMsg);
            return false;
        }
    }

    sqlite3_exec(db, "COMMIT", nullptr, nullptr, nullptr);
    AddLogMessage("SUCCESS: Imported " + std::to_string(data.size()) + " recipe lines");
    return true;
}

// Import data function (runs in separate thread)
void ImportDataThread()
{
    g_importInProgress = true;
    EnableWindow(g_hImportBtn, FALSE);
    UpdateProgress(0);

    char csvPath[MAX_PATH];
    char dbPath[MAX_PATH];

    GetWindowTextA(g_hCsvPathEdit, csvPath, MAX_PATH);
    GetWindowTextA(g_hDbPathEdit, dbPath, MAX_PATH);

    AddLogMessage("Starting import process...");

    if (strlen(csvPath) == 0)
    {
        AddLogMessage("ERROR: Please select CSV folder");
        EnableWindow(g_hImportBtn, TRUE);
        g_importInProgress = false;
        return;
    }

    if (strlen(dbPath) == 0)
    {
        strcpy_s(dbPath, "bakery.db");
    }

    // Check CSV files exist
    std::string matlistPath = std::string(csvPath) + "\\Matlist.csv";
    std::string recipeHeadPath = std::string(csvPath) + "\\Recipehead.csv";
    std::string recipeLinePath = std::string(csvPath) + "\\Recipeline.csv";

    if (!fs::exists(matlistPath) || !fs::exists(recipeHeadPath) || !fs::exists(recipeLinePath))
    {
        AddLogMessage("ERROR: Required CSV files not found in folder");
        EnableWindow(g_hImportBtn, TRUE);
        g_importInProgress = false;
        return;
    }

    AddLogMessage("All CSV files found");
    UpdateProgress(5);

    // Open database
    sqlite3 *db;
    if (sqlite3_open(dbPath, &db))
    {
        AddLogMessage("ERROR: Cannot open database: " + std::string(sqlite3_errmsg(db)));
        EnableWindow(g_hImportBtn, TRUE);
        g_importInProgress = false;
        return;
    }

    AddLogMessage("Database opened: " + std::string(dbPath));
    UpdateProgress(10);

    // Create tables
    if (!CreateTables(db))
    {
        sqlite3_close(db);
        EnableWindow(g_hImportBtn, TRUE);
        g_importInProgress = false;
        return;
    }

    try
    {
        // Import Matlist
        AddLogMessage("Reading Matlist.csv...");
        auto matlistData = ReadCSVWithEncoding(matlistPath);
        if (!matlistData.empty() && !InsertMatlist(db, matlistData))
        {
            sqlite3_close(db);
            EnableWindow(g_hImportBtn, TRUE);
            g_importInProgress = false;
            return;
        }

        // Import RecipeHead
        AddLogMessage("Reading Recipehead.csv...");
        auto recipeHeadData = ReadCSVWithEncoding(recipeHeadPath);
        if (!recipeHeadData.empty() && !InsertRecipeHead(db, recipeHeadData))
        {
            sqlite3_close(db);
            EnableWindow(g_hImportBtn, TRUE);
            g_importInProgress = false;
            return;
        }

        // Import RecipeLine
        AddLogMessage("Reading Recipeline.csv...");
        auto recipeLineData = ReadCSVWithEncoding(recipeLinePath);
        if (!recipeLineData.empty() && !InsertRecipeLine(db, recipeLineData))
        {
            sqlite3_close(db);
            EnableWindow(g_hImportBtn, TRUE);
            g_importInProgress = false;
            return;
        }

        UpdateProgress(100);
        AddLogMessage("SUCCESS: Import completed successfully!");
        AddLogMessage("Database saved to: " + std::string(dbPath));

        MessageBoxA(g_hMainWindow, "Import completed successfully!", "Success", MB_OK | MB_ICONINFORMATION);
    }
    catch (...)
    {
        AddLogMessage("ERROR: Exception during import");
    }

    sqlite3_close(db);
    EnableWindow(g_hImportBtn, TRUE);
    g_importInProgress = false;
}

// Browse for folder
std::string BrowseForFolder(HWND parent)
{
    char path[MAX_PATH] = "";

    BROWSEINFOA bi = {};
    bi.hwndOwner = parent;
    bi.lpszTitle = "Select CSV Files Folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    LPITEMIDLIST pidl = SHBrowseForFolderA(&bi);
    if (pidl)
    {
        SHGetPathFromIDListA(pidl, path);
        CoTaskMemFree(pidl);
    }

    return std::string(path);
}

// Browse for database file
std::string BrowseForDatabase(HWND parent)
{
    char filename[MAX_PATH] = "bakery.db";

    OPENFILENAMEA ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = parent;
    ofn.lpstrFile = filename;
    ofn.nMaxFile = sizeof(filename);
    ofn.lpstrFilter = "SQLite Database\0*.db\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrTitle = "Save Database As";
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
    ofn.lpstrDefExt = "db";

    if (GetSaveFileNameA(&ofn))
    {
        return std::string(filename);
    }

    return "";
}

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_CREATE:
        // Initialize common controls
        InitCommonControls();

        // Create controls
        CreateWindowA("STATIC", "Bakery CSV Import Tool",
                      WS_VISIBLE | WS_CHILD | SS_CENTER,
                      20, 10, 760, 30, hwnd, nullptr, GetModuleHandle(nullptr), nullptr);

        CreateWindowA("STATIC", "CSV Files Folder:",
                      WS_VISIBLE | WS_CHILD,
                      20, 50, 120, 20, hwnd, nullptr, GetModuleHandle(nullptr), nullptr);

        g_hCsvPathEdit = CreateWindowA("EDIT", "",
                                       WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
                                       20, 70, 600, 25, hwnd, (HMENU)ID_CSV_PATH_EDIT, GetModuleHandle(nullptr), nullptr);

        CreateWindowA("BUTTON", "Browse...",
                      WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                      640, 70, 80, 25, hwnd, (HMENU)ID_CSV_BROWSE_BTN, GetModuleHandle(nullptr), nullptr);

        CreateWindowA("STATIC", "Database Path:",
                      WS_VISIBLE | WS_CHILD,
                      20, 110, 120, 20, hwnd, nullptr, GetModuleHandle(nullptr), nullptr);

        g_hDbPathEdit = CreateWindowA("EDIT", "bakery.db",
                                      WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
                                      20, 130, 600, 25, hwnd, (HMENU)ID_DB_PATH_EDIT, GetModuleHandle(nullptr), nullptr);

        CreateWindowA("BUTTON", "Browse...",
                      WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                      640, 130, 80, 25, hwnd, (HMENU)ID_DB_BROWSE_BTN, GetModuleHandle(nullptr), nullptr);

        g_hImportBtn = CreateWindowA("BUTTON", "Start Import",
                                     WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                                     20, 170, 120, 35, hwnd, (HMENU)ID_IMPORT_BTN, GetModuleHandle(nullptr), nullptr);

        CreateWindowA("BUTTON", "Exit",
                      WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                      160, 170, 80, 35, hwnd, (HMENU)ID_EXIT_BTN, GetModuleHandle(nullptr), nullptr);

        CreateWindowA("STATIC", "Progress:",
                      WS_VISIBLE | WS_CHILD,
                      20, 220, 60, 20, hwnd, nullptr, GetModuleHandle(nullptr), nullptr);

        g_hProgressBar = CreateWindowA(PROGRESS_CLASS, nullptr,
                                       WS_VISIBLE | WS_CHILD,
                                       90, 220, 630, 20, hwnd, (HMENU)ID_PROGRESS_BAR, GetModuleHandle(nullptr), nullptr);
        SendMessage(g_hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

        CreateWindowA("STATIC", "Log:",
                      WS_VISIBLE | WS_CHILD,
                      20, 250, 40, 20, hwnd, nullptr, GetModuleHandle(nullptr), nullptr);

        g_hLogEdit = CreateWindowA("EDIT", "",
                                   WS_VISIBLE | WS_CHILD | WS_BORDER | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                                   20, 270, 740, 200, hwnd, (HMENU)ID_LOG_EDIT, GetModuleHandle(nullptr), nullptr);

        AddLogMessage("Bakery CSV Import Tool started");
        AddLogMessage("Please select CSV folder and database path");
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_CSV_BROWSE_BTN:
        {
            std::string folder = BrowseForFolder(hwnd);
            if (!folder.empty())
            {
                SetWindowTextA(g_hCsvPathEdit, folder.c_str());
                AddLogMessage("CSV folder selected: " + folder);
            }
            break;
        }
        case ID_DB_BROWSE_BTN:
        {
            std::string dbFile = BrowseForDatabase(hwnd);
            if (!dbFile.empty())
            {
                SetWindowTextA(g_hDbPathEdit, dbFile.c_str());
                AddLogMessage("Database path selected: " + dbFile);
            }
            break;
        }
        case ID_IMPORT_BTN:
            if (!g_importInProgress)
            {
                std::thread importThread(ImportDataThread);
                importThread.detach();
            }
            break;
        case ID_EXIT_BTN:
            PostQuitMessage(0);
            break;
        }
        break;

    case WM_CLOSE:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    // Initialize COM for shell functions
    CoInitialize(nullptr);

    // Register window class
    WNDCLASSA wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = "BakeryImportTool";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);

    if (!RegisterClassA(&wc))
    {
        MessageBoxA(nullptr, "Failed to register window class", "Error", MB_OK | MB_ICONERROR);
        CoUninitialize();
        return -1;
    }

    // Create main window
    g_hMainWindow = CreateWindowA(
        "BakeryImportTool",
        "Bakery CSV Import Tool",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 520,
        nullptr, nullptr, hInstance, nullptr);

    if (!g_hMainWindow)
    {
        MessageBoxA(nullptr, "Failed to create window", "Error", MB_OK | MB_ICONERROR);
        CoUninitialize();
        return -1;
    }

    ShowWindow(g_hMainWindow, nCmdShow);
    UpdateWindow(g_hMainWindow);

    // Message loop
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CoUninitialize();
    return (int)msg.wParam;
}