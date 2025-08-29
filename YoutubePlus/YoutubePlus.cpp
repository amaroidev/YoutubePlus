// YoutubePlus.cpp : Defines the entry point for the application.
//

#include "pch.h"
#include "framework.h"
#include "YoutubePlus.h"

#include <objbase.h>
#include <string>
#include <wrl.h>
#include <WebView2.h>
#include <CommCtrl.h>
#include <thread>
#include <vector>
#include <regex>
#include <ShlObj.h> // For SHBrowseForFolder
#include <fstream>
#include <mutex>
#include <sstream>
#include "nlohmann/json.hpp"
#include <Windows.h>
#include <winreg.h> // For registry functions
#include <shellapi.h> // For ShellExecute and SHELLEXECUTEINFO

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "WebView2LoaderStatic.lib")
#pragma comment(lib, "Shell32.lib") // For ShellExecute functions

using namespace Microsoft::WRL;

// Forward declarations
INT_PTR CALLBACK DownloadOptionsProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK DownloadProgressProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK DownloadManagerProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
DWORD WINAPI DownloadThread(LPVOID lpParam);
void SetLightMode();
void SetDarkMode();
void UseSystemTheme();
void ApplyThemeMode();
bool IsWebView2RuntimeInstalled();
bool InstallWebView2Runtime(HWND hWnd);

// Global Variables:
HINSTANCE hInst;                                // current instance
ICoreWebView2Controller* g_webViewController = nullptr;
ICoreWebView2* g_webView = nullptr;
bool g_isAudioOnly = false;

enum DownloadStatus {
    Downloading,
    Completed,
    Failed,
    Cancelled
};

// Struct to hold all info about a download
struct DownloadItem {
    std::wstring url;
    std::wstring resolution;
    std::wstring path;
    bool downloadSubtitles;
    DownloadStatus status;
    double progress;
    HANDLE hProcess;
    HANDLE hThread;  // Add handle to thread
    DWORD threadId;
    HWND progressDlg; // Handle to the progress dialog for this download
    DWORD startTime;  // Add start time for calculating progress
};

std::vector<DownloadItem*> g_downloadQueue;
std::mutex g_queueMutex;

// Application settings
struct AppSettings {
    bool adBlockOnStartup = true;
    std::wstring defaultDownloadPath = L"";
    
    // Theme settings
    enum ThemeMode {
        Light,
        Dark,
        System
    };
    ThemeMode themeMode = ThemeMode::System;
};
AppSettings g_settings;

// Helper function to get the path to the settings file
std::wstring GetSettingsPath() {
    PWSTR path = NULL;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &path))) {
        std::wstring settingsPath = path;
        CoTaskMemFree(path);
        settingsPath += L"\\YoutubePlus";
        CreateDirectoryW(settingsPath.c_str(), NULL);
        settingsPath += L"\\settings.json";
        return settingsPath;
    }
    return L"";
}

// Helper function to save settings
void SaveSettings() {
    nlohmann::json j;
    j["adBlockOnStartup"] = g_settings.adBlockOnStartup;
    j["themeMode"] = static_cast<int>(g_settings.themeMode);
   
    // Convert wstring to string for JSON serialization
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &g_settings.defaultDownloadPath[0], (int)g_settings.defaultDownloadPath.size(), NULL, 0, NULL, NULL);
    std::string path_str(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &g_settings.defaultDownloadPath[0], (int)g_settings.defaultDownloadPath.size(), &path_str[0], size_needed, NULL, NULL);
    j["defaultDownloadPath"] = path_str;

    std::wstring settingsPath = GetSettingsPath();
    if (!settingsPath.empty()) {
        std::ofstream o(settingsPath);
        o << j.dump(4) << std::endl;
    }
}

// Helper function to load settings
void LoadSettings() {
    std::wstring settingsPath = GetSettingsPath();
    if (!settingsPath.empty()) {
        std::ifstream i(settingsPath);
        if (i.good()) {
            nlohmann::json j;
            i >> j;
            if (j.contains("adBlockOnStartup")) {
                g_settings.adBlockOnStartup = j["adBlockOnStartup"];
            }
            if (j.contains("themeMode")) {
                g_settings.themeMode = static_cast<AppSettings::ThemeMode>(j["themeMode"].get<int>());
            }
            if (j.contains("defaultDownloadPath")) {
                std::string path_str = j["defaultDownloadPath"];
                // Convert string to wstring
                int size_needed = MultiByteToWideChar(CP_UTF8, 0, &path_str[0], (int)path_str.size(), NULL, 0);
                std::wstring w_path_str(size_needed, 0);
                MultiByteToWideChar(CP_UTF8, 0, &path_str[0], (int)path_str.size(), &w_path_str[0], size_needed);
                g_settings.defaultDownloadPath = w_path_str;
            }
        }
    }
}

// Struct to hold download progress information
struct DownloadProgressInfo {
    HWND hDlg;
    std::wstring videoUrl;
};

// Struct to hold download options
struct DownloadOptions {
    HWND hDlg; // Add hDlg to be passed to the thread
    std::wstring resolution;
    std::wstring path;
    std::wstring videoUrl;
    bool downloadSubtitles;
    HANDLE hProcess = NULL; // To hold the handle of the yt-dlp process
};

// Structure to hold playlist video info
struct PlaylistVideo {
    std::wstring title;
    std::wstring url;
    bool selected;
};

std::vector<PlaylistVideo> g_playlistVideos;

// Parse playlist page to extract videos
DWORD WINAPI FetchPlaylistVideosThread(LPVOID lpParam) {
    HWND hDlg = (HWND)lpParam;
    
    // Get the playlist URL safely
    std::wstring playlistUrl;
    
    // Get the playlist URL with safer access
    void* userData = (void*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
    if (!userData) {
        MessageBox(hDlg, L"Failed to get playlist URL from dialog.", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    try {
        playlistUrl = *(std::wstring*)userData;
    }
    catch (...) {
        MessageBox(hDlg, L"Error accessing playlist URL data.", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    if (playlistUrl.empty()) {
        MessageBox(hDlg, L"Playlist URL is empty.", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Support both playlist URLs and watch?v=XXX&list=XXX format
    bool isPlaylistPage = playlistUrl.find(L"youtube.com/playlist") != std::wstring::npos;
    bool isWatchListPage = playlistUrl.find(L"youtube.com/watch") != std::wstring::npos && 
                          playlistUrl.find(L"list=") != std::wstring::npos;
    
    if (!isPlaylistPage && !isWatchListPage) {
        MessageBox(hDlg, L"The URL does not appear to be a YouTube playlist.", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Get full path to yt-dlp.exe
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    std::wstring exeDir = exePath;
    size_t pos = exeDir.find_last_of(L"\\/");
    std::wstring ytdlpPath = L"yt-dlp.exe";
    std::wstring workingDir;
    if (pos != std::wstring::npos) {
        workingDir = exeDir.substr(0, pos);
        ytdlpPath = L"\"" + workingDir + L"\\yt-dlp.exe\"";
    }

    // Run yt-dlp to get playlist info in JSON format
    std::wstring command = ytdlpPath + L" --flat-playlist --dump-json \"" + playlistUrl + L"\"";
    
    HANDLE hChildStd_OUT_Rd = NULL;
    HANDLE hChildStd_OUT_Wr = NULL;
    HANDLE hChildStd_ERR_Rd = NULL;
    HANDLE hChildStd_ERR_Wr = NULL;
    
    SECURITY_ATTRIBUTES sa;
    sa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    // Create pipes for stdout and stderr with proper error checking
    if (!CreatePipe(&hChildStd_OUT_Rd, &hChildStd_OUT_Wr, &sa, 0) ||
        !CreatePipe(&hChildStd_ERR_Rd, &hChildStd_ERR_Wr, &sa, 0)) {
        MessageBox(hDlg, L"Failed to create pipes for process communication.", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    SetHandleInformation(hChildStd_OUT_Rd, HANDLE_FLAG_INHERIT, 0);
    SetHandleInformation(hChildStd_ERR_Rd, HANDLE_FLAG_INHERIT, 0);

    PROCESS_INFORMATION pi = { 0 };
    STARTUPINFOW si = { sizeof(si) };
    si.hStdError = hChildStd_ERR_Wr;
    si.hStdOutput = hChildStd_OUT_Wr;
    si.dwFlags |= STARTF_USESTDHANDLES;

    g_playlistVideos.clear();
    
    bool processCreated = CreateProcessW(nullptr, &command[0], nullptr, nullptr, TRUE, CREATE_NO_WINDOW, nullptr, workingDir.c_str(), &si, &pi);
    
    // Always close write handles after process creation
    CloseHandle(hChildStd_OUT_Wr);
    CloseHandle(hChildStd_ERR_Wr);
    
    if (!processCreated) {
        // Get the error message
        DWORD errorCode = GetLastError();
        wchar_t errorMsg[256];
        swprintf_s(errorMsg, L"Failed to start yt-dlp.exe. Error code: %d", errorCode);
        MessageBox(hDlg, errorMsg, L"Error", MB_OK | MB_ICONERROR);
        CloseHandle(hChildStd_OUT_Rd);
        CloseHandle(hChildStd_ERR_Rd);
        return 1;
    }
    
    // Read the output
    std::string jsonOutput;
    std::string errorOutput;
    char buffer[4096];
    DWORD bytesRead;
    
    // Read standard output
    while (ReadFile(hChildStd_OUT_Rd, buffer, sizeof(buffer) - 1, &bytesRead, NULL) && bytesRead > 0) {
        buffer[bytesRead] = '\0';
        jsonOutput += buffer;
    }
    
    // Read error output
    while (ReadFile(hChildStd_ERR_Rd, buffer, sizeof(buffer) - 1, &bytesRead, NULL) && bytesRead > 0) {
        buffer[bytesRead] = '\0';
        errorOutput += buffer;
    }
    
    // Wait for process to finish with timeout
    WaitForSingleObject(pi.hProcess, 30000); // 30 second timeout
    
    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    
    // Clean up process handles
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    CloseHandle(hChildStd_OUT_Rd);
    CloseHandle(hChildStd_ERR_Rd);

    if (exitCode != 0) {
        // Process failed
        wchar_t errorMsg[4096];
        if (!errorOutput.empty()) {
            // Convert error output to wstring for display
            int size_needed = MultiByteToWideChar(CP_UTF8, 0, errorOutput.c_str(), (int)errorOutput.size(), NULL, 0);
            std::wstring wErrorOutput(size_needed, 0);
            MultiByteToWideChar(CP_UTF8, 0, errorOutput.c_str(), (int)errorOutput.size(), &wErrorOutput[0], size_needed);
            
            swprintf_s(errorMsg, L"Failed to extract playlist data. yt-dlp.exe returned error code %d.\n\nError: %s", 
                       exitCode, wErrorOutput.c_str());
        } else {
            swprintf_s(errorMsg, L"Failed to extract playlist data. yt-dlp.exe returned error code %d.", exitCode);
        }
        
        MessageBox(hDlg, errorMsg, L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    if (jsonOutput.empty()) {
        MessageBox(hDlg, L"No playlist data was returned from yt-dlp.", L"Warning", MB_OK | MB_ICONWARNING);
        return 1;
    }

    // Parse each line as a separate JSON object
    std::istringstream stream(jsonOutput);
    std::string line;
    int videosAdded = 0;
    
    while (std::getline(stream, line)) {
        if (line.empty()) continue;
        
        try {
            auto json = nlohmann::json::parse(line);
            
            // Use "id" instead of "url" for better compatibility
            if (json.contains("title") && json.contains("id")) {
                std::string title;
                std::string id;
                
                try {
                    title = json["title"].get<std::string>();
                    id = json["id"].get<std::string>();
                }
                catch (...) {
                    continue; // Skip entries with invalid title or id
                }
                
                // Convert to wstring
                int titleSize = MultiByteToWideChar(CP_UTF8, 0, title.c_str(), -1, NULL, 0);
                if (titleSize <= 0) continue;
                
                std::vector<wchar_t> wTitle(titleSize);
                MultiByteToWideChar(CP_UTF8, 0, title.c_str(), -1, &wTitle[0], titleSize);
                
                // Construct the YouTube video URL directly from video ID
                std::wstring videoUrl = L"https://www.youtube.com/watch?v=" + std::wstring(id.begin(), id.end());
                
                // Add to our playlist videos list
                PlaylistVideo video = {&wTitle[0], videoUrl, true};
                g_playlistVideos.push_back(video);
                videosAdded++;
            }
        }
        catch (...) {
            // Skip any lines that don't parse correctly
            continue;
        }
    }
    
    if (videosAdded == 0) {
        MessageBox(hDlg, L"No videos were found in the playlist.", L"Warning", MB_OK | MB_ICONWARNING);
        return 1;
    }
    
    // Update the UI with the fetched videos
    if (IsWindow(hDlg)) {
        // Use PostMessage instead of SendMessage to avoid deadlocks
        PostMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_PLAYLIST_LIST, LBN_SELCHANGE), 0);
    }
    
    return 0;
}

// Dialog procedure for playlist selection
INT_PTR CALLBACK PlaylistDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static std::wstring* pPlaylistUrl = nullptr;
    
    switch (message) {
    case WM_INITDIALOG: {
        pPlaylistUrl = (std::wstring*)lParam;
        SetWindowLongPtr(hDlg, GWLP_USERDATA, (LONG_PTR)pPlaylistUrl);
        
        // Initialize the list view
        HWND hList = GetDlgItem(hDlg, IDC_PLAYLIST_LIST);
        
        // Add columns to the list view
        LVCOLUMNW lvc = { 0 };
        lvc.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        
        lvc.iSubItem = 0;
        lvc.pszText = (LPWSTR)L"Title";
        lvc.cx = 320;
        ListView_InsertColumn(hList, 0, &lvc);
        
        lvc.iSubItem = 1;
        lvc.pszText = (LPWSTR)L"URL";
        lvc.cx = 60;
        ListView_InsertColumn(hList, 1, &lvc);
        
        // Set extended style to enable checkboxes
        ListView_SetExtendedListViewStyle(hList, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);
        
        // Start thread to fetch playlist data
        CreateThread(NULL, 0, FetchPlaylistVideosThread, (LPVOID)hDlg, 0, NULL);
        
        return (INT_PTR)TRUE;
    }
    
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            // Process selected videos
            std::vector<std::wstring> selectedUrls;
            
            for (const auto& video : g_playlistVideos) {
                if (video.selected) {
                    selectedUrls.push_back(video.url);
                }
            }
            
            if (!selectedUrls.empty()) {
                // Ask for download options first
                DownloadOptions options = { nullptr, L"Best", g_settings.defaultDownloadPath, L"", false };
                if (DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_DOWNLOAD_OPTIONS), hDlg, DownloadOptionsProc, (LPARAM)&options) == IDOK) {
                    // Pass the selected URLs and options to the Download Manager
                    auto managerData = new std::pair<std::vector<std::wstring>, DownloadOptions>(selectedUrls, options);
                    DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_DOWNLOAD_MANAGER), NULL, DownloadManagerProc, (LPARAM)managerData);
                }
            }
            
            EndDialog(hDlg, IDOK);
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDC_BUTTON_SELECT_ALL) {
            HWND hList = GetDlgItem(hDlg, IDC_PLAYLIST_LIST);
            int count = ListView_GetItemCount(hList);
            
            for (int i = 0; i < count; i++) {
                ListView_SetCheckState(hList, i, TRUE);
                if (i < g_playlistVideos.size()) {
                    g_playlistVideos[i].selected = true;
                }
            }
            
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDC_BUTTON_DESELECT_ALL) {
            HWND hList = GetDlgItem(hDlg, IDC_PLAYLIST_LIST);
            int count = ListView_GetItemCount(hList);
            
            for (int i = 0; i < count; i++) {
                ListView_SetCheckState(hList, i, FALSE);
                if (i < g_playlistVideos.size()) {
                    g_playlistVideos[i].selected = false;
                }
            }
            
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDC_PLAYLIST_LIST && HIWORD(wParam) == LBN_SELCHANGE) {
            // Update the list view with the fetched videos
            HWND hList = GetDlgItem(hDlg, IDC_PLAYLIST_LIST);
            ListView_DeleteAllItems(hList);
            
            for (size_t i = 0; i < g_playlistVideos.size(); i++) {
                const auto& video = g_playlistVideos[i];
                
                LVITEMW lvi = { 0 };
                lvi.mask = LVIF_TEXT;
                lvi.iItem = i;
                lvi.iSubItem = 0;
                lvi.pszText = (LPWSTR)video.title.c_str();
                int index = ListView_InsertItem(hList, &lvi);
                
                // Set URL as the second column
                ListView_SetItemText(hList, index, 1, (LPWSTR)L"View");
                
                // Set checkbox state
                ListView_SetCheckState(hList, index, video.selected);
            }
            
            return (INT_PTR)TRUE;
        }
        break;
    
    case WM_NOTIFY: {
        NMHDR* pnmhdr = (NMHDR*)lParam;
        if (pnmhdr->idFrom == IDC_PLAYLIST_LIST && pnmhdr->code == LVN_ITEMCHANGED) {
            NMLISTVIEW* pnmlv = (NMLISTVIEW*)lParam;
            if (pnmlv->uChanged & LVIF_STATE) {
                int index = pnmlv->iItem;
                if (index >= 0 && index < g_playlistVideos.size()) {
                    g_playlistVideos[index].selected = (ListView_GetCheckState(pnmhdr->hwndFrom, index) != 0);
                }
            }
        }
        break;
    }
    }
    
    return (INT_PTR)FALSE;
}

// Function to parse yt-dlp output and update progress bar
void UpdateDownloadProgress(const std::string& output, HWND hDlg, int itemIndex) {
    if (output.empty() || itemIndex < 0) {
        return; // No output or invalid index
    }
    
    try {
        // More robust regex to match yt-dlp progress outputs
        std::regex progressRegex(R"(\[download\]\s+([0-9\.]+)%\s+of\s+~?)");
        std::smatch match;
        std::string lastMatch;
        
        // Find all matches and use the last one (most recent progress)
        std::string::const_iterator searchStart(output.cbegin());
        while (std::regex_search(searchStart, output.cend(), match, progressRegex)) {
            lastMatch = match[1].str();
            searchStart = match.suffix().first;
        }
        
        if (!lastMatch.empty()) {
            try {
                float progress = std::stof(lastMatch);
                {
                    std::lock_guard<std::mutex> lock(g_queueMutex);
                    if (itemIndex < g_downloadQueue.size() && g_downloadQueue[itemIndex] != nullptr) {
                        g_downloadQueue[itemIndex]->progress = progress;
                    }
                }
            }
            catch (const std::invalid_argument&) {
                // Invalid number format in progress output
            }
            catch (const std::out_of_range&) {
                // Progress number out of range
            }
        }
    }
    catch (const std::regex_error&) {
        // Regex error - don't crash the application
    }
    catch (...) {
        // Unknown error - don't crash the application
    }
}

// Thread function to run yt-dlp and capture its output
DWORD WINAPI DownloadThread(LPVOID lpParam) {
    if (!lpParam) {
        // Invalid parameter
        return 1;
    }

    DownloadItem* item = (DownloadItem*)lpParam;
    if (!item) {
        return 1;
    }
    
    // Make sure URL is not empty
    if (item->url.empty()) {
        item->status = Failed;
        return 1;
    }
    
    // Construct the command with user-selected options
    std::wstring command;
    try {
        // Get full path to yt-dlp.exe
        wchar_t exePath[MAX_PATH];
        GetModuleFileNameW(NULL, exePath, MAX_PATH);
        std::wstring exeDir = exePath;
        size_t pos = exeDir.find_last_of(L"\\/");
        std::wstring ytdlpPath = L"yt-dlp.exe";
        if (pos != std::wstring::npos) {
            ytdlpPath = L"\"" + exeDir.substr(0, pos) + L"\\yt-dlp.exe\"";
        }

        if (item->resolution == L"Audio Only (mp3)") {
            command = ytdlpPath + L" --progress --newline --no-playlist --no-check-certificates -x --audio-format mp3";
        }
        else {
            command = ytdlpPath + L" --progress --newline --no-playlist --no-check-certificates --merge-output-format mp4";
            if (item->resolution != L"Best") {
                std::wstring res = item->resolution;
                if (!res.empty() && res.back() == 'p') {
                    res.pop_back();
                }
                command += L" -f \"bestvideo[height<=" + res + L"]+bestaudio/best\"";
            }
        }

        if (item->downloadSubtitles) {
            command += L" --write-auto-sub";
        }

        std::wstring path = item->path;
        if (path.empty()) {
            // Default to Documents folder if path is empty
            PWSTR documentsPath = nullptr;
            if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, 0, nullptr, &documentsPath))) {
                path = documentsPath;
                CoTaskMemFree(documentsPath);
            }
        }
        
        if (!path.empty() && path.back() != L'\\' && path.back() != L'/') {
            path += L'\\';
        }
        
        // Create the output directory if it doesn't exist
        CreateDirectoryW(path.c_str(), NULL);
        
        command += L" -o \"" + path + L"%(title)s.%(ext)s\"";
        command += L" \"" + item->url + L"\"";
        
        // Log the command for debugging purposes
        OutputDebugStringW((L"Running command: " + command).c_str());
    }
    catch (...) {
        // Error constructing command
        item->status = Failed;
        return 1;
    }

    HANDLE hChildStd_OUT_Rd = NULL;
    HANDLE hChildStd_OUT_Wr = NULL;
    HANDLE hChildStd_ERR_Rd = NULL;
    HANDLE hChildStd_ERR_Wr = NULL;
    
    SECURITY_ATTRIBUTES sa;
    sa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    // Create pipes with proper error handling
    bool pipesCreated = true;
    if (!CreatePipe(&hChildStd_OUT_Rd, &hChildStd_OUT_Wr, &sa, 0)) {
        pipesCreated = false;
    }
    if (!CreatePipe(&hChildStd_ERR_Rd, &hChildStd_ERR_Wr, &sa, 0)) {
        pipesCreated = false;
        if (hChildStd_OUT_Rd) CloseHandle(hChildStd_OUT_Rd);
        if (hChildStd_OUT_Wr) CloseHandle(hChildStd_OUT_Wr);
    }
    
    if (!pipesCreated) {
        item->status = Failed;
        return 1;
    }

    // Set the pipe handles as non-inheritable
    SetHandleInformation(hChildStd_OUT_Rd, HANDLE_FLAG_INHERIT, 0);
    SetHandleInformation(hChildStd_ERR_Rd, HANDLE_FLAG_INHERIT, 0);

    PROCESS_INFORMATION pi = {0};
    STARTUPINFOW si = {sizeof(si)};
    si.hStdError = hChildStd_ERR_Wr;
    si.hStdOutput = hChildStd_OUT_Wr;
    si.dwFlags |= STARTF_USESTDHANDLES;

    bool success = false;
    bool processStarted = false;
    DWORD lastError = 0;
    
    try {
        // Get the current executable directory for yt-dlp.exe
        wchar_t exePath[MAX_PATH];
        GetModuleFileNameW(NULL, exePath, MAX_PATH);
        std::wstring exeDir = std::wstring(exePath);
        size_t lastSlash = exeDir.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos) {
            exeDir = exeDir.substr(0, lastSlash);
        }
        
        // Create the process with proper working directory
        processStarted = CreateProcessW(nullptr, &command[0], nullptr, nullptr, TRUE, 
                                      CREATE_NO_WINDOW, nullptr, exeDir.c_str(), &si, &pi);
        
        if (!processStarted) {
            lastError = GetLastError();
        }
    }
    catch (...) {
        processStarted = false;
        lastError = GetLastError();
    }
    
    // Close write handles after process creation attempt
    CloseHandle(hChildStd_OUT_Wr);
    CloseHandle(hChildStd_ERR_Wr);
    
    if (processStarted) {
        item->hProcess = pi.hProcess;
        
        char buffer[256];
        DWORD dwRead;
        std::string full_output;
        std::string error_output;
        
        // Read process output with better error handling
        try {
            // Read output in a more robust way using PeekNamedPipe
            while (true) {
                // Check if download was cancelled
                if (item->status == Cancelled) {
                    break;
                }
                
                DWORD available = 0;
                DWORD exitCode = STILL_ACTIVE;
                
                // First check if process is still running
                if (!GetExitCodeProcess(pi.hProcess, &exitCode) || exitCode != STILL_ACTIVE) {
                    break;
                }
                
                // Check if there's data to read
                if (!PeekNamedPipe(hChildStd_OUT_Rd, NULL, 0, NULL, &available, NULL) || available == 0) {
                    Sleep(100); // Wait a bit before trying again
                    continue;
                }
                
                // Read available data
                if (ReadFile(hChildStd_OUT_Rd, buffer, sizeof(buffer) - 1, &dwRead, NULL) && dwRead > 0) {
                    buffer[dwRead] = '\0';
                    full_output += buffer;
                    
                    // Parse progress information using more robust regex
                    try {
                        std::regex re(R"(\[download\]\s+([0-9\.]+)%)");
                        std::smatch match;
                        
                        // Process just the newly received data
                        std::string newOutput(buffer);
                        std::string::const_iterator searchStart(newOutput.cbegin());
                        while (std::regex_search(searchStart, newOutput.cend(), match, re)) {
                            // Update progress with the latest percentage
                            try {
                                float progress = std::stof(match[1].str());
                                item->progress = progress;
                            }
                            catch (...) {
                                // Ignore parsing errors
                            }
                            searchStart = match.suffix().first;
                        }
                    }
                    catch (...) {
                        // Ignore regex errors
                    }
                }
                else {
                    break;
                }
            }
            
            // Read any error output
            while (ReadFile(hChildStd_ERR_Rd, buffer, sizeof(buffer) - 1, &dwRead, NULL) && dwRead > 0) {
                buffer[dwRead] = '\0';
                error_output += buffer;
            }
            
            // If we have error output and the download failed, show it
            if (!error_output.empty()) {
                // Only log error output, don't show message box here to avoid UI blocks
                OutputDebugStringA(("yt-dlp error output:\n" + error_output).c_str());
            }
        }
        catch (...) {
            // Reading output failed, but we'll continue to check the exit code
        }
        
        // Get process exit code to determine success
        DWORD exitCode = 1; // Default to failure
        if (GetExitCodeProcess(pi.hProcess, &exitCode)) {
            success = (exitCode == 0);
        }

        // Clean up handles
        CloseHandle(pi.hThread);
        CloseHandle(hChildStd_OUT_Rd);
        CloseHandle(hChildStd_ERR_Rd);
        
        // Close process handle
        CloseHandle(pi.hProcess);
        item->hProcess = NULL;
    }
    else {
        // Process creation failed, clean up
        CloseHandle(hChildStd_OUT_Rd);
        CloseHandle(hChildStd_ERR_Rd);
        
        // Log the error for debugging
        std::wstring errorMsg = L"Failed to start yt-dlp process. Error code: " + std::to_wstring(lastError);
        errorMsg += L"\nCommand: " + command;
        OutputDebugStringW(errorMsg.c_str());
    }

    // Update item status if not already cancelled
    if (item->status != Cancelled) {
        item->status = success ? Completed : Failed;
    }

    return success ? 0 : 1;
}

// Dialog procedure for the download progress dialog
INT_PTR CALLBACK DownloadProgressProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static DownloadOptions* pOptions = nullptr;
    static DownloadItem* pItem = nullptr;
    static HANDLE hThread = NULL;
    static DWORD startTime = 0;
    static std::wstring downloadFilename;
    
    switch (message) {
    case WM_INITDIALOG: {
        if (lParam == 0) {
            // Invalid parameter
            MessageBox(hDlg, L"Invalid download options provided.", L"Error", MB_OK | MB_ICONERROR);
            DestroyWindow(hDlg);
            return (INT_PTR)TRUE;
        }
        
        try {
            pOptions = (DownloadOptions*)lParam;
            if (!pOptions) {
                MessageBox(hDlg, L"Invalid download options.", L"Error", MB_OK | MB_ICONERROR);
                DestroyWindow(hDlg);
                return (INT_PTR)TRUE;
            }
            
            // Store dialog handle in options
            pOptions->hDlg = hDlg;
            
            // Initialize progress bar
            HWND hProgress = GetDlgItem(hDlg, IDC_PROGRESS);
            if (hProgress) {
                SendMessage(hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
                SendMessage(hProgress, PBM_SETSTEP, 1, 0);
                SendMessage(hProgress, PBM_SETPOS, 0, 0);
            }
            
            // Update progress text with video URL
            if (!pOptions->videoUrl.empty()) {
                // Truncate long URLs for display
                std::wstring displayUrl = pOptions->videoUrl;
                if (displayUrl.length() > 80) {
                    displayUrl = displayUrl.substr(0, 77) + L"...";
                }
                SetDlgItemText(hDlg, IDC_PROGRESS_TEXT, displayUrl.c_str());
            } else {
                SetDlgItemText(hDlg, IDC_PROGRESS_TEXT, L"Downloading video...");
            }
            
            // Initialize progress text fields
            SetDlgItemText(hDlg, IDC_PROGRESS_PERCENT, L"0% completed");
            SetDlgItemText(hDlg, IDC_TIME_REMAINING, L"Time remaining: Calculating...");
            SetDlgItemText(hDlg, IDC_DOWNLOAD_SPEED, L"Download speed: Calculating...");
            
            // Create download item
            pItem = new DownloadItem();
            pItem->url = pOptions->videoUrl;
            pItem->resolution = pOptions->resolution;
            pItem->path = pOptions->path.empty() ? g_settings.defaultDownloadPath : pOptions->path;
            pItem->downloadSubtitles = pOptions->downloadSubtitles;
            pItem->status = Downloading;
            pItem->progress = 0;
            pItem->hProcess = NULL;
            pItem->threadId = 0;
            pItem->progressDlg = hDlg;
            
            // Record start time for calculating speed and remaining time
            startTime = GetTickCount();
            
            // Create download thread with appropriate error handling
            hThread = CreateThread(NULL, 0, DownloadThread, (LPVOID)pItem, 0, &pItem->threadId);
            if (hThread == NULL) {
                DWORD error = GetLastError();
                wchar_t buffer[256];
                swprintf_s(buffer, L"Failed to start download thread. Error code: %d", error);
                MessageBox(hDlg, buffer, L"Error", MB_OK | MB_ICONERROR);
                delete pItem;
                pItem = nullptr;
                DestroyWindow(hDlg);
                return (INT_PTR)TRUE;
            }
            
            // Set a timer to check download progress
            SetTimer(hDlg, 1, 500, NULL);
            
            return (INT_PTR)TRUE;
        }
        catch (...) {
            MessageBox(hDlg, L"An unexpected error occurred while initializing the download.", 
                      L"Error", MB_OK | MB_ICONERROR);
            if (pItem) {
                delete pItem;
                pItem = nullptr;
            }
            DestroyWindow(hDlg);
            return (INT_PTR)TRUE;
        }
    }
    
    case WM_TIMER:
        if (pItem) {
            // Calculate elapsed time in seconds
            DWORD currentTime = GetTickCount();
            double elapsedSec = (currentTime - startTime) / 1000.0;
            if (elapsedSec < 0.1) elapsedSec = 0.1; // Avoid division by zero
            
            // Update progress bar
            HWND hProgress = GetDlgItem(hDlg, IDC_PROGRESS);
            if (hProgress) {
                SendMessage(hProgress, PBM_SETPOS, (int)pItem->progress, 0);
            }
            
            // Update progress percentage text
            std::wstring progressText = std::to_wstring((int)pItem->progress) + L"% completed";
            SetDlgItemText(hDlg, IDC_PROGRESS_PERCENT, progressText.c_str());
            
            // Calculate and update download speed (assume average download size is ~30MB for a video)
            // This is an estimate since we don't know the actual bytes downloaded
            double estimatedTotalSize = 30 * 1024 * 1024; // 30 MB in bytes
            double bytesDownloaded = (pItem->progress / 100.0) * estimatedTotalSize;
            double bytesPerSec = bytesDownloaded / elapsedSec;
            
            std::wstring speedText;
            if (bytesPerSec > 1024 * 1024) {
                speedText = L"Download speed: " + std::to_wstring((int)(bytesPerSec / (1024 * 1024))) + L" MB/s";
            } else if (bytesPerSec > 1024) {
                speedText = L"Download speed: " + std::to_wstring((int)(bytesPerSec / 1024)) + L" KB/s";
            } else {
                speedText = L"Download speed: Calculating...";
            }
            SetDlgItemText(hDlg, IDC_DOWNLOAD_SPEED, speedText.c_str());
            
            // Calculate and update estimated time remaining
            if (pItem->progress > 0) {
                double remainingPercentage = 100.0 - pItem->progress;
                double timePerPercent = elapsedSec / pItem->progress;
                int remainingSeconds = (int)(remainingPercentage * timePerPercent);
                
                std::wstring timeText;
                if (remainingSeconds > 60) {
                    timeText = L"Time remaining: " + std::to_wstring(remainingSeconds / 60) + L" min " + 
                               std::to_wstring(remainingSeconds % 60) + L" sec";
                } else {
                    timeText = L"Time remaining: " + std::to_wstring(remainingSeconds) + L" sec";
                }
                SetDlgItemText(hDlg, IDC_TIME_REMAINING, timeText.c_str());
            }
            
            // Check if download is complete
            if (pItem->status == Completed) {
                KillTimer(hDlg, 1);
                MessageBox(hDlg, L"Download completed successfully!", L"Success", MB_OK | MB_ICONINFORMATION);
                
                // Clean up
                if (hThread) {
                    CloseHandle(hThread);
                    hThread = NULL;
                }
                
                delete pItem;
                pItem = nullptr;
                
                // Don't close the dialog automatically, wait for user to dismiss it
                SetDlgItemText(hDlg, IDC_PROGRESS_PERCENT, L"100% completed");
                SetDlgItemText(hDlg, IDC_TIME_REMAINING, L"Download completed successfully!");
                SetDlgItemText(hDlg, IDC_DOWNLOAD_SPEED, L"");
                
                // Change the cancel button text to "Close"
                SetDlgItemText(hDlg, IDCANCEL, L"Close");
            }
            else if (pItem->status == Failed) {
                KillTimer(hDlg, 1);
                MessageBox(hDlg, L"Download failed. Please try again.", L"Error", MB_OK | MB_ICONERROR);
                
                // Clean up
                if (hThread) {
                    CloseHandle(hThread);
                    hThread = NULL;
                }
                
                delete pItem;
                pItem = nullptr;
                
                DestroyWindow(hDlg);
            }
            else if (pItem->status == Cancelled) {
                KillTimer(hDlg, 1);
                
                // Clean up
                if (hThread) {
                    CloseHandle(hThread);
                    hThread = NULL;
                }
                
                delete pItem;
                pItem = nullptr;
                
                DestroyWindow(hDlg);
            }
        }
        return (INT_PTR)TRUE;
        
    case WM_COMMAND:
        if (LOWORD(wParam) == IDCANCEL) {
            try {
                KillTimer(hDlg, 1);
                
                if (pItem && pItem->hProcess) {
                    TerminateProcess(pItem->hProcess, 0);
                    pItem->status = Cancelled;
                }
                
                if (hThread) {
                    // Give the thread a moment to clean up
                    WaitForSingleObject(hThread, 1000);
                    CloseHandle(hThread);
                    hThread = NULL;
                }
                
                if (pItem) {
                    delete pItem;
                    pItem = nullptr;
                }
            }
            catch (...) {
                // Ignore errors when trying to terminate the process
            }
            DestroyWindow(hDlg);
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDC_HIDE_BUTTON) {
            // Hide the dialog but keep the download going
            ShowWindow(hDlg, SW_HIDE);
            return (INT_PTR)TRUE;
        }
        break;
        
    case WM_DESTROY:
        KillTimer(hDlg, 1);
        if (hThread) {
            CloseHandle(hThread);
            hThread = NULL;
        }
        if (pItem) {
            delete pItem;
            pItem = nullptr;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Dialog procedure for download options
INT_PTR CALLBACK DownloadOptionsProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static DownloadOptions* pOptions;
    switch (message) {
    case WM_INITDIALOG:
        pOptions = (DownloadOptions*)lParam;
        // Populate resolutions (example, you can get these from yt-dlp)
        SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_ADDSTRING, 0, (LPARAM)L"Best");
        SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_ADDSTRING, 0, (LPARAM)L"1080p");
        SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_ADDSTRING, 0, (LPARAM)L"720p");
        SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_ADDSTRING, 0, (LPARAM)L"Audio Only (mp3)");
        SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_SETCURSEL, 0, 0);
        SetDlgItemText(hDlg, IDC_EDIT_PATH, g_settings.defaultDownloadPath.c_str());
        return (INT_PTR)TRUE;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            wchar_t buffer[MAX_PATH];
            GetDlgItemText(hDlg, IDC_EDIT_PATH, buffer, MAX_PATH);
            pOptions->path = buffer;
            int sel = SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_GETCURSEL, 0, 0);
            SendDlgItemMessage(hDlg, IDC_COMBO_RESOLUTION, CB_GETLBTEXT, sel, (LPARAM)buffer);
            pOptions->resolution = buffer;
            pOptions->downloadSubtitles = IsDlgButtonChecked(hDlg, IDC_CHECK_SUBTITLES) == BST_CHECKED;
            EndDialog(hDlg, IDOK);
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDC_BUTTON_BROWSE) {
            BROWSEINFOW bi = { 0 };
            bi.lpszTitle = L"Select a download folder";
            bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
            if (pidl) {
                wchar_t path[MAX_PATH];
                if (SHGetPathFromIDListW(pidl, path)) {
                    SetDlgItemText(hDlg, IDC_EDIT_PATH, path);
                }
                CoTaskMemFree(pidl);
            }
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Dialog procedure for settings
INT_PTR CALLBACK SettingsProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        CheckDlgButton(hDlg, IDC_CHECK_ADBLOCK_STARTUP, g_settings.adBlockOnStartup ? BST_CHECKED : BST_UNCHECKED);
        SetDlgItemText(hDlg, IDC_EDIT_DEFAULT_PATH, g_settings.defaultDownloadPath.c_str());
        
        // Set theme radio buttons
        switch (g_settings.themeMode) {
            case AppSettings::ThemeMode::Light:
                CheckDlgButton(hDlg, IDC_RADIO_LIGHT_MODE, BST_CHECKED);
                break;
            case AppSettings::ThemeMode::Dark:
                CheckDlgButton(hDlg, IDC_RADIO_DARK_MODE, BST_CHECKED);
                break;
            case AppSettings::ThemeMode::System:
                CheckDlgButton(hDlg, IDC_RADIO_SYSTEM_THEME, BST_CHECKED);
                break;
        }
        
        return (INT_PTR)TRUE;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            g_settings.adBlockOnStartup = IsDlgButtonChecked(hDlg, IDC_CHECK_ADBLOCK_STARTUP) == BST_CHECKED;
            
            // Get theme mode from radio buttons
            if (IsDlgButtonChecked(hDlg, IDC_RADIO_LIGHT_MODE) == BST_CHECKED) {
                g_settings.themeMode = AppSettings::ThemeMode::Light;
            }
            else if (IsDlgButtonChecked(hDlg, IDC_RADIO_DARK_MODE) == BST_CHECKED) {
                g_settings.themeMode = AppSettings::ThemeMode::Dark;
            }
            else if (IsDlgButtonChecked(hDlg, IDC_RADIO_SYSTEM_THEME) == BST_CHECKED) {
                g_settings.themeMode = AppSettings::ThemeMode::System;
            }
            
            wchar_t path[MAX_PATH];
            GetDlgItemText(hDlg, IDC_EDIT_DEFAULT_PATH, path, MAX_PATH);
            g_settings.defaultDownloadPath = path;
            SaveSettings(); // Save settings when OK is clicked
            
            // Apply theme based on settings
            ApplyThemeMode();
            
            EndDialog(hDlg, IDOK);
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDC_BUTTON_DEFAULT_BROWSE) {
            BROWSEINFOW bi = { 0 };
            bi.lpszTitle = L"Select a default download folder";
            bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
            if (pidl) {
                wchar_t path[MAX_PATH];
                if (SHGetPathFromIDListW(pidl, path)) {
                    SetDlgItemText(hDlg, IDC_EDIT_DEFAULT_PATH, path);
                }
                CoTaskMemFree(pidl);
            }
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Helper to resize WebView2 when window size changes
void ResizeWebView2(HWND hWnd) {
    if (g_webViewController) {
        RECT bounds;
        GetClientRect(hWnd, &bounds);
        g_webViewController->put_Bounds(bounds);
    }
}

// Helper function to inject ad-blocking JavaScript
void InjectAdBlockScript() {
    if (g_webView) {
        // This script uses a MutationObserver to remove ads as they are dynamically added to the page.
        // It targets a wider range of ad-related selectors for better coverage.
        std::wstring script = LR"(
            const adSelectors = [
                '.ad-showing',
                '.video-ads',
                'ytd-ad-slot-renderer',
                'ytd-promoted-sparkles-web-renderer',
                'ytd-in-feed-ad-layout-renderer',
                '#player-ads',
                '#masthead-ad',
                '.ytd-ad-slot-renderer',
                '.ytp-ad-module'
            ];

            const removeAds = () => {
                const ads = document.querySelectorAll(adSelectors.join(', '));
                ads.forEach(ad => {
                    let parent = ad.parentElement;
                    while (parent && parent.tagName.toLowerCase() !== 'ytd-rich-item-renderer' && parent.tagName.toLowerCase() !== 'ytd-video-renderer' && parent.tagName.toLowerCase() !== 'body') {
                        parent = parent.parentElement;
                    }
                    if (parent && (parent.tagName.toLowerCase() === 'ytd-rich-item-renderer' || parent.tagName.toLowerCase() === 'ytd-video-renderer')) {
                        parent.style.display = 'none';
                    } else {
                        ad.style.display = 'none';
                    }
                });
            };

            const observer = new MutationObserver((mutations) => {
                removeAds();
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true
            });

            // Initial cleanup
            removeAds();
        )";
        g_webView->AddScriptToExecuteOnDocumentCreated(script.c_str(), nullptr);
    }
}

void ToggleDarkMode() {
    if (g_settings.themeMode == AppSettings::ThemeMode::Dark) {
        SetLightMode();
    } else {
        SetDarkMode();
    }
}

void TogglePictureInPicture() {
    if (g_webView) {
        // This script attempts to find the primary video element and request picture-in-picture.
        std::wstring script = LR"(
            const video = document.querySelector('video');
            if (video && document.pictureInPictureElement !== video) {
                video.requestPictureInPicture();
            } else if (document.pictureInPictureElement) {
                document.exitPictureInPicture();
            }
        )";
        g_webView->ExecuteScript(script.c_str(), nullptr);
    }
}

void ToggleAudioOnlyPlayer() {
    if (g_webViewController) {
        g_isAudioOnly = !g_isAudioOnly;
        g_webViewController->put_IsVisible(!g_isAudioOnly);
    }
}

// Helper to get current URL from WebView2
std::wstring GetCurrentWebViewUrl() {
    if (!g_webView) return L"";
    LPWSTR uri = nullptr;
    if (SUCCEEDED(g_webView->get_Source(&uri)) && uri) {
        std::wstring result(uri);
        CoTaskMemFree(uri);
        return result;
    }
    return L"";
}

// Determine if system theme is dark mode
bool IsSystemInDarkMode() {
    DWORD value = 0;
    DWORD dataSize = sizeof(value);
    HKEY hKey;
    
    if (RegOpenKeyExW(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
        0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        
        if (RegQueryValueExW(hKey, L"AppsUseLightTheme", nullptr, nullptr,
            reinterpret_cast<LPBYTE>(&value), &dataSize) == ERROR_SUCCESS) {
            RegCloseKey(hKey);
            return value == 0; // 0 means dark theme is enabled
        }
        RegCloseKey(hKey);
    }
    
    return false; // Default to light mode if registry read fails
}

// Apply theme based on current settings
void ApplyThemeMode() {
    if (!g_webView) return;
    
    bool shouldUseDarkMode = false;
    
    switch (g_settings.themeMode) {
        case AppSettings::ThemeMode::Light:
            shouldUseDarkMode = false;
            break;
        case AppSettings::ThemeMode::Dark:
            shouldUseDarkMode = true;
            break;
        case AppSettings::ThemeMode::System:
            shouldUseDarkMode = IsSystemInDarkMode();
            break;
    }
    
    // Apply the appropriate theme
    std::wstring script = LR"(
        const html = document.documentElement;
        if ()";
    
    script += shouldUseDarkMode ? L"true" : L"false";
    
    script += LR"() {
            if (!html.hasAttribute('dark')) {
                html.setAttribute('dark', 'true');
            }
        } else {
            if (html.hasAttribute('dark')) {
                html.removeAttribute('dark');
            }
        }
    )";
    
    g_webView->ExecuteScript(script.c_str(), nullptr);
}

// Theme-related functions
void SetLightMode() {
    if (!g_webView) return;
    
    g_settings.themeMode = AppSettings::ThemeMode::Light;
    SaveSettings();
    
    // Apply light theme to WebView
    std::wstring script = LR"(
        const html = document.documentElement;
        if (html.hasAttribute('dark')) {
            html.removeAttribute('dark');
        }
    )";
    
    g_webView->ExecuteScript(script.c_str(), nullptr);
    MessageBox(nullptr, L"Light mode enabled", L"Theme Changed", MB_OK | MB_ICONINFORMATION);
}

void SetDarkMode() {
    if (!g_webView) return;
    
    g_settings.themeMode = AppSettings::ThemeMode::Dark;
    SaveSettings();
    
    // Apply dark theme to WebView
    std::wstring script = LR"(
        const html = document.documentElement;
        if (!html.hasAttribute('dark')) {
            html.setAttribute('dark', 'true');
        }
    )";
    
    g_webView->ExecuteScript(script.c_str(), nullptr);
    MessageBox(nullptr, L"Dark mode enabled", L"Theme Changed", MB_OK | MB_ICONINFORMATION);
}

void UseSystemTheme() {
    if (!g_webView) return;
    
    g_settings.themeMode = AppSettings::ThemeMode::System;
    SaveSettings();
    
    // Apply theme based on system settings
    ApplyThemeMode();
    MessageBox(nullptr, L"Using system theme", L"Theme Changed", MB_OK | MB_ICONINFORMATION);
}

// Helper function to directly start a download
void StartDownload(const std::wstring& url, HWND parentWindow = NULL) {
    if (url.empty()) {
        MessageBox(NULL, L"Cannot download empty URL.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }
    
    try {
        DownloadOptions options;
        options.resolution = L"Best";
        options.path = g_settings.defaultDownloadPath;
        options.videoUrl = url;
        options.downloadSubtitles = false;
        options.hProcess = NULL;
        
        if (!parentWindow) {
            parentWindow = GetActiveWindow();
        }
        
        if (DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_DOWNLOAD_OPTIONS), parentWindow, 
                         DownloadOptionsProc, (LPARAM)&options) == IDOK) {
            // Start the download with progress dialog (modeless)
            HWND hProgress = CreateDialogParam(hInst, MAKEINTRESOURCE(IDD_DOWNLOAD_PROGRESS), parentWindow, DownloadProgressProc, (LPARAM)&options);
            ShowWindow(hProgress, SW_SHOW);
        }
    }
    catch (const std::exception& e) {
        char buffer[256];
        sprintf_s(buffer, "Failed to start download: %s", e.what());
        
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, buffer, -1, NULL, 0);
        std::wstring wErrorMsg(size_needed, 0);
        MultiByteToWideChar(CP_UTF8, 0, buffer, -1, &wErrorMsg[0], size_needed);
        
        MessageBox(NULL, wErrorMsg.c_str(), L"Download Error", MB_OK | MB_ICONERROR);
    }
    catch (...) {
        MessageBox(NULL, L"An unexpected error occurred when preparing download.", 
                  L"Download Error", MB_OK | MB_ICONERROR);
    }
}

#define MAX_LOADSTRING 100

// Global Variables:
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

// New dialog procedure for the Download Manager
INT_PTR CALLBACK DownloadManagerProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    static std::vector<DownloadItem*> downloadItems;
    static std::pair<std::vector<std::wstring>, DownloadOptions>* managerData = nullptr;

    switch (message) {
    case WM_INITDIALOG: {
        managerData = (std::pair<std::vector<std::wstring>, DownloadOptions>*)lParam;
        if (!managerData) {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)FALSE;
        }

        HWND hList = GetDlgItem(hDlg, IDC_DOWNLOAD_LIST);
        ListView_SetExtendedListViewStyle(hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

        // Add columns
        LVCOLUMNW lvc = { 0 };
        lvc.mask = LVCF_TEXT | LVCF_WIDTH;
        lvc.pszText = (LPWSTR)L"Video URL";
        lvc.cx = 200;
        ListView_InsertColumn(hList, 0, &lvc);

        lvc.pszText = (LPWSTR)L"Progress";
        lvc.cx = 80;
        ListView_InsertColumn(hList, 1, &lvc);

        lvc.pszText = (LPWSTR)L"Status";
        lvc.cx = 100;
        ListView_InsertColumn(hList, 2, &lvc);

        // Populate the list with videos to download
        for (const auto& url : managerData->first) {
            DownloadItem* newItem = new DownloadItem();
            newItem->url = url;
            newItem->resolution = managerData->second.resolution;
            newItem->path = managerData->second.path;
            newItem->downloadSubtitles = managerData->second.downloadSubtitles;
            newItem->status = Downloading;
            newItem->progress = 0;
            newItem->hProcess = NULL;
            newItem->hThread = NULL;
            newItem->threadId = 0;
            newItem->progressDlg = hDlg; // The manager is the dialog
            newItem->startTime = 0;
            downloadItems.push_back(newItem);

            LVITEMW lvi = { 0 };
            lvi.mask = LVIF_TEXT;
            lvi.iItem = downloadItems.size() - 1;
            lvi.pszText = (LPWSTR)url.c_str();
            ListView_InsertItem(hList, &lvi);
            ListView_SetItemText(hList, lvi.iItem, 1, (LPWSTR)L"0%");
            ListView_SetItemText(hList, lvi.iItem, 2, (LPWSTR)L"Queued");
        }

        // Start the first download
        if (!downloadItems.empty()) {
            downloadItems[0]->startTime = GetTickCount();
            downloadItems[0]->hThread = CreateThread(NULL, 0, DownloadThread, (LPVOID)downloadItems[0], 0, &downloadItems[0]->threadId);
        }

        SetTimer(hDlg, 1, 1000, NULL); // Timer to update progress
        return (INT_PTR)TRUE;
    }

    case WM_TIMER: {
        HWND hList = GetDlgItem(hDlg, IDC_DOWNLOAD_LIST);
        bool allDone = true;
        for (size_t i = 0; i < downloadItems.size(); ++i) {
            DownloadItem* item = downloadItems[i];
            if (item->status == Downloading) {
                allDone = false;
                wchar_t progressText[16];
                swprintf_s(progressText, L"%.1f%%", item->progress);
                ListView_SetItemText(hList, i, 1, progressText);

                // If process finished but progress didn't reach 100, we'll still rely on status
                if (item->progress >= 100.0) {
                    item->status = Completed;
                    ListView_SetItemText(hList, i, 2, (LPWSTR)L"Completed");
                    // Start next download if not already started
                    if (i + 1 < downloadItems.size()) {
                        DownloadItem* nextItem = downloadItems[i + 1];
                        if (nextItem->hThread == NULL && nextItem->status == Downloading) {
                            nextItem->startTime = GetTickCount();
                            nextItem->hThread = CreateThread(NULL, 0, DownloadThread, (LPVOID)nextItem, 0, &nextItem->threadId);
                        }
                    }
                }
            } else if (item->status == Completed) {
                // Ensure UI shows completed and progress set to 100
                ListView_SetItemText(hList, i, 2, (LPWSTR)L"Completed");
                ListView_SetItemText(hList, i, 1, (LPWSTR)L"100.0%");
                // Start next download if not already started
                if (i + 1 < downloadItems.size()) {
                    DownloadItem* nextItem = downloadItems[i + 1];
                    if (nextItem->hThread == NULL && nextItem->status == Downloading) {
                        nextItem->startTime = GetTickCount();
                        nextItem->hThread = CreateThread(NULL, 0, DownloadThread, (LPVOID)nextItem, 0, &nextItem->threadId);
                        allDone = false;
                    }
                }
            } else if (item->status == Failed) {
                ListView_SetItemText(hList, i, 2, (LPWSTR)L"Failed");
                // If a failed item hasn't started next, start it
                if (i + 1 < downloadItems.size()) {
                    DownloadItem* nextItem = downloadItems[i + 1];
                    if (nextItem->hThread == NULL && nextItem->status == Downloading) {
                        nextItem->startTime = GetTickCount();
                        nextItem->hThread = CreateThread(NULL, 0, DownloadThread, (LPVOID)nextItem, 0, &nextItem->threadId);
                        allDone = false;
                    }
                }
            }
        }
        if (allDone) {
            KillTimer(hDlg, 1);
        }
        return (INT_PTR)TRUE;
    }

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            KillTimer(hDlg, 1);
            for (auto& item : downloadItems) {
                if (item->hThread) {
                    if (item->hProcess) TerminateProcess(item->hProcess, 1);
                    CloseHandle(item->hThread);
                }
                delete item;
            }
            downloadItems.clear();
            delete managerData;
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;

    case WM_DESTROY:
        KillTimer(hDlg, 1);
        for (auto& item : downloadItems) {
            if (item->hThread) {
                if (item->hProcess) TerminateProcess(item->hProcess, 1);
                CloseHandle(item->hThread);
            }
            delete item;
        }
        downloadItems.clear();
        if(managerData) delete managerData;
        break;
    }
    return (INT_PTR)FALSE;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Initialize COM
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
    {
        MessageBox(nullptr, L"Failed to initialize COM.", L"Fatal Error", MB_OK | MB_ICONERROR);
        return FALSE;
    }

    LoadSettings(); // Load settings on startup

    // Check if WebView2 Runtime is installed
    if (!IsWebView2RuntimeInstalled()) {
        if (!InstallWebView2Runtime(NULL)) {
            MessageBox(NULL, L"WebView2 Runtime installation failed. The application will now exit.", L"Error", MB_OK | MB_ICONERROR);
            CoUninitialize();
            return FALSE;
        }
    }

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_YOUTUBEPLUS, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_YOUTUBEPLUS));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    CoUninitialize();
    return (int) msg.wParam;
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_YOUTUBEPLUS));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_YOUTUBEPLUS);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        // Check if WebView2 Runtime is installed
        if (!IsWebView2RuntimeInstalled()) {
            if (InstallWebView2Runtime(hWnd)) {
                // Restart the application after successful installation
                wchar_t exePath[MAX_PATH];
                GetModuleFileNameW(NULL, exePath, MAX_PATH);
                ShellExecuteW(NULL, NULL, exePath, NULL, NULL, SW_SHOWNORMAL);
                DestroyWindow(hWnd);
                return 0;
            }
        }
        
        // Initialize WebView2 with a custom user data folder
        {
            // Create a dedicated user data folder in AppData
            PWSTR appDataPath = NULL;
            if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &appDataPath))) {
                std::wstring userDataFolder = appDataPath;
                CoTaskMemFree(appDataPath);
                userDataFolder += L"\\YoutubePlus\\WebView2Data";
                
                // Create the directory if it doesn't exist (create parent dirs too)
                std::wstring dirPath = userDataFolder.substr(0, userDataFolder.find_last_of(L"\\"));
                CreateDirectoryW(dirPath.c_str(), NULL);
                CreateDirectoryW(userDataFolder.c_str(), NULL);
                
                // Initialize WebView2 with the user data folder and improved error handling
                HRESULT envResult = CreateCoreWebView2EnvironmentWithOptions(nullptr, userDataFolder.c_str(), nullptr,
                    Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
                        [hWnd](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
                            if (FAILED(result)) {
                                wchar_t errorMsg[256];
                                swprintf_s(errorMsg, L"Failed to create WebView2 environment (Error: 0x%08X). Please restart your computer and try again.", result);
                                MessageBox(hWnd, errorMsg, L"WebView2 Error", MB_OK | MB_ICONERROR);
                                return result;
                            }
                            
                            env->CreateCoreWebView2Controller(hWnd, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                                [hWnd](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                                    if (FAILED(result)) {
                                        wchar_t errorMsg[256];
                                        swprintf_s(errorMsg, L"Failed to create WebView2 controller (Error: 0x%08X).", result);
                                        MessageBox(hWnd, errorMsg, L"WebView2 Error", MB_OK | MB_ICONERROR);
                                        return result;
                                    }
                                    
                                    if (controller) {
                                        g_webViewController = controller;
                                        g_webViewController->AddRef();
                                        controller->get_CoreWebView2(&g_webView);
                                        if (g_webView) {
                                            g_webView->AddRef();
                                        }
                                        RECT bounds;
                                        GetClientRect(hWnd, &bounds);
                                        controller->put_Bounds(bounds);
                                        controller->put_IsVisible(TRUE); // Make the WebView visible
                                        
                                        if (g_webView) {
                                            // Add improved error handling for navigation
                                            EventRegistrationToken navToken;
                                            g_webView->add_NavigationCompleted(
                                                Callback<ICoreWebView2NavigationCompletedEventHandler>(
                                                    [](ICoreWebView2* webview, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
                                                        BOOL success;
                                                        args->get_IsSuccess(&success);
                                                        if (!success) {
                                                            COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus;
                                                            args->get_WebErrorStatus(&webErrorStatus);
                                                            
                                                            // Handle specific error cases like no internet
                                                            if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED) {
                                                                webview->ExecuteScript(
                                                                    L"document.body.innerHTML = '<div style=\"padding: 20px; text-align: center; font-family: Arial, sans-serif;\">"
                                                                    L"<h2>No Internet Connection</h2>"
                                                                    L"<p>Please check your connection and try again.</p>"
                                                                    L"<button onclick=\"window.location.reload()\">Reload</button>"
                                                                    L"</div>';",
                                                                    nullptr);
                                                            }
                                                        }
                                                        return S_OK;
                                                    }).Get(), &navToken);
                                            
                                            // Add content loading handler for settings application
                                            EventRegistrationToken contentToken;
                                            g_webView->add_ContentLoading(
                                                Callback<ICoreWebView2ContentLoadingEventHandler>(
                                                    [](ICoreWebView2* webview, ICoreWebView2ContentLoadingEventArgs* args) -> HRESULT {
                                                        // Apply settings when DOM content loads
                                                        if (g_settings.adBlockOnStartup) {
                                                            InjectAdBlockScript();
                                                        }
                                                        
                                                        // Apply theme mode
                                                        ApplyThemeMode();
                                                        
                                                        return S_OK;
                                                    }).Get(), &contentToken);
                                            
                                            // Navigate to YouTube
                                            g_webView->Navigate(L"https://www.youtube.com");
                                        }
                                    }
                                    return S_OK;
                                }).Get());
                            return S_OK;
                        }).Get());
                
                // Check if environment creation was successful
                if (FAILED(envResult)) {
                    MessageBox(hWnd, L"Failed to initialize WebView2. Please ensure WebView2 Runtime is properly installed.", 
                              L"WebView2 Error", MB_OK | MB_ICONERROR);
                }
            } else {
                MessageBox(hWnd, L"Failed to access AppData folder for WebView2 user data.", L"Error", MB_OK | MB_ICONERROR);
            }
        }
        break;
    case WM_SIZE:
        ResizeWebView2(hWnd);
        break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            case IDM_DOWNLOAD:
                // Get current video URL from WebView2
                if (g_webView) {
                    try {
                        std::wstring url = GetCurrentWebViewUrl();
                        
                        if (url.empty()) {
                            MessageBox(hWnd, L"Failed to get current URL. Please try again.", 
                                      L"Error", MB_OK | MB_ICONERROR);
                            break;
                        }
                        
                        if (url.find(L"youtube.com/watch") != std::wstring::npos) {
                            StartDownload(url, hWnd);
                        } else {
                            MessageBox(hWnd, L"Please navigate to a YouTube video page to download.", 
                                      L"Download", MB_OK | MB_ICONINFORMATION);
                        }
                    }
                    catch (const std::exception& e) {
                        char buffer[256];
                        sprintf_s(buffer, "Error during download preparation: %s", e.what());
                        
                        int size_needed = MultiByteToWideChar(CP_UTF8, 0, buffer, -1, NULL, 0);
                        std::wstring wErrorMsg(size_needed, 0);
                        MultiByteToWideChar(CP_UTF8, 0, buffer, -1, &wErrorMsg[0], size_needed);
                        
                        MessageBox(hWnd, wErrorMsg.c_str(), L"Error", MB_OK | MB_ICONERROR);
                    }
                    catch (...) {
                        MessageBox(hWnd, L"An unknown error occurred while preparing the download.", 
                                  L"Error", MB_OK | MB_ICONERROR);
                    }
                } else {
                    // More detailed error message for WebView initialization issue
                    wchar_t errorMsg[512];
                    swprintf_s(errorMsg, 
                        L"WebView2 not initialized. This could be due to:\n"
                        L"1. WebView2 Runtime not properly installed\n"
                        L"2. Missing WebView2Loader.dll\n"
                        L"3. Insufficient permissions\n\n"
                        L"Please reinstall the application or try running as administrator.");
                    MessageBox(hWnd, errorMsg, L"WebView2 Error", MB_OK | MB_ICONERROR);
                }
                break;
            case IDM_DOWNLOAD_PLAYLIST:
                // Get current playlist URL from WebView2
                if (g_webView) {
                    std::wstring url = GetCurrentWebViewUrl();
                    bool isPlaylistPage = url.find(L"youtube.com/playlist") != std::wstring::npos;
                    bool isWatchListPage = url.find(L"youtube.com/watch") != std::wstring::npos && 
                                          url.find(L"list=") != std::wstring::npos;
                                          
                    if (isPlaylistPage || isWatchListPage) {
                        // Pass the playlist URL to the playlist dialog
                        DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_PLAYLIST), hWnd, PlaylistDialogProc, (LPARAM)&url);
                    } else {
                        MessageBox(hWnd, L"Please navigate to a YouTube playlist page or a video that's part of a playlist.", 
                                  L"Download Playlist", MB_OK | MB_ICONINFORMATION);
                    }
                } else {
                    MessageBox(hWnd, L"WebView2 not initialized.", L"Download", MB_OK | MB_ICONERROR);
                }
                break;
            case IDM_ADBLOCK:
                InjectAdBlockScript();
                MessageBox(hWnd, L"AdBlock script injected (if WebView2 is running).", L"AdBlock", MB_OK | MB_ICONINFORMATION);
                break;
            case IDM_SETTINGS:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_SETTINGS), hWnd, SettingsProc);
                break;
            case IDM_TOGGLE_DARK_MODE:
                ToggleDarkMode();
                break;
            case IDM_LIGHT_MODE:
                SetLightMode();
                break;
            case IDM_DARK_MODE:
                SetDarkMode();
                break;
            case IDM_SYSTEM_THEME:
                UseSystemTheme();
                break;
            case IDM_PIP:
                TogglePictureInPicture();
                break;
            case IDM_AUDIO_ONLY_PLAYER:
                ToggleAudioOnlyPlayer();
                break;
            case IDM_NAV_BACK:
                if (g_webView) {
                    g_webView->GoBack();
                }
                break;
            case IDM_NAV_FORWARD:
                if (g_webView) {
                    g_webView->GoForward();
                }
                break;
            case IDM_NAV_RELOAD:
                if (g_webView) {
                    g_webView->Reload();
                }
                break;
            case IDM_NAV_HOME:
                if (g_webView) {
                    g_webView->Navigate(L"https://www.youtube.com");
                }
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        if (g_webViewController)
        {
            g_webViewController->Close();
            g_webViewController->Release();
            g_webViewController = nullptr;
        }
        if (g_webView)
        {
            g_webView->Release();
            g_webView = nullptr;
        }
        SaveSettings(); // Save settings on exit
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Helper function to check if WebView2 Runtime is installed
bool IsWebView2RuntimeInstalled() {
    HKEY hKey;
    const wchar_t* registryKeys[] = {
        L"SOFTWARE\\WOW6432Node\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
        L"SOFTWARE\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
        L"Software\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    };
    
    for (const auto& keyPath : registryKeys) {
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
            RegCloseKey(hKey);
            return true;
        }
        if (RegOpenKeyExW(HKEY_CURRENT_USER, keyPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
            RegCloseKey(hKey);
            return true;
        }
    }
    
    // Alternative check: Try to load WebView2Loader.dll directly
    // This is useful in MSI installed scenarios where the registry might not be properly set
    HMODULE hWebView2Loader = LoadLibraryW(L"WebView2Loader.dll");
    if (hWebView2Loader) {
        FreeLibrary(hWebView2Loader);
        return true;
    }
    
    return false;
}

// Helper function to install WebView2 Runtime
bool InstallWebView2Runtime(HWND hWnd) {
    // Get the current executable directory
    wchar_t exePath[MAX_PATH];
    if (!GetModuleFileNameW(NULL, exePath, MAX_PATH)) {
        return false;
    }
    
    std::wstring exeDir = exePath;
    size_t pos = exeDir.find_last_of(L"\\/");
    if (pos == std::wstring::npos) {
        return false;
    }
    
    // Look for WebView2 installer in the application directory
    std::wstring installerPath = exeDir.substr(0, pos) + L"\\MicrosoftEdgeWebview2Setup.exe";
    
    // Check if the installer exists
    DWORD attrs = GetFileAttributesW(installerPath.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) {
        // Also check in the current directory
        installerPath = L"MicrosoftEdgeWebview2Setup.exe";
        attrs = GetFileAttributesW(installerPath.c_str());
        
        if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) {
            // Installer not found, ask user to download
            int result = MessageBox(hWnd, 
                L"WebView2 Runtime is not installed. Would you like to download and install it now?", 
                L"WebView2 Runtime Required", 
                MB_YESNO | MB_ICONQUESTION);
                
            if (result == IDYES) {
                ShellExecuteW(NULL, L"open", L"https://go.microsoft.com/fwlink/p/?LinkId=2124703", NULL, NULL, SW_SHOWNORMAL);
            }
            return false;
        }
    }
    
    // Run the installer with elevated privileges
    SHELLEXECUTEINFOW sei = { 0 };
    sei.cbSize = sizeof(SHELLEXECUTEINFOW);
    sei.lpVerb = L"runas";
    sei.lpFile = installerPath.c_str();
    sei.lpParameters = L"/silent /install";
    sei.nShow = SW_SHOW;
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    
    if (!ShellExecuteExW(&sei)) {
        MessageBox(hWnd, 
            L"Failed to launch WebView2 installer. Please install it manually.", 
            L"Installation Failed", 
            MB_OK | MB_ICONERROR);
        return false;
    }
    
    // Wait for installer to finish with a timeout
    if (sei.hProcess) {
        MessageBox(hWnd, 
            L"WebView2 Runtime is being installed. The application will restart after installation completes.", 
            L"Installing WebView2", 
            MB_OK | MB_ICONINFORMATION);
            
        // Wait with a 2-minute timeout
        WaitForSingleObject(sei.hProcess, 120000);
        CloseHandle(sei.hProcess);
        return true;
    }
    
    return false;
}