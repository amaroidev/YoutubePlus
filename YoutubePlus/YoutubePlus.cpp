// YoutubePlus.cpp : Defines the entry point for the application.
//

#include "pch.h"
#include "framework.h"
#include "YoutubePlus.h"

#include <objbase.h>
#include <string>
#include <wrl.h>
#include <WebView2.h>

#pragma comment(lib, "WebView2LoaderStatic.lib")

using namespace Microsoft::WRL;

// WebView2 global pointers
ICoreWebView2Controller* g_webViewController = nullptr;
ICoreWebView2* g_webView = nullptr;

// Helper to resize WebView2 when window size changes
void ResizeWebView2(HWND hWnd) {
    if (g_webViewController) {
        RECT bounds;
        GetClientRect(hWnd, &bounds);
        g_webViewController->put_Bounds(bounds);
    }
}

// Helper function to launch yt-dlp for downloads
void StartDownload(const std::wstring& videoUrl) {
    std::wstring command = L"yt-dlp.exe \"" + videoUrl + L"\"";
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    if (CreateProcessW(nullptr, &command[0], nullptr, nullptr, FALSE, 0, nullptr, nullptr, &si, &pi)) {
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    } else {
        MessageBoxW(nullptr, L"Failed to start yt-dlp. Make sure yt-dlp.exe is available.", L"Download Error", MB_OK | MB_ICONERROR);
    }
}

// Helper function to inject ad-blocking JavaScript
void InjectAdBlockScript() {
    if (g_webView) {
        std::wstring script = L"document.querySelectorAll('ytd-ad-slot-renderer, .ad-container, .video-ads').forEach(e => e.remove());";
        g_webView->AddScriptToExecuteOnDocumentCreated(script.c_str(), nullptr);
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

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

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
        // Initialize WebView2 and navigate to YouTube
        CreateCoreWebView2EnvironmentWithOptions(nullptr, nullptr, nullptr,
            Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
                [hWnd](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
                    if (FAILED(result)) {
                        MessageBox(hWnd, L"Failed to create WebView2 environment. Please ensure the WebView2 Runtime is installed.", L"Error", MB_OK | MB_ICONERROR);
                        return result;
                    }
                    env->CreateCoreWebView2Controller(hWnd, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [hWnd](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                            if (FAILED(result)) {
                                MessageBox(hWnd, L"Failed to create WebView2 controller.", L"Error", MB_OK | MB_ICONERROR);
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
                                    g_webView->Navigate(L"https://www.youtube.com");
                                }
                            }
                            return S_OK;
                        }).Get());
                    return S_OK;
                }).Get());
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
                    std::wstring url = GetCurrentWebViewUrl();
                    if (url.find(L"youtube.com/watch") != std::wstring::npos) {
                        StartDownload(url);
                    } else {
                        MessageBox(hWnd, L"Please navigate to a YouTube video page to download.", L"Download", MB_OK | MB_ICONINFORMATION);
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
                MessageBox(hWnd, L"Settings feature coming soon!", L"Settings", MB_OK | MB_ICONINFORMATION);
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
