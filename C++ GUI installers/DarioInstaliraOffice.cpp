#define UNICODE
#define _UNICODE
#include <windows.h>
#include <string>
#include <urlmon.h>
#include <shellapi.h>
#include <thread>

#pragma comment(lib, "urlmon.lib")

const wchar_t* versions[] = {
    L"Office 2019 Pro Plus HR",
    L"Office 2021 Pro Plus HR",
    L"Office 2024 Pro Plus HR",
    L"Office 365 Enterprise"
};

const wchar_t* urls[] = {
    L"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_2019_Hrvatski.exe",
    L"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_2021_Hrvatski.exe",
    L"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_2024_Hrvatski.exe",
    L"https://github.com/Batman123n/Office-installation-script/releases/download/Office_installers-v1/Office_365_Hrvatski.exe"
};

const wchar_t* files[] = {
    L"Office_2019_Hrvatski.exe",
    L"Office_2021_Hrvatski.exe",
    L"Office_2024_Hrvatski.exe",
    L"Office_365_Hrvatski.exe"
};

HWND hComboBox;
HWND hStatusText;
HFONT hFont;
HBRUSH hBackgroundBrush = NULL;

bool DownloadFile(const wchar_t* url, const wchar_t* filename) {
    return SUCCEEDED(URLDownloadToFile(NULL, url, filename, 0, NULL));
}

void RunFileAndDeleteAfter(const wchar_t* filename) {
    SHELLEXECUTEINFO sei = { sizeof(SHELLEXECUTEINFO) };
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb = L"open";
    sei.lpFile = filename;
    sei.nShow = SW_SHOWNORMAL;

    if (ShellExecuteEx(&sei)) {
        WaitForSingleObject(sei.hProcess, INFINITE);
        CloseHandle(sei.hProcess);
        DeleteFile(filename);
    }
}

void UpdateStatus(HWND hwnd, const wchar_t* msg) {
    PostMessage(hwnd, WM_USER + 1, 0, (LPARAM)msg);
}

void InstallSelectedVersion() {
    int sel = (int)SendMessage(hComboBox, CB_GETCURSEL, 0, 0);
    if (sel >= 0 && sel < 4) {
        const wchar_t* url = urls[sel];
        const wchar_t* file = files[sel];

        UpdateStatus(hStatusText, L"Downloading installer...");

        std::thread([url, file]() {
            if (DownloadFile(url, file)) {
                UpdateStatus(hStatusText, L"Running installer...");
                RunFileAndDeleteAfter(file);
                UpdateStatus(hStatusText, L"Installation complete.");
            }
            else {
                UpdateStatus(hStatusText, L"Download failed.");
            }
            }).detach();
    }
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        hFont = CreateFont(
            -16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
            L"Segoe UI"
        );

        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        int winWidth = rcClient.right - rcClient.left;

        int comboWidth = 200;
        int comboHeight = 150;
        int buttonWidth = 120;
        int buttonHeight = 30;
        int centerX = winWidth / 2;

        hComboBox = CreateWindow(L"COMBOBOX", NULL,
            CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL,
            centerX - comboWidth / 2, 100, comboWidth, comboHeight,
            hwnd, NULL, NULL, NULL);

        SendMessage(hComboBox, WM_SETFONT, (WPARAM)hFont, TRUE);

        for (int i = 0; i < 4; i++) {
            SendMessage(hComboBox, CB_ADDSTRING, 0, (LPARAM)versions[i]);
        }

        SendMessage(hComboBox, CB_SETCURSEL, 0, 0);

        HWND hButton = CreateWindow(L"BUTTON", L"Install",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            centerX - buttonWidth / 2, 160, buttonWidth, buttonHeight,
            hwnd, (HMENU)1, NULL, NULL);

        SendMessage(hButton, WM_SETFONT, (WPARAM)hFont, TRUE);

        hStatusText = CreateWindow(
            L"STATIC", L"Odaberi verziju i stisni install.",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            30, 230, winWidth - 60, 20,
            hwnd, NULL, NULL, NULL);

        SendMessage(hStatusText, WM_SETFONT, (WPARAM)hFont, TRUE);

        break;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == 1) {
            InstallSelectedVersion();
        }
        break;
    case WM_USER + 1:
        SetWindowText(hStatusText, (LPCWSTR)lParam);
        break;
    case WM_DESTROY:
        DeleteObject(hFont);
        DeleteObject(hBackgroundBrush);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int nCmdShow) {
    const wchar_t CLASS_NAME[] = L"DarioInstaliraOfficeWin";

    hBackgroundBrush = CreateSolidBrush(RGB(144, 164, 183));

    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInst;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = hBackgroundBrush;

    RegisterClass(&wc);

    int winWidth = 525;
    int winHeight = 450;

    int screenWidth = GetSystemMetrics(SM_CXSCREEN);
    int screenHeight = GetSystemMetrics(SM_CYSCREEN);
    int x = (screenWidth - winWidth) / 2;
    int y = (screenHeight - winHeight) / 2;

    HWND hwnd = CreateWindowEx(0, CLASS_NAME, L"Dario instalira office",
        WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX & ~WS_THICKFRAME,
        x, y, winWidth, winHeight,
        NULL, NULL, hInst, NULL);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}

