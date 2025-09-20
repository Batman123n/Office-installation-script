#define UNICODE
#define _UNICODE
#include <windows.h>
#include <thread>
#include <shellapi.h>
#include <wininet.h>
#include <cstdio>

#pragma comment(lib, "wininet.lib")

const wchar_t* versions[] = {
    L"Office 2019 Pro Plus HR",
    L"Office 2021 Pro Plus HR",
    L"Office 2024 Pro Plus HR",
    L"Office 365 Enterprise HR"
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
HFONT hFont;
HBRUSH hBackgroundBrush = NULL;

bool DownloadFile(const wchar_t* url, const wchar_t* filename) {
    HINTERNET hInternet = InternetOpen(L"MyInstaller", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);
    if (!hInternet) return false;

    HINTERNET hFile = InternetOpenUrl(hInternet, url, NULL, 0, INTERNET_FLAG_RELOAD, 0);
    if (!hFile) { InternetCloseHandle(hInternet); return false; }

    FILE* out = _wfopen(filename, L"wb");
    if (!out) { InternetCloseHandle(hFile); InternetCloseHandle(hInternet); return false; }

    char buffer[4096];
    DWORD bytesRead;
    while (InternetReadFile(hFile, buffer, sizeof(buffer), &bytesRead) && bytesRead > 0)
        fwrite(buffer, 1, bytesRead, out);

    fclose(out);
    InternetCloseHandle(hFile);
    InternetCloseHandle(hInternet);
    return true;
}

void RunFileAndDelete(const wchar_t* filename) {
    SHELLEXECUTEINFO sei = { sizeof(sei) };
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

void InstallSelected() {
    int sel = (int)SendMessage(hComboBox, CB_GETCURSEL, 0, 0);
    if (sel >= 0 && sel < 4) {
        const wchar_t* url = urls[sel];
        const wchar_t* file = files[sel];

        std::thread([url, file]() {
            if (DownloadFile(url, file)) {
                RunFileAndDelete(file);
            }
        }).detach();
    }
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        hFont = CreateFont(-16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                           DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                           CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

        RECT rc; GetClientRect(hwnd, &rc);
        int cx = (rc.right - rc.left) / 2;

        hComboBox = CreateWindow(L"COMBOBOX", NULL,
            CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL,
            cx - 100, 100, 200, 150, hwnd, NULL, NULL, NULL);
        SendMessage(hComboBox, WM_SETFONT, (WPARAM)hFont, TRUE);
        for (int i = 0; i < 4; i++) SendMessage(hComboBox, CB_ADDSTRING, 0, (LPARAM)versions[i]);
        SendMessage(hComboBox, CB_SETCURSEL, 0, 0);

        HWND hButton = CreateWindow(L"BUTTON", L"Instaliraj",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            cx - 60, 160, 120, 30, hwnd, (HMENU)1, NULL, NULL);
        SendMessage(hButton, WM_SETFONT, (WPARAM)hFont, TRUE);
        break;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == 1) InstallSelected();
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
    const wchar_t CLASS_NAME[] = L"Dario";
    hBackgroundBrush = CreateSolidBrush(RGB(144, 164, 183));

    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInst;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = hBackgroundBrush;
    RegisterClass(&wc);

    int w = 525, h = 450;
    int x = (GetSystemMetrics(SM_CXSCREEN) - w) / 2;
    int y = (GetSystemMetrics(SM_CYSCREEN) - h) / 2;

    HWND hwnd = CreateWindowEx(0, CLASS_NAME, L"Dario instalira office",
        WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX & ~WS_THICKFRAME,
        x, y, w, h, NULL, NULL, hInst, NULL);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}

