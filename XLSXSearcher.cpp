#include <windows.h>
#include "XLSXUtils.h"
#include"resource.h"
#include <libxl.h>
#include <ShObjIdl.h>
#include <Shlwapi.h>


#define BUTTON_IN 39101
#define ID_EDIT 39201
#define BUTTON_OUT 39102
#define ID_OUT 39202
#define BUTTON_DONE 39103
#define EDIT_NAME 39203
#define EDIT_ID 39204
#define IDC_PROGRESS1 39301

const char g_szClassName[] = "myWindowClass";
LPWSTR g_lpstrFilePATH = nullptr;
LPWSTR g_lpstrOutPATH = nullptr;
std::vector<std::string> files;
std::vector<std::string> XMLfiles;
std::vector<std::string> PDFfiles;
std::string uName;
std::string uID;

HANDLE hConsole = 0;
HWND hEdit;
HWND hEdit2;
HWND hName;
HWND hID;
HWND hWndProgress;
HWND hWndButtonDone;

std::string LPWSTRToString(LPWSTR pwszStr) {
    int nLen = WideCharToMultiByte(CP_ACP, 0, pwszStr, -1, NULL, 0, NULL, NULL);
    if (nLen <= 0) {
        return "";
    }

    char* pszStr = new char[nLen];
    WideCharToMultiByte(CP_ACP, 0, pwszStr, -1, pszStr, nLen, NULL, NULL);

    std::string str(pszStr);
    delete[] pszStr;

    return str;
}

BOOL CALLBACK EnumChildProc(HWND hWnd, LPARAM lParam) {
    SendMessage(hWnd, WM_SETFONT, lParam, TRUE);
    return TRUE;
}

BOOL CheckValidPath(HWND hWnd) {
    WCHAR szPath1[MAX_PATH];
    WCHAR szPath2[MAX_PATH];
    GetWindowTextW(hEdit, szPath1, MAX_PATH);
    GetWindowTextW(hEdit2, szPath2, MAX_PATH);
    if (szPath1 == L"" && szPath2 == L"") {
        MessageBox(hWnd, "Please Enter Path!", "Path Validation", MB_OK | MB_ICONWARNING);
        return 0;
    }
    if (!(PathFileExistsW(szPath1))) {
        MessageBox(hWnd, "Invalid Path!", "Path Validation", MB_OK|MB_ICONWARNING);
        return 0;
    }
    return 1;
}

// Function to open a folder dialog and return the selected folder path
HRESULT OpenFolderDialog(HWND hwndOwner, LPWSTR lpstrFolder, DWORD nMaxFolder)
{
    IFileDialog* pfd = NULL;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
    if (SUCCEEDED(hr))
    {
        // Set the options on the dialog to select folders.
        DWORD dwFlags;
        hr = pfd->GetOptions(&dwFlags);
        if (SUCCEEDED(hr))
        {
            hr = pfd->SetOptions(dwFlags | FOS_PICKFOLDERS);
            if (SUCCEEDED(hr))
            {
                // Show the dialog
                hr = pfd->Show(hwndOwner);
                if (SUCCEEDED(hr))
                {
                    // Obtain the result once the user clicks the 'Open' button.
                    IShellItem* psiResult;
                    hr = pfd->GetResult(&psiResult);
                    if (SUCCEEDED(hr))
                    {
                        PWSTR pszFolderPath = NULL;
                        hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFolderPath);
                        if (SUCCEEDED(hr))
                        {
                            // Copy the folder path to the buffer passed in
                            wcsncpy_s(lpstrFolder, nMaxFolder, pszFolderPath, nMaxFolder);
                            CoTaskMemFree(pszFolderPath);
                        }
                        psiResult->Release();
                    }
                }
            }
        }
        pfd->Release();
    }
    return hr;
}

HRESULT OpenFileDlg(HWND hWndOwner, LPWSTR lpstrFile, DWORD nMaxFile) {
    // Create a new instance of IFileDialog
    IFileDialog* pfd = NULL;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
    if (SUCCEEDED(hr))
    {
        // Set the options on the dialog.
        DWORD dwFlags;
        hr = pfd->GetOptions(&dwFlags);
        if (SUCCEEDED(hr))
        {
            // Add the FOS_FORCEFILESYSTEM flag to ensure it's a file system item.
            hr = pfd->SetOptions(dwFlags | FOS_FORCEFILESYSTEM);
            if (SUCCEEDED(hr))
            {
                // Show the dialog
                hr = pfd->Show(hWndOwner);
                if (SUCCEEDED(hr))
                {
                    // Obtain the result once the user clicks the 'Open' button.
                    // The result is an IShellItem object.
                    IShellItem* psiResult;
                    hr = pfd->GetResult(&psiResult);
                    if (SUCCEEDED(hr))
                    {
                        // We can now get the file path from the IShellItem object.
                        PWSTR pszFilePath = NULL;
                        hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
                        if (SUCCEEDED(hr))
                        {
                            // Copy the file path to the buffer passed in
                            wcsncpy_s(lpstrFile, nMaxFile, pszFilePath, nMaxFile);
                            CoTaskMemFree(pszFilePath);
                        }
                        psiResult->Release();
                    }
                }
            }
        }
        pfd->Release();
    }
    return hr;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_COMMAND:
    {
        if (LOWORD(wParam) == EDIT_NAME && HIWORD(wParam) == EN_KILLFOCUS) {
            TCHAR szText[100]; 
            GetWindowText(hName, szText, 100);
            WriteConsole(hConsole, szText, sizeof(szText), NULL, NULL);
            uName = szText;
        }
        if (LOWORD(wParam) == EDIT_ID && HIWORD(wParam) == EN_KILLFOCUS) {
            TCHAR szText[100];
            GetWindowText(hID, szText, 100);
            WriteConsole(hConsole, szText, sizeof(szText), NULL, NULL);
            uID = szText;
        }
        if (LOWORD(wParam) == BUTTON_IN && HIWORD(wParam) == BN_CLICKED) {
            g_lpstrFilePATH = new WCHAR[MAX_PATH];
            HRESULT hr = OpenFolderDialog(hWnd, g_lpstrFilePATH, MAX_PATH);
            if (SUCCEEDED(hr)) {
                SetWindowTextW(hEdit, g_lpstrFilePATH);
            } 
        }
        if (LOWORD(wParam) == BUTTON_OUT && HIWORD(wParam) == BN_CLICKED) {
            g_lpstrOutPATH = new WCHAR[MAX_PATH];
            HRESULT hr = OpenFolderDialog(hWnd, g_lpstrOutPATH, MAX_PATH);
            if (SUCCEEDED(hr)) {
                SetWindowTextW(hEdit2, g_lpstrOutPATH);
            }
        }
        if (LOWORD(wParam) == BUTTON_DONE && HIWORD(wParam) == BN_CLICKED) {
            if (CheckValidPath(hWnd)) {
                EnableWindow(hWndButtonDone, FALSE);
                WriteConsole(hConsole, uName.c_str(), static_cast<DWORD>(uName.size()), NULL, NULL);
                XLSXUtils::GetAllFiles(g_lpstrFilePATH, files, XMLfiles, PDFfiles); 
                SendMessage(hWndProgress, PBM_SETPOS, 5, 0);
                XLSXUtils::DoContain(uID, uName, files, XMLfiles, hWndProgress);
                /*for each (std::string path in files)
                {

                    WriteConsole(hConsole, path.c_str(), static_cast<DWORD>(path.size()), NULL, NULL);
                }
                for each (std::string path in XMLfiles)
                {
                    WriteConsole(hConsole, path.c_str(), static_cast<DWORD>(path.size()), NULL, NULL);
                }*/
                std::string outputPath = LPWSTRToString(g_lpstrOutPATH);
                XLSXUtils::Export(outputPath, files, XMLfiles, hWndProgress);
            }
            else {
                
            }
        }
        break;
    }
    case WM_CREATE:
    {
        HFONT hFont = CreateFont(
            -MulDiv(10, GetDeviceCaps(GetDC(NULL), LOGPIXELSY), 72), // Height of font
            0,                    // Average character width
            0,                    // Angle of escapement
            0,                    // Base-line orientation angle
            FW_NORMAL,            // Font weight
            FALSE,                // Italic attribute option
            FALSE,                // Underline attribute option
            FALSE,                // Strikeout attribute option
            DEFAULT_CHARSET,      // Character set identifier
            OUT_DEFAULT_PRECIS,   // Output precision
            CLIP_DEFAULT_PRECIS,  // Clipping precision
            DEFAULT_QUALITY,      // Output quality
            DEFAULT_PITCH | FF_SWISS, // Pitch and family
            TEXT("Microsoft YaHei")); // Typeface name

        HWND hWndButton = CreateWindow("BUTTON", "IN", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 10, 10, 40, 29, hWnd, (HMENU)BUTTON_IN, (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE), NULL);

        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "",WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL, 60, 10, 290, 29, hWnd, (HMENU)ID_EDIT, GetModuleHandle(NULL), NULL);

        HWND hWndButton2 = CreateWindow("BUTTON", "OUT", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 10, 40, 40, 29, hWnd, (HMENU)BUTTON_OUT, (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE), NULL);

        hEdit2 = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 60, 40, 290, 29, hWnd, (HMENU)ID_OUT, GetModuleHandle(NULL), NULL);

        hName = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "����", WS_CHILD | WS_VISIBLE , 10, 80, 140, 28, hWnd, (HMENU)EDIT_NAME, GetModuleHandle(NULL), NULL);

        hID = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "ѧ��", WS_CHILD | WS_VISIBLE , 160, 80, 140, 28, hWnd, (HMENU)EDIT_ID, GetModuleHandle(NULL), NULL);

        hWndButtonDone = CreateWindow("BUTTON", "OK", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 300, 80, 40, 28, hWnd, (HMENU)BUTTON_DONE, (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE), NULL);

        hWndProgress = CreateWindowEx(0, PROGRESS_CLASS, "", WS_CHILD | WS_VISIBLE,
            10, 120, 340, 20, hWnd, (HMENU)IDC_PROGRESS1, NULL, NULL);

        SetWindowRgn(hWndButton, CreateRoundRectRgn(0, 0, 100, 25, 10, 10), TRUE);
        SetWindowRgn(hWndButton2, CreateRoundRectRgn(0, 0, 100, 25, 10, 10), TRUE);

        HDC  hWndButtonIn = GetDC(hWndButton);
        HDC  hWndButtonOut = GetDC(hWndButton2);
        SetBkColor(hWndButtonIn, RGB(235, 235, 235));
        SetBkColor(hWndButtonOut, RGB(235, 235, 235));

        SetTextColor(hWndButtonIn, RGB(0, 0, 0));
        SetTextColor(hWndButtonOut, RGB(0, 0, 0));

        SetFocus(hName);
        SetFocus(hID);

        SendMessage(hWndProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

        EnumChildWindows(hWnd, EnumChildProc, (LPARAM)hFont);
        EnableWindow(hEdit, FALSE);
        EnableWindow(hEdit2, FALSE);
        break;
    }
    case WM_CLOSE:
        DestroyWindow(hWnd);
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;   
    }
    return DefWindowProc(hWnd, msg, wParam, lParam);
}

int CALLBACK WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{

    WNDCLASSEX wc;
    HWND hwnd;
    MSG Msg;

    AllocConsole();
    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = 0;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = g_szClassName;
    wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wc))
    {
        MessageBox(NULL, "Window Registration Failed!", "Error!",
            MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    char szText[256] = {0};
    LoadString(hInstance, IDS_TITLE, szText, 256);

    // Step 2: Creating the Window
    hwnd = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        g_szClassName,
        szText,
        WS_OVERLAPPEDWINDOW & ~ WS_THICKFRAME,
        CW_USEDEFAULT, CW_USEDEFAULT, 390, 200,
        NULL, NULL, hInstance, NULL);

    if (hwnd == NULL)
    {
        MessageBox(NULL, "Window Creation Failed!", "Error!",
            MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Step 3: The Message Loop
    while (GetMessage(&Msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&Msg);
        DispatchMessage(&Msg);
    }
    return Msg.wParam;
}