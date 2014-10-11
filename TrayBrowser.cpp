//  TrayBrowser.cpp
//

#include <stdio.h>
#include <stdlib.h>
#include <Windows.h>
#include <StrSafe.h>
#include <ExDisp.h>
#include "Resource.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")

const LPCWSTR TRAYBROWSER_NAME = L"TrayBrowser";
const LPCWSTR TRAYBROWSER_WNDCLASS = L"TrayBrowserClass";
const LPCWSTR TASKBAR_CREATED = L"TaskbarCreated";
static UINT WM_TASKBAR_CREATED;
const UINT WM_NOTIFY_ICON = WM_USER+1;

class TrayBrowser
{
    int icon_id;
    IWebBrowser2* pBrowser2;

public:
    TrayBrowser(int id) : icon_id(id) { } 
    int getIconId() { return icon_id; }
    void initialize();
    void unInitialize();
    void show() { }
};

void TrayBrowser::initialize()
{
    CoCreateInstance(CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER,
                     IID_IWebBrowser2, (void**)&pBrowser2);
    if (pBrowser2 != NULL) {
        pBrowser2->put_Visible(VARIANT_TRUE);
    }
}

void TrayBrowser::unInitialize()
{
    if (pBrowser2) {
        pBrowser2->Quit();
        pBrowser2->Release();
        pBrowser2 = NULL;
    }
};

// logging
static FILE* logfp = NULL;

//  trayBrowserWndProc
//
static LRESULT CALLBACK trayBrowserWndProc(
    HWND hWnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
    //fwprintf(stderr, L"msg: %x, hWnd=%p, wParam=%p\n", uMsg, hWnd, wParam);

    switch (uMsg) {
    case WM_CREATE:
    {
        // Initialization.
	CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
	TrayBrowser* self = (TrayBrowser*)(cs->lpCreateParams);
        if (self != NULL) {
            self->initialize();
	    SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)self);
	    SendMessage(hWnd, WM_TASKBAR_CREATED, 0, 0);
        }
	return FALSE;
    }
    
    case WM_DESTROY:
    {
        // Clean up.
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	TrayBrowser* self = (TrayBrowser*)lp;
        if (self != NULL) {
            self->unInitialize();
	    // Unregister the icon.
	    NOTIFYICONDATA nidata = {0};
	    nidata.cbSize = sizeof(nidata);
	    nidata.hWnd = hWnd;
	    nidata.uID = self->getIconId();
	    Shell_NotifyIcon(NIM_DELETE, &nidata);
        }
	PostQuitMessage(0);
	return FALSE;
    }

    case WM_COMMAND:
    {
        // Command specified.
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	TrayBrowser* self = (TrayBrowser*)lp;
	switch (LOWORD(wParam)) {
	case IDM_OPEN:
	    break;
	case IDM_EXIT:
	    SendMessage(hWnd, WM_CLOSE, 0, 0);
	    break;
	}
	return FALSE;
    }

    case WM_CLOSE:
	DestroyWindow(hWnd);
	return FALSE;

    case WM_NOTIFY_ICON:
    {
        // UI event handling.
	POINT pt;
        HMENU menu = GetMenu(hWnd);
        if (menu != NULL) {
            menu = GetSubMenu(menu, 0);
        }
	switch (lParam) {
	case WM_LBUTTONDBLCLK:
            if (menu != NULL) {
                UINT item = GetMenuDefaultItem(menu, FALSE, 0);
                SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(item, 1), NULL);
            }
	    break;
	case WM_LBUTTONUP:
            {
                LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
                TrayBrowser* self = (TrayBrowser*)lp;
                if (self != NULL) {
                    self->show();
                }
            }
	    break;
	case WM_RBUTTONUP:
	    if (GetCursorPos(&pt)) {
                SetForegroundWindow(hWnd);
                if (menu != NULL) {
                    TrackPopupMenu(menu, TPM_LEFTALIGN, 
                                   pt.x, pt.y, 0, hWnd, NULL);
                }
		PostMessage(hWnd, WM_NULL, 0, 0);
	    }
	    break;
	}
	return FALSE;
    }

    default:
        if (uMsg == WM_TASKBAR_CREATED) {
            LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
            TrayBrowser* self = (TrayBrowser*)lp;
            if (self != NULL) {
                // Register the icon.
                NOTIFYICONDATA nidata = {0};
                nidata.cbSize = sizeof(nidata);
                nidata.hWnd = hWnd;
                nidata.uID = self->getIconId();
                nidata.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
                nidata.uCallbackMessage = WM_NOTIFY_ICON;
                nidata.hIcon = (HICON)GetClassLongPtr(hWnd, GCLP_HICON);
                StringCchPrintf(nidata.szTip, _countof(nidata.szTip), 
                                L"%s", TRAYBROWSER_NAME);
                Shell_NotifyIcon(NIM_ADD, &nidata);
            }
        }
	return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }
}


//  TrayBrowserMain
// 
int TrayBrowserMain(
    HINSTANCE hInstance, 
    HINSTANCE hPrevInstance, 
    int nCmdShow,
    int argc, LPWSTR* argv)
{
    // Prevent a duplicate process.
    HANDLE mutex = CreateMutex(NULL, TRUE, TRAYBROWSER_NAME);
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
	CloseHandle(mutex);
	return 0;
    }
    
    // Register the window class.
    ATOM atom;
    {
	WNDCLASS klass;
	ZeroMemory(&klass, sizeof(klass));
	klass.lpfnWndProc = trayBrowserWndProc;
	klass.hInstance = hInstance;
	//klass.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TRAYBROWSER));
        klass.hIcon = LoadIcon(NULL, MAKEINTRESOURCE(IDI_APPLICATION));
	klass.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
        klass.lpszMenuName = MAKEINTRESOURCE(IDM_POPUPMENU);
	klass.lpszClassName = TRAYBROWSER_WNDCLASS;
	atom = RegisterClass(&klass);
    }

    // Register the window message.
    WM_TASKBAR_CREATED = RegisterWindowMessage(TASKBAR_CREATED);
    
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    // Create a TrayBrowser object.
    TrayBrowser* browser = new TrayBrowser(1);
    
    // Create a SysTray window.
    HWND hWnd = CreateWindow(
	(LPCWSTR)atom,
	TRAYBROWSER_NAME,
	(WS_OVERLAPPED | WS_SYSMENU),
	CW_USEDEFAULT, CW_USEDEFAULT,
	CW_USEDEFAULT, CW_USEDEFAULT,
	NULL, NULL, hInstance, browser);
    UpdateWindow(hWnd);
    {
        // Set the default item.
        HMENU menu = GetMenu(hWnd);
        if (menu != NULL) {
            menu = GetSubMenu(menu, 0);
            if (menu != NULL) {
                SetMenuDefaultItem(menu, IDM_OPEN, FALSE);
            }
        }
    }

    // Event loop.
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Clean up.
    delete browser;

    CoUninitialize();

    return (int)msg.wParam;
}


// WinMain and wmain
#ifdef WINDOWS
int WinMain(HINSTANCE hInstance, 
	    HINSTANCE hPrevInstance, 
	    LPSTR lpCmdLine,
	    int nCmdShow)
{
    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    return TrayBrowserMain(hInstance, hPrevInstance, nCmdShow, argc, argv);
}
#else
int wmain(int argc, wchar_t* argv[])
{
    logfp = stderr;
    return TrayBrowserMain(GetModuleHandle(NULL), NULL, 0, argc, argv);
}
#endif
