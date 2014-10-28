//  TrayBrowser.cpp
//

#include <stdio.h>
#include <stdlib.h>
#include <Windows.h>
#include <StrSafe.h>
#include <ExDisp.h>
#include <Mshtmhst.h>
#include "Resource.h"
#include "AXClientSite.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")

const LPCWSTR TRAYBROWSER_NAME = L"TrayBrowser";
const LPCWSTR TRAYBROWSER_WNDCLASS = L"TrayBrowserClass";
const LPCWSTR TASKBAR_CREATED = L"TaskbarCreated";
static UINT WM_TASKBAR_CREATED;
const UINT WM_NOTIFY_ICON = WM_USER+1;
// logging
static FILE* logfp = NULL;


//  TrayBrowser
//
class TrayBrowser
{
    int _iconId;
    HWND _hWnd;
    HMENU _hMenu;
    
    AXClientSite* _site;
    IStorage* _storage;
    IOleObject* _ole;
    IWebBrowser2* _browser2;

public:
    TrayBrowser(int id, HMENU menu);
    void initialize(HWND hWnd, RECT* rect);
    void unInitialize();
    void registerIcon();
    void unregisterIcon();
    void resize(RECT* rect);
    void handleUI(LPARAM lParam, POINT pt);
};

TrayBrowser::TrayBrowser(int id, HMENU menu)
{
    _iconId = id;
    _hMenu = menu;
}

void TrayBrowser::initialize(HWND hWnd, RECT* rect)
{
    HRESULT hr;

    _hWnd = hWnd;
    _site = new AXClientSite(hWnd);
    if (logfp) fwprintf(logfp, L" site=%p\n", _site);

    hr = StgCreateDocfile(
        NULL,
        (STGM_READWRITE | STGM_SHARE_EXCLUSIVE |
         STGM_CREATE | STGM_DIRECT),
        0, &_storage);
    if (logfp) fwprintf(logfp, L" storage=%p\n", _storage);

    if (_site != NULL && _storage != NULL) {
        hr = OleCreate(
            CLSID_WebBrowser,
            IID_IOleObject,
            OLERENDER_DRAW,
            0,
            _site,
            _storage,
            (void**)&_ole);
        if (logfp) fwprintf(logfp, L" ole=%p\n", _ole);
    }

    if (_ole != NULL) {
        hr = _ole->DoVerb(
            OLEIVERB_INPLACEACTIVATE,
            NULL,
            _site,
            0, hWnd, rect);
    
        hr = _ole->QueryInterface(
            IID_IWebBrowser2,
            (void**)&_browser2);
        if (logfp) fwprintf(logfp, L" browser2=%p\n", _browser2);
    }

    if (_browser2 != NULL) {
        _browser2->Navigate(L"http://www.yahoo.co.jp/", 0,0,0,0);
    }
}

void TrayBrowser::unInitialize()
{
    if (_browser2) {
        _browser2->Quit();
        _browser2->Release();
        _browser2 = NULL;
    }

    if (_ole != NULL) {
        IOleInPlaceObject* iib = NULL;
        if (SUCCEEDED(_ole->QueryInterface(
                          IID_IOleInPlaceObject,
                          (void**)&iib))) {
            iib->UIDeactivate();
            iib->InPlaceDeactivate();
            iib->Release();
        }
        _ole->Release();
        _ole = NULL;
    }

    if (_storage != NULL) {
        _storage->Release();
        _storage = NULL;
    }

    if (_site != NULL) {
        _site->Release();
        delete _site;
    }
};

void TrayBrowser::registerIcon()
{
    NOTIFYICONDATA nidata = {0};
    nidata.cbSize = sizeof(nidata);
    nidata.hWnd = _hWnd;
    nidata.uID = _iconId;
    nidata.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nidata.uCallbackMessage = WM_NOTIFY_ICON;
    nidata.hIcon = (HICON)GetClassLongPtr(_hWnd, GCLP_HICON);
    StringCchPrintf(nidata.szTip, _countof(nidata.szTip), 
                    L"%s", TRAYBROWSER_NAME);
    Shell_NotifyIcon(NIM_ADD, &nidata);
}

void TrayBrowser::unregisterIcon()
{
    NOTIFYICONDATA nidata = {0};
    nidata.cbSize = sizeof(nidata);
    nidata.hWnd = _hWnd;
    nidata.uID = _iconId;
    Shell_NotifyIcon(NIM_DELETE, &nidata);
}

void TrayBrowser::resize(RECT* rect)
{
    if (_ole != NULL) {
        IOleInPlaceObject* iib = NULL;
        if (SUCCEEDED(_ole->QueryInterface(
                          IID_IOleInPlaceObject,
                          (void**)&iib))) {
            if (logfp) fwprintf(logfp, L" resize\n");
            iib->SetObjectRects(rect, rect);
            iib->Release();
        }
    }
}

void TrayBrowser::handleUI(LPARAM lParam, POINT pt)
{
    switch (lParam) {
    case WM_LBUTTONDBLCLK:
        if (_hMenu != NULL) {
            UINT item = GetMenuDefaultItem(_hMenu, FALSE, 0);
            SendMessage(_hWnd, WM_COMMAND, MAKEWPARAM(item, 1), NULL);
        }
        break;
        
    case WM_LBUTTONUP:
        if (IsWindowVisible(_hWnd)) {
            ShowWindow(_hWnd, SW_HIDE);
        } else {
            ShowWindow(_hWnd, SW_SHOWNORMAL);
            SetForegroundWindow(_hWnd);
        }
        break;
        
    case WM_RBUTTONUP:
        if (_hMenu != NULL) {
            TrackPopupMenu(_hMenu, TPM_LEFTALIGN, 
                           pt.x, pt.y, 0, _hWnd, NULL);
        }
        PostMessage(_hWnd, WM_NULL, 0, 0);
        break;
    }
}


//  trayBrowserWndProc
//
static LRESULT CALLBACK trayBrowserWndProc(
    HWND hWnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
    //fwprintf(logfp, L"msg: %x, hWnd=%p, wParam=%p\n", uMsg, hWnd, wParam);

    switch (uMsg) {
    case WM_CREATE:
    {
        // Initialization.
	CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
	TrayBrowser* self = (TrayBrowser*)(cs->lpCreateParams);
        if (self != NULL) {
            RECT rect;
	    SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)self);
            GetClientRect(hWnd, &rect);
            self->initialize(hWnd, &rect);
            self->registerIcon();
        }
	return FALSE;
    }
    
    case WM_DESTROY:
    {
        // Clean up.
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	TrayBrowser* self = (TrayBrowser*)lp;
        if (self != NULL) {
            self->unregisterIcon();
            self->unInitialize();
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
	    DestroyWindow(hWnd);
	    break;
	}
	return FALSE;
    }

    case WM_SIZE:
    {
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	TrayBrowser* self = (TrayBrowser*)lp;
        if (self != NULL) {
            RECT rect;
            GetClientRect(hWnd, &rect);
            self->resize(&rect);
        }
        return FALSE;
    }
        
    case WM_CLOSE:
	ShowWindow(hWnd, SW_HIDE);
	return FALSE;

    case WM_NOTIFY_ICON:
    {
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	TrayBrowser* self = (TrayBrowser*)lp;
        if (self != NULL) {
            // UI event handling.
            POINT pt;
            if (GetCursorPos(&pt)) {
                self->handleUI(lParam, pt);
            }
        }
        return FALSE;
    }

    default:
        if (uMsg == WM_TASKBAR_CREATED) {
            LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
            TrayBrowser* self = (TrayBrowser*)lp;
            if (self != NULL) {
                self->registerIcon();
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
        klass.lpszClassName = TRAYBROWSER_WNDCLASS;
	atom = RegisterClass(&klass);
    }

    // Register the window message.
    WM_TASKBAR_CREATED = RegisterWindowMessage(TASKBAR_CREATED);
    
    // Initialize the menu.
    HMENU menu = LoadMenu(hInstance, MAKEINTRESOURCE(IDM_POPUPMENU));
    if (menu != NULL) {
        menu = GetSubMenu(menu, 0);
        if (menu != NULL) {
            SetMenuDefaultItem(menu, IDM_OPEN, FALSE);
        }
    }
    
    OleInitialize(0);

    // Create a TrayBrowser object.
    TrayBrowser* browser = new TrayBrowser(1, menu);
    
    // Create a main window.
    HWND hWnd = CreateWindow(
	(LPCWSTR)atom,
	TRAYBROWSER_NAME,
	WS_OVERLAPPEDWINDOW,
	CW_USEDEFAULT, CW_USEDEFAULT,
	CW_USEDEFAULT, CW_USEDEFAULT,
	NULL, NULL, hInstance, browser);
    ShowWindow(hWnd, nCmdShow);

    // Event loop.
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Clean up.
    delete browser;

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
    _wfopen_s(&logfp, L"log.txt", L"a");
    return TrayBrowserMain(hInstance, hPrevInstance, nCmdShow, argc, argv);
}
#else
int wmain(int argc, wchar_t* argv[])
{
    logfp = stderr;
    return TrayBrowserMain(GetModuleHandle(NULL), NULL, SW_SHOWDEFAULT, argc, argv);
}
#endif
