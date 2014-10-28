//  TrayBrowser.cpp
//

#include <stdio.h>
#include <stdlib.h>
#include <Windows.h>
#include <StrSafe.h>
#include <ExDisp.h>
#include <OleAuto.h>
#include <Mshtmhst.h>
#include "Resource.h"
#include "AXClientSite.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "ole32.lib")

const LPCWSTR TRAYBROWSER_NAME = L"TrayBrowser";
const LPCWSTR TRAYBROWSER_WNDCLASS = L"TrayBrowserClass";
const LPCWSTR TASKBAR_CREATED = L"TaskbarCreated";
const UINT MAX_URL_CHARS = 1024;
const UINT WM_NOTIFY_ICON = WM_USER+1;
static UINT WM_TASKBAR_CREATED;
static FILE* logfp = NULL;      // logging


static INT_PTR CALLBACK showTextInputDialogProc(
    HWND hWnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
    switch (uMsg) {
    case WM_INITDIALOG:
        SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)lParam);
        {
            WCHAR* text = (WCHAR*)lParam;
            if (text != NULL) {
                SetDlgItemText(hWnd, IDC_EDIT_URL, text);
            }
        }
        return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDOK:
            {
                WCHAR* text = (WCHAR*)GetWindowLongPtr(hWnd, GWLP_USERDATA);
                if (text != NULL) {
                    UINT n = GetDlgItemText(hWnd, IDC_EDIT_URL, text, MAX_URL_CHARS-1);
                    text[n] = L'\0';
                }
                EndDialog(hWnd, IDOK);
            }
            break;
        case IDCANCEL:
            EndDialog(hWnd, IDCANCEL);
            break;
        }
        return TRUE;
    default:
        return FALSE;
    }
}

static WCHAR* showTextInputDialog(HWND hWnd, WCHAR* src)
{
    WCHAR* text = (WCHAR*)calloc(MAX_URL_CHARS, sizeof(WCHAR));
    StringCchCopy(text, MAX_URL_CHARS, src);
    LPARAM result = DialogBoxParam(
        NULL, MAKEINTRESOURCE(IDD_TEXTINPUT),
        hWnd, showTextInputDialogProc, (LPARAM)text);
    if (result != IDOK) {
        free(text);
        return NULL;
    }
    return text;
}


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
    void openURL();

public:
    TrayBrowser(int id);
    void initialize(HWND hWnd, RECT* rect);
    void unInitialize();
    void registerIcon();
    void unregisterIcon();
    void resize(RECT* rect);
    void handleIconUI(LPARAM lParam, POINT pt);
    void doCommand(WPARAM wParam);
};

TrayBrowser::TrayBrowser(int id)
{
    _iconId = id;
    HMENU menu = LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDM_POPUPMENU));
    if (menu != NULL) {
        _hMenu = GetSubMenu(menu, 0);
        if (_hMenu != NULL) {
            SetMenuDefaultItem(_hMenu, IDM_OPEN, FALSE);
        }
    }
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

void TrayBrowser::handleIconUI(LPARAM lParam, POINT pt)
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

void TrayBrowser::doCommand(WPARAM wParam)
{
    switch (LOWORD(wParam)) {
    case IDM_EXIT:
        DestroyWindow(_hWnd);
        break;
    case IDM_OPEN:
        openURL();
        break;
    case IDM_PIN:
        break;
    }
}

void TrayBrowser::openURL()
{
    if (_browser2 != NULL) {
        BSTR bstrSrc = NULL;
        _browser2->get_LocationURL(&bstrSrc);
        if (bstrSrc != NULL) {
            WCHAR* url = showTextInputDialog(_hWnd, (WCHAR*)bstrSrc);;
            if (url != NULL) {
                if (logfp) fwprintf(logfp, L" openURL: url=%s\n", url);
                BSTR bstrUrl = SysAllocString(url);
                if (bstrUrl != NULL) {
                    _browser2->Navigate(bstrUrl, NULL, NULL, NULL, NULL);
                    SysFreeString(bstrUrl);
                }
                free(url);
            }
            SysFreeString(bstrSrc);
        }
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
        if (self != NULL) {
            self->doCommand(wParam);
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
                self->handleIconUI(lParam, pt);
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
    
    OleInitialize(0);

    // Create a TrayBrowser object.
    TrayBrowser* browser = new TrayBrowser(1);
    
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
