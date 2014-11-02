//  TrayBrowser.cpp
//

#include <stdio.h>
#include <stdlib.h>
#include <Windows.h>
#include <StrSafe.h>
#include <ExDisp.h>
#include <OleAuto.h>
#include <Mshtmhst.h>
#include <Shlwapi.h>
#include "Resource.h"
#include "AXClientSite.h"
#include "ini.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")

#define MAX_URL_LENGTH 256
const LPCWSTR TRAYBROWSER_NAME = L"TrayBrowser";
const LPCWSTR TRAYBROWSER_INI = L"TrayBrowser.ini";
const LPCWSTR TRAYBROWSER_WNDCLASS = L"TrayBrowserClass";
const LPCWSTR TASKBAR_CREATED = L"TaskbarCreated";
const UINT WM_NOTIFY_ICON = WM_USER+1;
static UINT WM_TASKBAR_CREATED;
static FILE* logfp = NULL;      // logging

static BOOL getMenuItemChecked(
    HMENU menu,
    UINT item)
{
    MENUITEMINFO mii = {0};
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STATE;
    GetMenuItemInfo(menu, item, FALSE, &mii);
    return (mii.fState & MFS_CHECKED);
}

static void setMenuItemChecked(
    HMENU menu,
    UINT item,
    BOOL checked)
{
    MENUITEMINFO mii = {0};
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STATE;
    GetMenuItemInfo(menu, item, FALSE, &mii);
    if (checked) {
        mii.fState |= MFS_CHECKED;
    } else {
        mii.fState &= ~MFS_CHECKED;
    }
    SetMenuItemInfo(menu, item, FALSE, &mii);
}

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
            WCHAR** pText = (WCHAR**)lParam;
            if (pText != NULL && *pText != NULL) {
                SetDlgItemText(hWnd, IDC_EDIT_URL, *pText);
            }
        }
        return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDOK:
            {
                WCHAR** pText = (WCHAR**)GetWindowLongPtr(hWnd, GWLP_USERDATA);
                if (pText != NULL) {
                    HWND item = GetDlgItem(hWnd, IDC_EDIT_URL);
                    if (item != NULL) {
                        int n = GetWindowTextLength(item);
                        WCHAR* text = (WCHAR*) calloc(n+1, sizeof(WCHAR));
                        if (text != NULL) {
                            GetWindowText(item, text, n+1);
                            *pText = text;
                        }
                    }
                }
                EndDialog(hWnd, IDOK);
            }
            break;
        case IDCANCEL:
            {
                WCHAR** pText = (WCHAR**)GetWindowLongPtr(hWnd, GWLP_USERDATA);
                if (pText != NULL) {
                    *pText = NULL;
                }
            }
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
    WCHAR* text = src;
    DialogBoxParam(
        NULL, MAKEINTRESOURCE(IDD_TEXTINPUT),
        hWnd, showTextInputDialogProc, (LPARAM)&text);
    return text;
}

//  trayBrowserINIHandler
//   Called for each config line.
static int trayBrowserINIHandler(
    void* user, const char* section,
    const char* name, const char* value)
{
    fprintf(stderr, " section=%s, name=%s, value=%s\n", section, name, value);

    wchar_t wvalue[MAX_URL_LENGTH];
    mbstowcs_s(NULL, wvalue, _countof(wvalue), value, _TRUNCATE);
    HMENU menu = (HMENU)user;
    UINT uid = IDM_RECENT + GetMenuItemCount(menu);
    AppendMenu(menu, MF_STRING, uid, wvalue);
    return 1;
}


//  TrayBrowser
//
class TrayBrowser
{
    int _iconId;
    HWND _hWnd;
    HMENU _hMenu;
    HMENU _hBookmarks;
    BOOL _modal;
    
    AXClientSite* _site;
    IStorage* _storage;
    IOleObject* _ole;
    IWebBrowser2* _browser2;

    void openURL(const WCHAR* url);
    void openURLDialog();
    void togglePin();
    void toggleShow();

public:
    TrayBrowser(int id);
    void loadIni(const WCHAR* path);
    void saveIni(const WCHAR* path);
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
    _modal = FALSE;
    HMENU menu = LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDM_POPUPMENU));
    if (menu != NULL) {
        _hMenu = GetSubMenu(menu, 0);
        if (_hMenu != NULL) {
            SetMenuDefaultItem(_hMenu, IDM_OPEN, FALSE);
            _hBookmarks = GetSubMenu(_hMenu, 2);
            DeleteMenu(_hBookmarks, IDM_RECENT, MF_BYCOMMAND);
        }
    }
}

void TrayBrowser::loadIni(const WCHAR* path)
{
    if (logfp) fwprintf(logfp, L" loadIni: path=%s\n", path);

    // Open the .ini file.
    {
        FILE* fp = NULL;
        if (_wfopen_s(&fp, path, L"r") == 0) {
            ini_parse_file(fp, trayBrowserINIHandler, _hBookmarks);
            fclose(fp);
        }
    }
}

void TrayBrowser::saveIni(const WCHAR* path)
{
    if (logfp) fwprintf(logfp, L" saveIni: path=%s\n", path);

    // Open the .ini file.
    if (_hBookmarks != NULL) {
        FILE* fp = NULL;
        if (_wfopen_s(&fp, path, L"w") == 0) {
            fwprintf(fp, L"[bookmarks]\n");
            for (int i = 0; i < GetMenuItemCount(_hBookmarks); i++) {
                WCHAR value[MAX_URL_LENGTH];
                GetMenuString(_hBookmarks, i, value, _countof(value),
                              MF_BYPOSITION);
                fwprintf(fp, L"url%d = %s\n", i, value);
            }
            fwprintf(fp, L"\n");
            fclose(fp);
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
        SendMessage(_hWnd, WM_COMMAND, MAKEWPARAM(IDM_SHOW, 1), NULL);
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
    UINT uid = LOWORD(wParam);
    switch (uid) {
    case IDM_EXIT:
        DestroyWindow(_hWnd);
        break;
    case IDM_OPEN:
        openURLDialog();
        break;
    case IDM_PIN:
        togglePin();
        break;
    case IDM_SHOW:
        toggleShow();
        break;
    default:
        if (IDM_RECENT <= uid) {
            WCHAR value[MAX_URL_LENGTH];
            if (GetMenuString(_hBookmarks, uid, value, _countof(value),
                              MF_BYCOMMAND)) {
                openURL(value);
            }
        }
        break;
    }
}

void TrayBrowser::openURL(const WCHAR* url)
{
    if (_browser2 != NULL) {
        BSTR bstrUrl = SysAllocString(url);
        if (bstrUrl != NULL) {
            _browser2->Navigate(bstrUrl, NULL, NULL, NULL, NULL);
            SysFreeString(bstrUrl);
        }
    }
 }

void TrayBrowser::openURLDialog()
{
    if (_modal) return;
    
    _modal = TRUE;
    if (_browser2 != NULL) {
        BSTR bstrSrc = NULL;
        _browser2->get_LocationURL(&bstrSrc);
        if (bstrSrc != NULL) {
            WCHAR* url = showTextInputDialog(_hWnd, (WCHAR*)bstrSrc);;
            if (url != NULL) {
                if (logfp) fwprintf(logfp, L" openURL: url=%s\n", url);
                openURL(url);
                free(url);
            }
            SysFreeString(bstrSrc);
        }
    }
    _modal = FALSE;
}

void TrayBrowser::togglePin()
{
    BOOL pinned = getMenuItemChecked(_hMenu, IDM_PIN);
    pinned = !pinned;
    if (logfp) fwprintf(logfp, L" togglePin: %d\n", pinned);
    setMenuItemChecked(_hMenu, IDM_PIN, pinned);
    
    HWND hwndAfter = (pinned)? HWND_TOPMOST : HWND_NOTOPMOST;
    SetWindowPos(_hWnd, hwndAfter, 0,0,0,0, 
                 (SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE));
    
    DWORD exStyle = GetWindowLongPtr(_hWnd, GWL_EXSTYLE);
    if (pinned) {
        exStyle |= WS_EX_NOACTIVATE;
    } else {
        exStyle &= ~WS_EX_NOACTIVATE;
    }
    SetWindowLongPtr(_hWnd, GWL_EXSTYLE, exStyle);
}

void TrayBrowser::toggleShow()
{
    if (_modal) return;

    if (IsWindowVisible(_hWnd)) {
        ShowWindow(_hWnd, SW_HIDE);
    } else {
        ShowWindow(_hWnd, SW_SHOWNORMAL);
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
	WNDCLASS klass = {0};
	klass.lpfnWndProc = trayBrowserWndProc;
	klass.hInstance = hInstance;
	klass.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_TRAYBROWSER));
	klass.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
        klass.lpszClassName = TRAYBROWSER_WNDCLASS;
	atom = RegisterClass(&klass);
    }

    // Register the window message.
    WM_TASKBAR_CREATED = RegisterWindowMessage(TASKBAR_CREATED);
    
    OleInitialize(0);

    // Create a TrayBrowser object.
    TrayBrowser* browser = new TrayBrowser(1);

    // Load a .ini file.
    WCHAR path[MAX_PATH];
    GetModuleFileName(hInstance, path, _countof(path));
    PathRemoveFileSpec(path);
    PathAppend(path, TRAYBROWSER_INI);
    browser->loadIni(path);
    
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
    browser->saveIni(path);
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
    //_wfopen_s(&logfp, L"log.txt", L"a");
    return TrayBrowserMain(hInstance, hPrevInstance, nCmdShow, argc, argv);
}
#else
int wmain(int argc, wchar_t* argv[])
{
    logfp = stderr;
    return TrayBrowserMain(GetModuleHandle(NULL), NULL, SW_SHOWDEFAULT, argc, argv);
}
#endif
