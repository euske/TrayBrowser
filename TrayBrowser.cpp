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
const LPCSTR INI_SECTION_BOOKMARKS = "bookmarks";
static UINT WM_TASKBAR_CREATED;
static FILE* logfp = NULL;      // logging

// getMenuItemChecked: return TRUE if the item is checked.
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

// setMenuItemChecked: check/uncheck a menu item.
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

// showTextInputDialogProc: dialog proc.
static INT_PTR CALLBACK showTextInputDialogProc(
    HWND hWnd,
    UINT uMsg,
    WPARAM wParam,
    LPARAM lParam)
{
    switch (uMsg) {
    case WM_INITDIALOG:
        // Setting up.
        SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)lParam);
        {
            WCHAR** pText = (WCHAR**)lParam;
            if (pText != NULL && *pText != NULL) {
                SetDlgItemText(hWnd, IDC_EDIT_URL, *pText);
            }
        }
        return TRUE;
        
    case WM_COMMAND:
        // Respond to button press.
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

// showTextInputDialog: prompt with a text field.
static WCHAR* showTextInputDialog(HWND hWnd, WCHAR* src)
{
    WCHAR* text = src;
    DialogBoxParam(
        NULL, MAKEINTRESOURCE(IDD_TEXTINPUT),
        hWnd, showTextInputDialogProc, (LPARAM)&text);
    return text;
}

// trayBrowserINIHandler
//   Called for each config line.
static int trayBrowserINIHandler(
    void* user, const char* section,
    const char* name, const char* value)
{
    fprintf(stderr, " section=%s, name=%s, value=%s\n", section, name, value);

    if (stricmp(section, INI_SECTION_BOOKMARKS) == 0) {
        wchar_t wvalue[MAX_URL_LENGTH];
        mbstowcs_s(NULL, wvalue, _countof(wvalue), value, _TRUNCATE);
        HMENU menu = (HMENU)user;
        UINT uid = IDM_RECENT + GetMenuItemCount(menu);
        AppendMenu(menu, MF_STRING, uid, wvalue);
    }
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
    void pinWindow(BOOL pinned);
    void showWindow();

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

// TrayBrowser(id):
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

// loadIni(path): load a .ini file.
void TrayBrowser::loadIni(const WCHAR* path)
{
    if (logfp) fwprintf(logfp, L"loadIni: path=%s\n", path);

    if (_hBookmarks != NULL) {
        FILE* fp = NULL;
        if (_wfopen_s(&fp, path, L"r") == 0) {
            ini_parse_file(fp, trayBrowserINIHandler, _hBookmarks);
            fclose(fp);
        }
    }
}

// saveIni(path): save a .ini file.
void TrayBrowser::saveIni(const WCHAR* path)
{
    if (logfp) fwprintf(logfp, L"saveIni: path=%s\n", path);

    if (_hBookmarks != NULL) {
        FILE* fp = NULL;
        if (_wfopen_s(&fp, path, L"w") == 0) {
            // Write bookmark items.
            fprintf(fp, "[%s]\n", INI_SECTION_BOOKMARKS);
            for (int i = 0; i < GetMenuItemCount(_hBookmarks); i++) {
                WCHAR wvalue[MAX_URL_LENGTH];
                char value[MAX_URL_LENGTH*2];
                GetMenuString(_hBookmarks, i, wvalue, _countof(wvalue),
                              MF_BYPOSITION);
                wcstombs_s(NULL, value, _countof(value), wvalue, _TRUNCATE);
                fprintf(fp, "url%d = %s\n", i, value);
            }
            fprintf(fp, "\n");
            fclose(fp);
        }
    }
}

// initialize: set up a browser object.
void TrayBrowser::initialize(HWND hWnd, RECT* rect)
{
    HRESULT hr;
    if (logfp) fwprintf(logfp, L"initialize\n");

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

// unInitialize: tear down a browser object.
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

// registerIcon: create a SysTray icon.
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

// unregisterIcon: destroy a SysTray icon.
void TrayBrowser::unregisterIcon()
{
    NOTIFYICONDATA nidata = {0};
    nidata.cbSize = sizeof(nidata);
    nidata.hWnd = _hWnd;
    nidata.uID = _iconId;
    Shell_NotifyIcon(NIM_DELETE, &nidata);
}

// resize: resize the window.
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

// handleIconUI: UI handing.
void TrayBrowser::handleIconUI(LPARAM lParam, POINT pt)
{
    switch (lParam) {
    case WM_LBUTTONDBLCLK:
        // Double click - choose the default item.
        if (_hMenu != NULL) {
            UINT item = GetMenuDefaultItem(_hMenu, FALSE, 0);
            SendMessage(_hWnd, WM_COMMAND, MAKEWPARAM(item, 1), NULL);
        }
        break;
        
    case WM_LBUTTONUP:
        // Single click - show the window.
        SendMessage(_hWnd, WM_COMMAND, MAKEWPARAM(IDM_SHOW, 1), NULL);
        break;
        
    case WM_RBUTTONUP:
        // Right click - open a popup menu.
        SetForegroundWindow(_hWnd);
        if (_hMenu != NULL) {
            TrackPopupMenu(_hMenu,
                           TPM_LEFTALIGN | TPM_RIGHTBUTTON, 
                           pt.x, pt.y, 0, _hWnd, NULL);
        }
        PostMessage(_hWnd, WM_NULL, 0, 0);
        break;
    }
}

// doCommand: respond to a menu choice.
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
        pinWindow(!getMenuItemChecked(_hMenu, IDM_PIN));
        break;
        
    case IDM_SHOW:
        showWindow();
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

// openURLDialog(): open a url input dialog.
void TrayBrowser::openURLDialog()
{
    if (_modal) return;
    if (logfp) fwprintf(logfp, L"openURLDialog\n");
    
    _modal = TRUE;
    if (_browser2 != NULL) {
        BSTR bstrSrc = NULL;
        _browser2->get_LocationURL(&bstrSrc);
        if (bstrSrc != NULL) {
            WCHAR* url = showTextInputDialog(_hWnd, (WCHAR*)bstrSrc);;
            if (url != NULL) {
                openURL(url);
                free(url);
            }
            SysFreeString(bstrSrc);
        }
    }
    _modal = FALSE;
}

// openURL(url): navigate to URL.
void TrayBrowser::openURL(const WCHAR* url)
{
    if (logfp) fwprintf(logfp, L"openURL: url=%s\n", url);
    
    if (_browser2 != NULL) {
        BSTR bstrUrl = SysAllocString(url);
        if (bstrUrl != NULL) {
            _browser2->Navigate(bstrUrl, NULL, NULL, NULL, NULL);
            SysFreeString(bstrUrl);
        }
    }

    // Get a new item ID.
    UINT maxuid = IDM_RECENT;
    for (int i = 0; i < GetMenuItemCount(_hBookmarks); i++) {
        UINT uid = GetMenuItemID(_hBookmarks, i);
        maxuid = (maxuid < uid)? uid : maxuid;
    }
    // Remove a previous item.
    for (int i = 0; i < GetMenuItemCount(_hBookmarks); i++) {
        WCHAR value[MAX_URL_LENGTH];
        if (GetMenuString(_hBookmarks, i, value, _countof(value), MF_BYPOSITION)) {
            if (wcscmp(url, value) == 0) {
                DeleteMenu(_hBookmarks, i, MF_BYPOSITION);
                break;
            }
        }
    }
    InsertMenu(_hBookmarks, 0, MF_BYPOSITION | MF_STRING, maxuid+1, url);
 }

// pinWindow(pinned): set window pinning.
void TrayBrowser::pinWindow(BOOL pinned)
{
    if (logfp) fwprintf(logfp, L"pinWindow: pinned=%d\n", pinned);
    setMenuItemChecked(_hMenu, IDM_PIN, pinned);

    // Bring the window to the topmost.
    HWND hwndAfter = (pinned)? HWND_TOPMOST : HWND_NOTOPMOST;
    SetWindowPos(_hWnd, hwndAfter, 0,0,0,0, 
                 (SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE));

    // Change the window style.
    DWORD exStyle = GetWindowLongPtr(_hWnd, GWL_EXSTYLE);
    if (pinned) {
        exStyle |= WS_EX_TOOLWINDOW;
    } else {
        exStyle &= ~WS_EX_TOOLWINDOW;
    }
    SetWindowLongPtr(_hWnd, GWL_EXSTYLE, exStyle);
}

// showWindow(): show the window.
void TrayBrowser::showWindow()
{
    if (_modal) return;
    if (logfp) fwprintf(logfp, L"showWindow\n");

    if (!IsWindowVisible(_hWnd)) {
        ShowWindow(_hWnd, SW_SHOWNORMAL);
    }
    SetForegroundWindow(_hWnd);
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
