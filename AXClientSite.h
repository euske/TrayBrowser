// -*- mode: c++ -*-
//  AXClientSite.h
// 
#include <Windows.h>
#include <OleIdl.h>

//  AXClientSite
//
class AXClientSite :
    public IDispatch,
    public IOleClientSite,
    public IOleInPlaceSite,
    public IOleInPlaceFrame
{
private:
    int _refCount;
    HWND _hWnd;

public:
    AXClientSite(HWND hWnd);

    // IUnknown methods
    STDMETHODIMP QueryInterface(REFIID iid, void**ppvObject);
    STDMETHODIMP_(ULONG) AddRef()
        { _refCount++; return _refCount; }
    STDMETHODIMP_(ULONG) Release()
        { _refCount--; return _refCount; }
    
    // IDispatch methods
    STDMETHODIMP GetTypeInfoCount(UINT*)
        { return E_NOTIMPL; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**)
        { return E_NOTIMPL; }
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*)
        { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*)
        { return E_NOTIMPL; }
    
    // IOleClientSite methods
    STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** ppmk)
        { *ppmk = NULL; return E_NOTIMPL; }
    STDMETHODIMP GetContainer(IOleContainer** ppc)
        { *ppc = NULL; return E_NOINTERFACE; }
    STDMETHODIMP SaveObject()
        { return S_OK; }
    STDMETHODIMP ShowObject()
        { return S_OK; }
    STDMETHODIMP RequestNewObjectLayout()
        { return S_OK; }
    STDMETHODIMP OnShowWindow(BOOL fShow);
    
    // IOleInPlaceSite methods
    STDMETHODIMP CanInPlaceActivate()
        { return S_OK; }
    STDMETHODIMP OnInPlaceActivate()
        { return S_OK; }
    STDMETHODIMP OnInPlaceDeactivate()
        { return S_OK; }
    STDMETHODIMP OnUIActivate()
        { return S_OK; }
    STDMETHODIMP OnUIDeactivate(int)
        { return S_OK; }
    STDMETHODIMP DiscardUndoState()
        { return S_OK; }
    STDMETHODIMP DeactivateAndUndo()
        { return S_OK; }
    STDMETHODIMP OnPosRectChange(LPCRECT)
        { return S_OK; }
    STDMETHODIMP Scroll(SIZE s)
        { return E_NOTIMPL; }
    STDMETHODIMP GetWindowContext(
        IOleInPlaceFrame** ppFrame,
        IOleInPlaceUIWindow** ppDoc,
        LPRECT lprcPosRect,
        LPRECT lprcClipRect,
        LPOLEINPLACEFRAMEINFO lpFrameInfo);

    // IOleWindow methods
    STDMETHODIMP GetWindow(HWND* phwnd)
        { *phwnd = _hWnd; return S_OK; }
    STDMETHODIMP ContextSensitiveHelp(BOOL)
        { return E_NOTIMPL; }

    // IOleInPlaceUIWindow methods
    STDMETHODIMP GetBorder(LPRECT lprectBorder)
        { GetClientRect(_hWnd, lprectBorder); return S_OK; }
    STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS)
        { return E_NOTIMPL; }
    STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS)
        { return S_OK; }
    STDMETHODIMP SetActiveObject(IOleInPlaceActiveObject*, LPCOLESTR)
        { return S_OK; }

    // IOleInPlaceFrame methods
    STDMETHODIMP EnableModeless(BOOL)
        { return E_NOTIMPL; }
    STDMETHODIMP TranslateAccelerator(LPMSG, WORD)
        { return E_NOTIMPL; }
    STDMETHODIMP SetStatusText(LPCOLESTR)
        { return E_NOTIMPL; }
    STDMETHODIMP InsertMenus(HMENU, LPOLEMENUGROUPWIDTHS)
        { return E_NOTIMPL; }
    STDMETHODIMP RemoveMenus(HMENU)
        { return E_NOTIMPL; }
    STDMETHODIMP SetMenu(HMENU, HOLEMENU, HWND)
        { return E_NOTIMPL; }
};
