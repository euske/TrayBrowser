//  AXClientSite.cpp
// 
#include <Windows.h>
#include <ExDisp.h>
#include <OleIdl.h>
#include "AXClientSite.h"


//  AXClientSite
//
AXClientSite::AXClientSite(HWND hWnd)
{
    _refCount = 0;
    _hWnd = hWnd;
}

// IUnknown methods
STDMETHODIMP AXClientSite::QueryInterface(REFIID iid, void** ppvObject)
{
    if (iid == IID_IUnknown) {
        *ppvObject = this;
    } else if (iid == IID_IDispatch) {
        *ppvObject = (IDispatch*)this;
    } else if (iid == IID_IAdviseSink) {
        *ppvObject = (IAdviseSink*)this;
    } else if (iid == IID_IOleClientSite) {
        *ppvObject = (IOleClientSite*)this;
    } else if (iid == IID_IOleInPlaceSite) {
        *ppvObject = (IOleInPlaceSite*)this;
    } else if (iid == IID_IOleInPlaceFrame) {
        *ppvObject = (IOleInPlaceFrame*)this;
    } else if (iid == IID_IOleInPlaceUIWindow) {
        *ppvObject = (IOleInPlaceUIWindow*)this;
    } else {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

// IOleClientSite methods
STDMETHODIMP AXClientSite::OnShowWindow(BOOL f)
{
    InvalidateRect(_hWnd, 0, TRUE);
    return S_OK;
}

// IOleInPlaceSite methods
STDMETHODIMP AXClientSite::GetWindowContext(
    IOleInPlaceFrame** ppFrame,
    IOleInPlaceUIWindow** ppDoc,
    LPRECT lprcPosRect,
    LPRECT lprcClipRect,
    LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
    *ppFrame = (IOleInPlaceFrame*)this;
    *ppDoc = NULL;
    GetClientRect(_hWnd, lprcPosRect);
    GetClientRect(_hWnd, lprcClipRect);
    
    lpFrameInfo->cb = sizeof(*lpFrameInfo);
    lpFrameInfo->fMDIApp = FALSE;
    lpFrameInfo->hwndFrame = _hWnd;
    lpFrameInfo->haccel = NULL;
    lpFrameInfo->cAccelEntries = 0;
    return S_OK;
}
