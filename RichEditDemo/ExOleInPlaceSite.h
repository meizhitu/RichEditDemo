#pragma once

class IExOleInPlaceSiteImpl : public  IOleInPlaceSite 
{
public :
	STDMETHOD(GetWindow)( HWND * phwnd )
	{
		return S_OK;
	}

	STDMETHOD(ContextSensitiveHelp)( BOOL fEnterMode ) {
		return E_NOTIMPL;
	}

	STDMETHOD(CanInPlaceActivate)(void)	{
		return S_FALSE;
	}

	STDMETHOD(OnInPlaceActivate)(void) {
		m_bInPlaceActive = true;
		return S_OK;
	}

	STDMETHOD(OnUIActivate)(void) {
		return E_NOTIMPL;
	}

	STDMETHOD(GetWindowContext)(IOleInPlaceFrame **ppFrame, IOleInPlaceUIWindow **ppDoc,
		LPRECT lprcPosRect, LPRECT lprcClipRect,
		LPOLEINPLACEFRAMEINFO lpFrameInfo)	
	{

		return S_OK;
	}

	STDMETHOD(Scroll)(SIZE scrollExtant)	{
		return E_NOTIMPL;
	}

	STDMETHOD(OnUIDeactivate)(BOOL fUndoable) {
		return E_NOTIMPL;
	}

	STDMETHOD(OnInPlaceDeactivate)( void) {
		m_bInPlaceActive = false;
		return S_OK;
	}

	STDMETHOD(DiscardUndoState)( void) {
		return E_NOTIMPL;
	}

	STDMETHOD(DeactivateAndUndo)( void)	{
		return E_NOTIMPL;
	}

	STDMETHOD(OnPosRectChange)(LPCRECT lprcPosRect)	{
		return E_NOTIMPL;
	}

	STDMETHOD(QueryInterface) (REFIID iid, LPVOID FAR * ppvObject)
	{
		HRESULT hr = S_OK;
		*ppvObject = NULL;

		if ( iid == IID_IUnknown ||
			iid == IID_IRichEditOleCallback )
		{
			*ppvObject = this;
			AddRef();
			hr = NOERROR;
		}
		else
		{
			hr = E_NOINTERFACE;
		}

		return hr;
	}

	STDMETHOD_(ULONG,AddRef) (void)
	{
		return ++m_dwRef;
	}

	STDMETHOD_(ULONG,Release) (void)
	{
		if ( --m_dwRef == 0 )
		{
			delete this;
			return 0;
		}

		return m_dwRef;
	}

	DWORD m_dwRef;

	bool m_bInPlaceActive;
};
