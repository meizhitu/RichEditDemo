#pragma once

class IExOleClientSiteImpl : public IOleClientSite 
{
public:
	STDMETHOD(SaveObject)(void)	
	{
		return E_NOTIMPL;
	}

	STDMETHOD(GetMoniker)(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(GetContainer)(IOleContainer **ppContainer)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(ShowObject)(void)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(OnShowWindow)(BOOL fShow)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(RequestNewObjectLayout)(void)
	{
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
};
