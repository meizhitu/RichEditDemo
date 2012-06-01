// DynamicOleCom.h : Declaration of the CDynamicOleCom

#pragma once
#include "resource.h"       // main symbols

#include "DynamicOle_i.h"
#include <vector>
using namespace std;
#include <gdiplus.h>
using namespace Gdiplus;

#pragma comment(lib,"gdiplus.lib")

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif
class CDynamicOleCom;

#define WM_OLE_REFRESHMESSAGE WM_USER+100

class CWinHiddenMock : 
	public CWindowImpl<CWinHiddenMock, CWindow, CNullTraits>
{
	BEGIN_MSG_MAP(CWinHiddenMock)
		MESSAGE_HANDLER(WM_OLE_REFRESHMESSAGE, OnRefreshOle)
	END_MSG_MAP()
private:
	CDynamicOleCom*   m_dynamicOleCom;

public:
	CWinHiddenMock() 
	{
		
	}
	~CWinHiddenMock()
	{
		if (m_hWnd != NULL)
		{
			::DestroyWindow(m_hWnd);
			m_hWnd = NULL;
		}
	}

public:
	LRESULT OnRefreshOle(UINT uMsg, WPARAM wParam, 
		LPARAM lParam, BOOL& bHandled);
	void AttachOle(CDynamicOleCom* ctl);
	BOOL PostRefreshOle()
	{	
		if (::IsWindow(m_hWnd))
			return ::SendMessage(m_hWnd,WM_OLE_REFRESHMESSAGE,0,0);
		return FALSE;
	}

};


struct TGifFrame
{
	int m_FrameIndex;
	int m_lpause;
	int m_FrameCount;
};
// CDynamicOleCom

class ATL_NO_VTABLE CDynamicOleCom :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CDynamicOleCom, &CLSID_DynamicOleCom>,
	public IDispatchImpl<IDynamicOleCom, &IID_IDynamicOleCom, &LIBID_DynamicOleLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IOleControlImpl<CDynamicOleCom>,
	public IOleObjectImpl<CDynamicOleCom>,
	public IOleInPlaceActiveObjectImpl<CDynamicOleCom>,
	public IOleInPlaceObjectWindowlessImpl<CDynamicOleCom>,
	public IViewObjectExImpl<CDynamicOleCom>,
	public CComControl<CDynamicOleCom>
	
{
public:
	CDynamicOleCom()
	{
		m_currentFrame = 0;
		m_mockWindow.AttachOle(this);
		if (m_mockWindow.m_hWnd == NULL)
		{
			RECT rt = {0, 0, 0, 0};
			m_mockWindow.Create(NULL, rt);
		}
		Status sta = GdiplusStartup(&m_gdiplusToken, &m_gdiplusStartupInput, NULL);		//GDI+≥ı ºªØ
	}
	~CDynamicOleCom()
	{
		if (m_objectType == 2)//GIF
		{
			delete m_gifImg;
		}
		GdiplusShutdown(m_gdiplusToken);
	}

	DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |  
	OLEMISC_CANTLINKINSIDE |  
		OLEMISC_INSIDEOUT |  
		OLEMISC_ACTIVATEWHENVISIBLE |  
		OLEMISC_SETCLIENTSITEFIRST  
		)  

DECLARE_REGISTRY_RESOURCEID(IDR_DYNAMICOLECOM)


BEGIN_COM_MAP(CDynamicOleCom)
	COM_INTERFACE_ENTRY(IDynamicOleCom)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IViewObjectEx)  
	COM_INTERFACE_ENTRY(IViewObject2)  
	COM_INTERFACE_ENTRY(IViewObject)  
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)  
	COM_INTERFACE_ENTRY(IOleInPlaceObject)  
	COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)  
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)  
	COM_INTERFACE_ENTRY(IOleControl)  
	COM_INTERFACE_ENTRY(IOleObject)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	STDMETHOD(DoVerb)(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
	{
		return S_OK;
	}

	STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid,
		LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult,
		EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		switch (dispidMember)
		{
		case DISPID_AMBIENT_SHOWHATCHING:
			{
				V_VT(pvarResult) = VT_BOOL;
				V_BOOL(pvarResult) = 0;
				return S_OK;
			}
			break;
		}
		return _tih.Invoke((IDispatch*)this, dispidMember, riid, lcid,
			wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr);
	}

	STDMETHOD(CanInPlaceActivate)(void)	{
		return S_FALSE;
	}
	
public:
	HRESULT OnDraw(ATL_DRAWINFO& di);
	STDMETHOD (InsertGif)(LONG img);
	STDMETHOD (SetHostWindow)(LONG hWnd);
protected:
	GdiplusStartupInput		m_gdiplusStartupInput;
	ULONG_PTR				m_gdiplusToken;
private:
	int m_objectType;
	CWinHiddenMock m_mockWindow;
	HWND m_hostWnd;
	void RedrawHostWindow();
	//GIF
	Image *m_gifImg;
	int m_currentFrame;
	static VOID CALLBACK OnGifTimer( HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime );
	vector<TGifFrame> m_vGif;
	void LoadGif(Image* gif);
	
};

OBJECT_ENTRY_AUTO(__uuidof(DynamicOleCom), CDynamicOleCom)
