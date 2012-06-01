#include "stdafx.h"
#include "ExRichEdit.h"
#include "..\DynamicOle\DynamicOle_i.h"
#include <comdefsp.h>
#include "ExRichEditData.h"

const IID IID_ITextServices = {
	// 由于此 IID 在 SDK 中的 riched20.lib 是错误的 , 需重新定义
	0x8d33f740, 0xcf58, 0x11ce, {0xa8, 0x9d, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}

};

#define ID_EVENT_UPDATE_OLE 1                       //更新OLE对象的TIMER的ID
#define ELAPSE_TIME_UPDATE_OLE 100                  //更新OLE对象的时间间隔100ms

CExRichEdit::CExRichEdit()
{
	
}

CExRichEdit::~CExRichEdit()
{
	if (m_hWnd && ::IsWindow(m_hWnd))
		KillTimer(ID_EVENT_UPDATE_OLE);
}

void CExRichEdit::StartTimer()
{
	SetTimer(ID_EVENT_UPDATE_OLE,ELAPSE_TIME_UPDATE_OLE);
}

LRESULT CExRichEdit::OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (ID_EVENT_UPDATE_OLE == static_cast<int>(wParam))
	{
		if (m_hWnd && ::IsWindow(m_hWnd))
			::RedrawWindow(m_hWnd,NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
		return TRUE;
	}
	return FALSE;
}

LRESULT CExRichEdit::OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	bHandled = TRUE;
	return 1L;
}

LRESULT CExRichEdit::OnPaint( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CRect clientRect(0, 0, 0, 0);
	GetClientRect(&clientRect);
	m_memDC.Create(clientRect.Width(), clientRect.Height());
	if (m_memDC.IsReady())
	{
		PAINTSTRUCT ps;
		HDC hClientDC = BeginPaint(&ps);
		HBRUSH hbrush = ::CreateSolidBrush(RGB(255, 255, 255));
		HBRUSH holdBrush = HBRUSH(::SelectObject(m_memDC.GetSafeHdc(), hbrush));
		::Rectangle(m_memDC.GetSafeHdc(), -1, -1, clientRect.Width() + 1, clientRect.Height() + 1);
		::SelectObject(m_memDC.GetSafeHdc(), holdBrush);
		::DeleteObject(hbrush);
		
		HRESULT hr = CRichEditCtrl::GetOleInterface()->QueryInterface(IID_ITextServices, (void**)&m_textServices);
		m_textServices->TxDraw(DVASPECT_CONTENT, NULL, NULL, NULL, m_memDC.GetSafeHdc(), NULL, LPCRECTL(&clientRect), NULL, NULL, NULL, NULL, NULL);

		::BitBlt(hClientDC, 0, 0, clientRect.Width(), clientRect.Height(), m_memDC.GetSafeHdc(), 0, 0, SRCCOPY);

		EndPaint(&ps);
		bHandled = TRUE;
		return 0L;
	}
	bHandled = FALSE;
	return 1L;
}

LRESULT CExRichEdit::OnLButtonDblClk(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	WORD selType = CRichEditCtrl::GetSelectionType();
	bHandled = FALSE;
	if (selType == SEL_OBJECT)
	{
		REOBJECT reObj;
		SecureZeroMemory(&reObj, sizeof(REOBJECT));
		reObj.cbStruct = sizeof(REOBJECT);

		CHARRANGE cr;
		CRichEditCtrl::GetSel(cr);

		reObj.cp = cr.cpMin;

		IRichEditOle* pRichEditOle = GetOleInterface();
		HRESULT hr = pRichEditOle->GetObject(REO_IOB_USE_CP, &reObj, REO_GETOBJ_NO_INTERFACES);
		if (S_OK == GetScode(hr))
		{
			CExRichEditData *pdata = (CExRichEditData*)reObj.dwUser;
			if (pdata->m_dataType == GIF)
			{
				//bHandled = TRUE;
			}
		}
	}
	
	return 0L;
}

void CExRichEdit::InsertGif(LONG gif)
{
	CComQIPtr<IDynamicOleCom> spDyn;
	HRESULT hr = spDyn.CoCreateInstance(STR_PROGID);
	if(SUCCEEDED(hr))
	{
		LPOLEOBJECT lpOleObject = NULL;
		HRESULT hr = spDyn->QueryInterface(IID_IOleObject, (void**)&lpOleObject);

		IUnknownPtr lpUnk = lpOleObject;
		hr = lpUnk->QueryInterface(IID_IOleObject, (LPVOID*)&lpOleObject);
		if (lpOleObject == NULL)
			throw(E_OUTOFMEMORY);
		//hr = lpOleObject->SetClientSite( static_cast<IOleClientSite *>( this ) );
		IViewObject2Ptr lpViewObject;// IViewObject for IOleObject above
		hr = lpOleObject->QueryInterface(IID_IViewObject2, (LPVOID*)&lpViewObject);
		if (hr != S_OK)
		{
			AtlThrow(hr);
		}
		IRichEditOle* pRichEditOle = GetOleInterface();
		////获取RichEdit的OLEClientSite
		IOleClientSitePtr lpClientSite;
		hr = pRichEditOle->GetClientSite(&lpClientSite);

		if (hr != S_OK)
		{
			AtlThrow(hr);
		}
		REOBJECT reobject;
		ZeroMemory(&reobject,sizeof(REOBJECT));
		reobject.cbStruct = sizeof(REOBJECT);

		CLSID clsid;
		hr = lpOleObject->GetUserClassID(&clsid);
		if (hr != S_OK)
		{
			AtlThrow(hr);
		}

		reobject.clsid = clsid;
		reobject.cp = -1;
		//reobject.cp = REO_CP_SELECTION;
		reobject.dvaspect = DVASPECT_CONTENT;//DVASPECT_OPAQUE;
		reobject.poleobj = lpOleObject;
		reobject.polesite = lpClientSite;
		//reobject.pstg = lpStorage;
		SIZEL sizel;
		sizel.cx = sizel.cy = 0; // let richedit determine initial size

		Image* img = (Image*)gif;
		SIZEL sizeInPix = {img->GetWidth(), img->GetHeight()};
		SIZEL sizeInHiMetric;
		AtlPixelToHiMetric(&sizeInPix, &sizeInHiMetric);

		reobject.sizel = sizeInHiMetric;
		reobject.dwFlags = REO_BELOWBASELINE|REO_STATIC;//REO_RESIZABLE

		CExRichEditData* pdata = new CExRichEditData;
		pdata->m_dataType = GIF;
		reobject.dwUser = (DWORD)pdata;//TODO 用户数据

		lpOleObject->SetClientSite(lpClientSite); 
		hr = pRichEditOle->InsertObject(&reobject);
		lpOleObject->SetExtent(DVASPECT_CONTENT, &sizeInHiMetric);
		OleSetContainedObject(lpOleObject, TRUE);
		lpOleObject->Release();
		spDyn->SetHostWindow((LONG)m_hWnd);
		spDyn->InsertGif(gif);
	}
}

void CExRichEdit::InsertBitmap(CString& filePath)
{
	//创建输入数据源
	LPSTORAGE lpStorage;
	//分配内存
	LPLOCKBYTES lpLockBytes = NULL;
	SCODE hr = ::CreateILockBytesOnHGlobal(NULL,TRUE,&lpLockBytes);
	if(hr != S_OK)
		AtlThrow(hr);
	ATLASSERT(lpLockBytes != NULL);
	hr = ::StgCreateDocfileOnILockBytes(lpLockBytes,STGM_SHARE_EXCLUSIVE|STGM_CREATE|STGM_READWRITE,0,&lpStorage);
	if(hr != S_OK)
	{
		ATLVERIFY(lpLockBytes->Release() == 0);
		lpLockBytes = NULL;
		AtlThrow(hr);
	}
	ATLASSERT(lpStorage != NULL);
	LPOLEOBJECT lpOleObject;
	hr = ::OleCreateFromFile(CLSID_NULL, filePath,
		IID_IUnknown, OLERENDER_DRAW, NULL, NULL, 
		lpStorage, (void **)&lpOleObject);
	if (hr != S_OK)
	{
		AtlThrow(hr);
	}
	ATLASSERT(lpOleObject != NULL);
	IUnknownPtr lpUnk = lpOleObject;
	hr = lpUnk->QueryInterface(IID_IOleObject, (LPVOID*)&lpOleObject);
	if (lpOleObject == NULL)
		throw(E_OUTOFMEMORY);
	IViewObject2Ptr lpViewObject;// IViewObject for IOleObject above
	hr = lpOleObject->QueryInterface(IID_IViewObject2, (LPVOID*)&lpViewObject);
	if (hr != S_OK)
	{
		AtlThrow(hr);
	}
	IRichEditOle* pRichEditOle = GetOleInterface();
	////获取RichEdit的OLEClientSite
	IOleClientSitePtr lpClientSite;
	hr = pRichEditOle->GetClientSite(&lpClientSite);
	if (hr != S_OK)
	{
		AtlThrow(hr);
	}
	REOBJECT reobject;
	ZeroMemory(&reobject,sizeof(REOBJECT));
	reobject.cbStruct = sizeof(REOBJECT);

	CLSID clsid;
	hr = lpOleObject->GetUserClassID(&clsid);
	if (hr != S_OK)
	{
		AtlThrow(hr);
	}

	reobject.clsid = clsid;
	reobject.cp = -1;
	//reobject.cp = REO_CP_SELECTION;
	reobject.dvaspect = DVASPECT_CONTENT;
	reobject.poleobj = lpOleObject;
	reobject.polesite = lpClientSite;
	//reobject.pstg = lpStorage;
	SIZEL sizel;
	sizel.cx = sizel.cy = 0; // let richedit determine initial size

	SIZEL sizeInPix = {20, 20};
	SIZEL sizeInHiMetric;
	AtlPixelToHiMetric(&sizeInPix, &sizeInHiMetric);

	//reobject.sizel = sizeInHiMetric;
	reobject.dwFlags = REO_BELOWBASELINE;//REO_RESIZABLE
	CExRichEditData* pdata = new CExRichEditData;
	pdata->m_dataType = IMAGE;
	reobject.dwUser = (DWORD)pdata;//TODO 用户数据


	lpOleObject->SetClientSite(lpClientSite); 
	hr = pRichEditOle->InsertObject(&reobject);
	OleSetContainedObject(lpOleObject, TRUE);
	lpOleObject->Release();
}