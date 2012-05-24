#include "stdafx.h"
#include "ExRichEdit.h"
#include "..\DynamicOle\DynamicOle_i.h"
#include <comdefsp.h>

CExRichEdit::CExRichEdit()
{
}

CExRichEdit::~CExRichEdit()
{

}

LRESULT CExRichEdit::OnLButtonDblClk(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	WORD selType = CRichEditCtrl::GetSelectionType();
	if (selType != SEL_OBJECT)
	{
		bHandled = FALSE;
	}
	else
	{
		bHandled = FALSE;
	}
	
	return 0L;
}

void CExRichEdit::InsertGif()
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
		reobject.dvaspect = DVASPECT_OPAQUE;//DVASPECT_OPAQUE;
		reobject.poleobj = lpOleObject;
		reobject.polesite = lpClientSite;
		//reobject.pstg = lpStorage;
		SIZEL sizel;
		sizel.cx = sizel.cy = 0; // let richedit determine initial size

		SIZEL sizeInPix = {20, 20};
		SIZEL sizeInHiMetric;
		AtlPixelToHiMetric(&sizeInPix, &sizeInHiMetric);

		reobject.sizel = sizeInHiMetric;
		reobject.dwFlags = REO_BELOWBASELINE;//REO_RESIZABLE
		reobject.dwUser = 0;//TODO 用户数据

		lpOleObject->SetClientSite(lpClientSite); 
		hr = pRichEditOle->InsertObject(&reobject);
		OleSetContainedObject(lpOleObject, TRUE);
		lpOleObject->Release();
		spDyn->InsertGif(0);
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
	reobject.dwUser = 0;//TODO 用户数据

	lpOleObject->SetClientSite(lpClientSite); 
	hr = pRichEditOle->InsertObject(&reobject);
	OleSetContainedObject(lpOleObject, TRUE);
	lpOleObject->Release();
}