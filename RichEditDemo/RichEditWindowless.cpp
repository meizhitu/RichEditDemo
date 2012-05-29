#include "stdafx.h"
#include "RichEditWindowless.h"
#include "..\DynamicOle\DynamicOle_i.h"
#include "ExRichEditData.h"

CExRichEditWindowless::CExRichEditWindowless()
{
}

CExRichEditWindowless::~CExRichEditWindowless()
{

}

IRichEditOle* CExRichEditWindowless::GetIRichEditOle()
{
	IRichEditOle *pRichItem = NULL;
	GetTextServices()->TxSendMessage(EM_GETOLEINTERFACE, 0, (LPARAM)&pRichItem, 0);
	return pRichItem;
}

BOOL CExRichEditWindowless::SetOLECallback(IRichEditOleCallback* pCallback)
{
	GetTextServices()->TxSendMessage(EM_SETOLECALLBACK, 0, (LPARAM)pCallback, 0);
	return TRUE;
}

int CExRichEditWindowless::InsertText(long nInsertAfterChar, LPCTSTR lpstrText, bool bCanUndo)
{
	int nRet = SetSel(nInsertAfterChar, nInsertAfterChar);
	ReplaceSel(lpstrText, bCanUndo);
	return nRet;
}

int CExRichEditWindowless::AppendText(LPCTSTR lpstrText, bool bCanUndo)
{
	int nRet = SetSel(-1, -1);
	ReplaceSel(lpstrText, bCanUndo);
	return nRet;
}

int CExRichEditWindowless::SetSel(long nStartChar, long nEndChar)
{
	CHARRANGE cr;
	cr.cpMin = nStartChar;
	cr.cpMax = nEndChar;
	LRESULT lResult;
	GetTextServices()->TxSendMessage(EM_EXSETSEL, 0, (LPARAM)&cr, &lResult); 
	return (int)lResult;
}

void CExRichEditWindowless::ReplaceSel(LPCTSTR lpszNewText, bool bCanUndo)
{
	GetTextServices()->TxSendMessage(EM_REPLACESEL, (WPARAM) bCanUndo, (LPARAM)lpszNewText, 0); 
}

void CExRichEditWindowless::InsertGif(LONG gif)
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
		IRichEditOle* pRichEditOle = GetIRichEditOle();
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
		spDyn->SetHostWindow((LONG)hwndParent);
		spDyn->InsertGif(gif);
	}
}


