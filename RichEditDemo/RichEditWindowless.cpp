#include "stdafx.h"
#include "RichEditWindowless.h"
#include "..\DynamicOle\DynamicOle_i.h"
#include "ExRichEditData.h"
#include "dragdrop.h"

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

CString CExRichEditWindowless::GetTextRange(long nStartChar, long nEndChar)
{
	TEXTRANGEW tr = { 0 };
	tr.chrg.cpMin = nStartChar;
	tr.chrg.cpMax = nEndChar;
	LPWSTR lpText = NULL;
	lpText = new WCHAR[nEndChar - nStartChar + 1];
	::ZeroMemory(lpText, (nEndChar - nStartChar + 1) * sizeof(WCHAR));
	tr.lpstrText = lpText;
	GetTextServices()->TxSendMessage(EM_GETTEXTRANGE, 0, (LPARAM)&tr, 0);
	CString sText(lpText);
	delete[] lpText;
	return sText;
}

STDMETHODIMP CExRichEditWindowless::QueryAcceptData(LPDATAOBJECT lpdataobj, CLIPFORMAT FAR *lpcfFormat,DWORD reco, BOOL fReally, HGLOBAL hMetaPict)
{
	if ( NULL == lpdataobj )
		return E_INVALIDARG;	

	HRESULT hr;
	FORMATETC formatEtc;
	STGMEDIUM stgMedium;

	UINT uFormat = m_uOwnOleClipboardFormat;
	if (0 == uFormat)
	{
		return E_FAIL;
	}
	SecureZeroMemory(&formatEtc, sizeof(FORMATETC));
	formatEtc.cfFormat = uFormat;
	formatEtc.dwAspect = DVASPECT_CONTENT;
	formatEtc.lindex = -1;
	formatEtc.ptd = NULL;
	formatEtc.tymed = TYMED_HGLOBAL;

	hr = lpdataobj->GetData(&formatEtc, &stgMedium);
	if (S_OK == hr)
	{
		LPTSTR pDest = LPTSTR(::GlobalLock(stgMedium.hGlobal));
		if (pDest != NULL)
		{
			ReplaceSel(CString(pDest),true);
			::GlobalUnlock(stgMedium.hGlobal);
			::ReleaseStgMedium( &stgMedium );
			return S_OK;
		}
		else
		{
			::ReleaseStgMedium( &stgMedium );
			return E_FAIL;
		}
	}
	return S_OK;  

}

STDMETHODIMP CExRichEditWindowless::GetClipboardData(CHARRANGE FAR *lpchrg, DWORD reco, LPDATAOBJECT FAR *lplpdataobj)
{
	switch(reco)
	{
	case RECO_COPY:
		{
			HRESULT hr = E_NOTIMPL;
			FORMATETC formatEtc;
			STGMEDIUM stgMedium;
			CDataObject *pDataObject = new CDataObject(NULL);

			CString text = GetTextRange(lpchrg->cpMin,lpchrg->cpMax);
			CString embedding(WCH_EMBEDDING);
			text.Replace(embedding,_T("<objtct/>"));
			if(!text.IsEmpty())
			{
				UINT uFormat = m_uOwnOleClipboardFormat;
				if (0 != uFormat)
				{
					HGLOBAL hMemBlock = ::GlobalAlloc(GMEM_MOVEABLE, (text.GetLength() + 1) * sizeof(TCHAR));
					if (NULL != hMemBlock)
					{
						LPTSTR pDest = LPTSTR(::GlobalLock(hMemBlock));
						_tcscpy(pDest, LPCTSTR(text));
						::GlobalUnlock(hMemBlock);

						SecureZeroMemory(&formatEtc, sizeof(FORMATETC));
						formatEtc.cfFormat = uFormat;
						formatEtc.dwAspect = DVASPECT_CONTENT;
						formatEtc.lindex = -1;
						formatEtc.ptd = NULL;
						formatEtc.tymed = TYMED_HGLOBAL;

						SecureZeroMemory(&stgMedium, sizeof(STGMEDIUM));
						stgMedium.tymed = TYMED_HGLOBAL;
						stgMedium.hGlobal = hMemBlock;
						hr = pDataObject->SetData(&formatEtc, &stgMedium, TRUE);
						if (FAILED(hr))
						{
							::GlobalFree(hMemBlock);
							hr &= E_FAIL;
						}
					}
					else
					{
						hr &= E_OUTOFMEMORY;
					}
				}
				else
				{
					hr &= E_FAIL;
				}
			}
			else
			{
				return E_UNEXPECTED;
			}
			if (SUCCEEDED(hr))
			{
				(*lplpdataobj) = pDataObject;
				(*lplpdataobj)->AddRef();
			}
			return hr;
		}
	case RECO_CUT:
	case RECO_DRAG:
	case RECO_DROP:
	case RECO_PASTE:
	default:
		return E_NOTIMPL;
	}
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
		pdata->m_data = (void*)gif;
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


