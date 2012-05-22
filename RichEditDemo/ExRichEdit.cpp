#include "stdafx.h"
#include "ExRichEdit.h"

CExRichEdit::CExRichEdit()
{
}

CExRichEdit::~CExRichEdit()
{

}

void CExRichEdit::InsertBitmap(CString& filePath)
{
	//创建输入数据源
	LPSTORAGE lpStorage;
	//分配内存
	LPLOCKBYTES lpLockBytes = NULL;
	SCODE sc = ::CreateILockBytesOnHGlobal(NULL,TRUE,&lpLockBytes);
	if(sc != S_OK)
		AtlThrow(sc);
	ATLASSERT(lpLockBytes != NULL);
	sc = ::StgCreateDocfileOnILockBytes(lpLockBytes,STGM_SHARE_EXCLUSIVE|STGM_CREATE|STGM_READWRITE,0,&lpStorage);
	if(sc != S_OK)
	{
		ATLVERIFY(lpLockBytes->Release() == 0);
		lpLockBytes = NULL;
		AtlThrow(sc);
	}
	ATLASSERT(lpStorage != NULL);
	LPOLEOBJECT lpOleObject;
	sc = ::OleCreateFromFile(CLSID_NULL, filePath,
		IID_IUnknown, OLERENDER_DRAW, NULL, NULL, 
		lpStorage, (void **)&lpOleObject);
	if (sc != S_OK)
	{
		AtlThrow(sc);
	}
	ATLASSERT(lpOleObject != NULL);
	LPUNKNOWN lpUnk = lpOleObject;
	lpUnk->QueryInterface(IID_IOleObject, (LPVOID*)&lpOleObject);
	lpUnk->Release();
	if (lpOleObject == NULL)
		throw(E_OUTOFMEMORY);
	LPVIEWOBJECT2 lpViewObject;// IViewObject for IOleObject above
	sc = lpOleObject->QueryInterface(IID_IViewObject2, (LPVOID*)&lpViewObject);
	if (sc != S_OK)
	{
		AtlThrow(sc);
	}
	OleSetContainedObject(lpOleObject, TRUE);
	IRichEditOle* pRichEditOle = GetOleInterface();
	////获取RichEdit的OLEClientSite
	LPOLECLIENTSITE lpClientSite;
	sc = pRichEditOle->GetClientSite(&lpClientSite);
	if (sc != S_OK)
	{
		AtlThrow(sc);
	}
	REOBJECT reobject;
	ZeroMemory(&reobject,sizeof(REOBJECT));
	reobject.cbStruct = sizeof(REOBJECT);

	CLSID clsid;
	sc = lpOleObject->GetUserClassID(&clsid);
	if (sc != S_OK)
	{
		AtlThrow(sc);
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
	//reobject.sizel = sizel;
	reobject.dwFlags = REO_BELOWBASELINE;//REO_RESIZABLE
	reobject.dwUser = 0;//TODO 用户数据

	sc = pRichEditOle->InsertObject(&reobject);

	lpViewObject->Release();
	lpLockBytes->Release();
	lpOleObject->Release();
	lpStorage->Release();
}