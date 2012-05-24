#pragma once

class CRichEditOleCallback: public IRichEditOleCallback
{
public:
	CRichEditOleCallback();
	virtual ~CRichEditOleCallback();
	int m_iNumStorages;
	IStorage* pStorage;
	DWORD m_dwRef;

#pragma region  IUnknown methods 
	STDMETHOD(QueryInterface) (REFIID riid, LPVOID FAR * lplpObj);

	STDMETHOD_(ULONG,AddRef) (void);

	STDMETHOD_(ULONG,Release) (void);
#pragma endregion

#pragma region IRichEditOleCallback methods
	STDMETHOD(GetNewStorage) (LPSTORAGE FAR * lplpstg);

	STDMETHOD(GetInPlaceContext) (LPOLEINPLACEFRAME FAR * lplpFrame, LPOLEINPLACEUIWINDOW FAR * lplpDoc, LPOLEINPLACEFRAMEINFO lpFrameInfo);

	STDMETHOD(ShowContainerUI) (BOOL fShow);

	STDMETHOD(QueryInsertObject) (LPCLSID lpclsid, LPSTORAGE lpstg, LONG cp);

	STDMETHOD(DeleteObject) (LPOLEOBJECT lpoleobj);

	STDMETHOD(QueryAcceptData) (LPDATAOBJECT lpdataobj, CLIPFORMAT FAR * lpcfFormat, DWORD reco, BOOL fReally, HGLOBAL hMetaPict);

	STDMETHOD(ContextSensitiveHelp) (BOOL fEnterMode);

	STDMETHOD(GetClipboardData) (CHARRANGE FAR * lpchrg, DWORD reco, LPDATAOBJECT FAR * lplpdataobj);

	STDMETHOD(GetDragDropEffect) (THIS_ BOOL fDrag, DWORD grfKeyState, LPDWORD pdwEffect);

	STDMETHOD(GetContextMenu) (THIS_ WORD seltype, LPOLEOBJECT lpoleobj, CHARRANGE FAR * lpchrg, HMENU FAR * lphmenu);
#pragma endregion
};