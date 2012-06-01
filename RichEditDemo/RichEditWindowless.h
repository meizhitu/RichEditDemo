#pragma once
#include <tom.h>
#include "TxtWinHost.h"
#include <richole.h>
#include "RichEditOleCallback.h"

class CExRichEditWindowless : public CTxtWinHost,public CRichEditOleCallback
{
public:
	CExRichEditWindowless();
	~CExRichEditWindowless();
public:
	IRichEditOle* GetIRichEditOle();
	BOOL SetOLECallback(IRichEditOleCallback* pCallback);

	STDMETHOD(QueryAcceptData) (LPDATAOBJECT lpdataobj, CLIPFORMAT FAR * lpcfFormat, DWORD reco, BOOL fReally, HGLOBAL hMetaPict);
	STDMETHOD(GetClipboardData) (CHARRANGE FAR * lpchrg, DWORD reco, LPDATAOBJECT FAR * lplpdataobj);

public:
	void InsertGif(LONG gif);
	int InsertText(long nInsertAfterChar, LPCTSTR lpstrText, bool bCanUndo);

	int AppendText(LPCTSTR lpstrText, bool bCanUndo);

	int SetSel(long nStartChar, long nEndChar);

	void ReplaceSel(LPCTSTR lpszNewText, bool bCanUndo);
	CString GetTextRange(long nStartChar, long nEndChar);

};