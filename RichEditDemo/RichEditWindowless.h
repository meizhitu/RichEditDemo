#pragma once
#include <tom.h>
#include "TxtWinHost.h"
#include <richole.h>

class CExRichEditWindowless : public CTxtWinHost
{
public:
	CExRichEditWindowless();
	~CExRichEditWindowless();
public:
	IRichEditOle* GetIRichEditOle();
	BOOL SetOLECallback(IRichEditOleCallback* pCallback);

public:
	void InsertGif(LONG gif);
	int InsertText(long nInsertAfterChar, LPCTSTR lpstrText, bool bCanUndo);

	int AppendText(LPCTSTR lpstrText, bool bCanUndo);

	int SetSel(long nStartChar, long nEndChar);

	void ReplaceSel(LPCTSTR lpszNewText, bool bCanUndo);

};