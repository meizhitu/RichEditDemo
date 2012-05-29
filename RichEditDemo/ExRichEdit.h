#pragma once
#include <TextServ.h>
#include "BitmapDC.h"

//风格2 = (0x56201004, 0x00000000)，使用场景：会话窗口输入框、历史框
typedef CWinTraits<WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_WANTRETURN>	CExRichEditWinTraits;


class CExRichEdit :
	public CWindowImpl<CExRichEdit, CRichEditCtrl, CExRichEditWinTraits>,
	public CRichEditCommands<CExRichEdit>
{
public:
	void InsertBitmap(CString& filePath) ;
	void InsertGif(LONG gif);
	void StartTimer();
public:
	CExRichEdit();
	~CExRichEdit();

	DECLARE_WND_SUPERCLASS(_T("CExRichEdit"), MSFTEDIT_CLASS)

	BEGIN_MSG_MAP(CExRichEdit)
		MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)
		MESSAGE_HANDLER(WM_ERASEBKGND,OnEraseBkgnd)
		MESSAGE_HANDLER(WM_PAINT, OnPaint)
		CHAIN_MSG_MAP_ALT(CRichEditCommands<CExRichEdit>, 1)
		DEFAULT_REFLECTION_HANDLER()
	END_MSG_MAP()
protected:
	LRESULT OnLButtonDblClk(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnPaint( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );

private:
	ITextServices* m_textServices;
	CBitmapDC m_memDC;

};