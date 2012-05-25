#pragma once
#include "ExOleInPlaceSite.h"
#include "ExOleClientSite.h"

//风格2 = (0x56201004, 0x00000000)，使用场景：会话窗口输入框、历史框
typedef CWinTraits<WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_WANTRETURN>	CExRichEditWinTraits;

const LPTSTR STR_PROGID = _T("DynamicOle.DynamicOleCom.1");
class CExRichEdit :
	public CWindowImpl<CExRichEdit, CRichEditCtrl, CExRichEditWinTraits>,
	public CRichEditCommands<CExRichEdit>,
	public IExOleInPlaceSiteImpl,
	public IExOleClientSiteImpl
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
		CHAIN_MSG_MAP_ALT(CRichEditCommands<CExRichEdit>, 1)
		DEFAULT_REFLECTION_HANDLER()
	END_MSG_MAP()
protected:
	LRESULT OnLButtonDblClk(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );

};