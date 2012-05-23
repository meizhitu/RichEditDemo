#pragma once

//风格2 = (0x56201004, 0x00000000)，使用场景：会话窗口输入框、历史框
typedef CWinTraits<WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_WANTRETURN>	CExRichEditWinTraits;

const LPTSTR STR_PROGID = _T("DynamicOle.DynamicOleCom.1");
class CExRichEdit :
	public CWindowImpl<CExRichEdit, CRichEditCtrl, CExRichEditWinTraits>,
	public CRichEditCommands<CExRichEdit>
{
public:
	void InsertBitmap(CString& filePath);
	void InsertGif();
public:
	CExRichEdit();
	~CExRichEdit();

	DECLARE_WND_SUPERCLASS(_T("CExRichEdit"), MSFTEDIT_CLASS)

	BEGIN_MSG_MAP(CExRichEdit)

	END_MSG_MAP()

};