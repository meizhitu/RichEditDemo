// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include "ExRichEdit.h"
#include <atltypes.h>
#include "RichEditOleCallback.h"
#include "RichEditWindowless.h"
#include "TxtWinHost.h"


class CMainDlg : public CDialogImpl<CMainDlg>, public CUpdateUI<CMainDlg>,
	public CMessageFilter, public CIdleHandler
{
private:
	CExRichEdit m_richEdit;
	CExRichEditWindowless m_richEditLess;
public:
	enum { IDD = IDD_MAINDLG };

	virtual BOOL PreTranslateMessage(MSG* pMsg)
	{
		if(pMsg-> message   ==   WM_CHAR) 
		{
			return FALSE;
		}
		return CWindow::IsDialogMessage(pMsg);
	}

	virtual BOOL OnIdle()
	{
		return FALSE;
	}

	BEGIN_UPDATE_UI_MAP(CMainDlg)
	END_UPDATE_UI_MAP()

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_PAINT, OnPaint)
		MESSAGE_HANDLER(WM_KEYDOWN, OnMessage)
		MESSAGE_HANDLER(WM_MOUSEMOVE, OnMessage)
		MESSAGE_HANDLER(WM_KEYUP, OnMessage)
		MESSAGE_HANDLER(WM_LBUTTONDOWN, OnMessage)
		MESSAGE_HANDLER(WM_LBUTTONUP, OnMessage)
		MESSAGE_HANDLER(WM_SETFOCUS, OnMessage)
		MESSAGE_HANDLER(WM_CHAR, OnChar)
		MESSAGE_HANDLER( WM_SYSKEYDOWN, OnMessage )
		MESSAGE_HANDLER( WM_SYSKEYUP, OnMessage )
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(IDOK, OnOK)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
	END_MSG_MAP()

	// Handler prototypes (uncomment arguments if needed):
	//LRESULT MessageHandler(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	//	LRESULT CommandHandler(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
	//	LRESULT NotifyHandler(int /*idCtrl*/, LPNMHDR /*pnmh*/, BOOL& /*bHandled*/)

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		HWND richEditWnd = m_richEdit.Create(
			m_hWnd,
			CRect(0,0,300,150),
			L"",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_LEFT |WS_VSCROLL,
			0);
		CRichEditOleCallback *pRichEditOle = new CRichEditOleCallback;
		m_richEdit.SetOleCallback(pRichEditOle);

		TCHAR exeFullPath[MAX_PATH]; // MAX_PATH
		GetModuleFileName(NULL,exeFullPath,MAX_PATH);//得到程序模块名称，全路径
		CString fileAPath;
		fileAPath.Format(_T("%s\\..\\%s"),exeFullPath,_T("a.bmp"));
		CString fileGPath;
		fileGPath.Format(_T("%s\\..\\%s"),exeFullPath,_T("c.gif"));

		m_richEdit.InsertBitmap(fileAPath);
		Image* gif = Image::FromFile(fileGPath);
		m_richEdit.InsertGif((LONG)gif);
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.InsertGif((LONG)gif->Clone());
		m_richEdit.StartTimer();


		CREATESTRUCT cs;
		// No text for text control yet
		cs.lpszName = _T("abccccccccccccc");
		// No parent for this window
		cs.hwndParent = m_hWnd;
		// Style of control that we want
		cs.style = ES_MULTILINE;
		cs.dwExStyle = 0;
		// Rectangle for the control
		cs.x = 0;
		cs.y = 150;
		cs.cy = 300;
		cs.cx = 300;

		m_richEditLess.Init(m_hWnd,&cs,NULL);
		//CRichEditOleCallback *pRichEditOleLess = new CRichEditOleCallback;
		m_richEditLess.SetOLECallback(&m_richEditLess);

		Image* gifLess = Image::FromFile(fileGPath);
		m_richEditLess.InsertGif((LONG)gifLess);

		m_richEditLess.OnTxInPlaceActivate(NULL);
		m_richEditLess.AppendText(_T("windowless edit"),true);

		//CreateControlWindow(m_hWnd,NULL);
		// center the dialog on the screen
		CenterWindow();

		// set icons
		HICON hIcon = (HICON)::LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), 
			IMAGE_ICON, ::GetSystemMetrics(SM_CXICON), ::GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
		SetIcon(hIcon, TRUE);
		HICON hIconSmall = (HICON)::LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), 
			IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
		SetIcon(hIconSmall, FALSE);


		// register object for message filtering and idle updates
		CMessageLoop* pLoop = _Module.GetMessageLoop();
		ATLASSERT(pLoop != NULL);
		pLoop->AddMessageFilter(this);
		pLoop->AddIdleHandler(this);


		UIAddChildWindowContainer(m_hWnd);

		return TRUE;
	}


	LRESULT OnMessage( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
	{
		return m_richEditLess.TxWindowProc(m_hWnd, uMsg, wParam,lParam  );
	}

	LRESULT OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
	{
		return m_richEditLess.TxWindowProc(m_hWnd, uMsg, wParam,lParam  );
	}

	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		// unregister message filtering and idle updates
		CMessageLoop* pLoop = _Module.GetMessageLoop();
		ATLASSERT(pLoop != NULL);
		pLoop->RemoveMessageFilter(this);
		pLoop->RemoveIdleHandler(this);

		return 0;
	}

	LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		CPaintDC dc(m_hWnd);
		//m_richEditLess.OnPaint(dc.m_hDC,m_richEditLess.m_rect);
		m_richEditLess.TxWindowProc(m_hWnd, uMsg, wParam, lParam);
		return 0;
	}

	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
	{
		CAboutDlg dlg;
		dlg.DoModal();
		return 0;
	}

	LRESULT OnOK(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
	{
		// TODO: Add validation code 
		CloseDialog(wID);
		return 0;
	}

	LRESULT OnCancel(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
	{
		CloseDialog(wID);
		return 0;
	}

	void CloseDialog(int nVal)
	{
		DestroyWindow();
		::PostQuitMessage(nVal);
	}
};