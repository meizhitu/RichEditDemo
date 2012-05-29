#include "stdafx.h"
/**************************************************************************
   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
   PARTICULAR PURPOSE.

   Copyright 2000 Microsoft Corporation.  All Rights Reserved.
**************************************************************************/


/*************************************************************************
 HOST.C	-- Text Host for Windowless Rich Edit Control 
 
***************************************************************************/

#include "windows.h"
#include "stddef.h"
#include "windowsx.h"
#include "richedit.h"
#include "memory.h"
#include <imm.h>
#include "TxtWinHost.h"
#include <tchar.h>

BOOL fInAssert = FALSE;

// HIMETRIC units per inch (used for conversion)
#define HIMETRIC_PER_INCH 2540

// Convert Himetric along the X axis to X pixels
LONG HimetricXtoDX(LONG xHimetric, LONG xPerInch)
{
	return (LONG) MulDiv(xHimetric, xPerInch, HIMETRIC_PER_INCH);
}

// Convert Himetric along the Y axis to Y pixels
LONG HimetricYtoDY(LONG yHimetric, LONG yPerInch)
{
	return (LONG) MulDiv(yHimetric, yPerInch, HIMETRIC_PER_INCH);
}

// Convert Pixels on the X axis to Himetric
LONG DXtoHimetricX(LONG dx, LONG xPerInch)
{
	return (LONG) MulDiv(dx, HIMETRIC_PER_INCH, xPerInch);
}

// Convert Pixels on the Y axis to Himetric
LONG DYtoHimetricY(LONG dy, LONG yPerInch)
{
	return (LONG) MulDiv(dy, HIMETRIC_PER_INCH, yPerInch);
}


// These constants are for backward compatibility. They are the 
// sizes used for initialization and reset in RichEdit 1.0

const LONG cInitTextMax = (32 * 1024) - 1;
const LONG cResetTextMax = (64 * 1024);


#define ibPed 0
#define SYS_ALTERNATE 0x20000000




INT	cxBorder, cyBorder;	    // GetSystemMetricx(SM_CXBORDER)...
INT	cxDoubleClk, cyDoubleClk;   // Double click distances
INT	cxHScroll, cxVScroll;	    // Width/height of scrlbar arw bitmap
INT	cyHScroll, cyVScroll;	    // Width/height of scrlbar arw bitmap
INT	dct;			    // Double Click Time in milliseconds
INT     nScrollInset;
COLORREF crAuto = 0;

LONG CTxtWinHost::xWidthSys = 0;    		            // average char width of system font
LONG CTxtWinHost::yHeightSys = 0;				// height of system font
LONG CTxtWinHost::yPerInch = 0;				// y pixels per inch
LONG CTxtWinHost::xPerInch = 0;				// x pixels per inch


//EXTERN_C const IID IID_ITextServices = { // 8d33f740-cf58-11ce-a89d-00aa006cadc5
//    0x8d33f740,
//    0xcf58,
//    0x11ce,
//    {0xa8, 0x9d, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}
//  };
//
//EXTERN_C const IID IID_ITextHost = { /* c5bdd8d0-d26e-11ce-a89e-00aa006cadc5 */
//    0xc5bdd8d0,
//    0xd26e,
//    0x11ce,
//    {0xa8, 0x9e, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}
//  };
//
//
EXTERN_C const IID IID_ITextEditControl = { /* f6642620-d266-11ce-a89e-00aa006cadc5 */
    0xf6642620,
    0xd266,
    0x11ce,
    {0xa8, 0x9e, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}
  };

void GetSysParms(void)
{
	crAuto		= GetSysColor(COLOR_WINDOWTEXT);
	cxBorder	= GetSystemMetrics(SM_CXBORDER);	// Unsizable window border
	cyBorder	= GetSystemMetrics(SM_CYBORDER);	//  widths
	cxHScroll	= GetSystemMetrics(SM_CXHSCROLL);	// Scrollbar-arrow bitmap 
	cxVScroll	= GetSystemMetrics(SM_CXVSCROLL);	//  dimensions
	cyHScroll	= GetSystemMetrics(SM_CYHSCROLL);	//
	cyVScroll	= GetSystemMetrics(SM_CYVSCROLL);	//
	cxDoubleClk	= GetSystemMetrics(SM_CXDOUBLECLK);
	cyDoubleClk	= GetSystemMetrics(SM_CYDOUBLECLK);
	dct			= GetDoubleClickTime();
    
    nScrollInset =
        GetProfileIntA( "windows", "ScrollInset", DD_DEFSCROLLINSET );
}

HRESULT InitDefaultCharFormat(CHARFORMATW * pcf, HFONT hfont) 
{
	HWND hwnd;
	LOGFONT lf;
	HDC hdc;
	LONG yPixPerInch;

	// Get LOGFONT for default font
	if (!hfont)
		hfont = (HFONT)GetStockObject(SYSTEM_FONT);

	// Get LOGFONT for passed hfont
	if (!GetObject(hfont, sizeof(LOGFONT), &lf))
		return E_FAIL;

	// Set CHARFORMAT structure
	pcf->cbSize = sizeof(CHARFORMAT2);
	
	hwnd = GetDesktopWindow();
	hdc = GetDC(hwnd);
	yPixPerInch = GetDeviceCaps(hdc, LOGPIXELSY);
	pcf->yHeight = lf.lfHeight * LY_PER_INCH / yPixPerInch;
	ReleaseDC(hwnd, hdc);

	pcf->yOffset = 0;
	pcf->crTextColor = crAuto;

	pcf->dwEffects = CFM_EFFECTS | CFE_AUTOBACKCOLOR;
	pcf->dwEffects &= ~(CFE_PROTECTED | CFE_LINK);

	if(lf.lfWeight < FW_BOLD)
		pcf->dwEffects &= ~CFE_BOLD;
	if(!lf.lfItalic)
		pcf->dwEffects &= ~CFE_ITALIC;
	if(!lf.lfUnderline)
		pcf->dwEffects &= ~CFE_UNDERLINE;
	if(!lf.lfStrikeOut)
		pcf->dwEffects &= ~CFE_STRIKEOUT;

	pcf->dwMask = CFM_ALL | CFM_BACKCOLOR;
	pcf->bCharSet = lf.lfCharSet;
	pcf->bPitchAndFamily = lf.lfPitchAndFamily;
#ifdef UNICODE
	_tcscpy(pcf->szFaceName, lf.lfFaceName);
#else
	//need to thunk pcf->szFaceName to a standard char string.in this case it's easy because our thunk is also our copy
	MultiByteToWideChar(CP_ACP, 0, lf.lfFaceName, LF_FACESIZE, pcf->szFaceName, LF_FACESIZE) ;
#endif
	
	return S_OK;
}

HRESULT InitDefaultParaFormat(PARAFORMAT * ppf) 
{	
	memset(ppf, 0, sizeof(PARAFORMAT));

	ppf->cbSize = sizeof(PARAFORMAT);
	ppf->dwMask = PFM_ALL;
	ppf->wAlignment = PFA_LEFT;
	ppf->cTabCount = 1;
   ppf->rgxTabs[0] = lDefaultTab;

	return S_OK;
}



LRESULT MapHresultToLresult(HRESULT hr, UINT msg)
{
	LRESULT lres = hr;

	switch(msg)
	{
	case EM_GETMODIFY:


		lres = (hr == S_OK) ? -1 : 0;

		break;

		// These messages must return TRUE/FALSE rather than an hresult.
	case EM_UNDO:
	case WM_UNDO:
	case EM_CANUNDO:
	case EM_CANPASTE:
	case EM_LINESCROLL:

		// Hresults are backwards from TRUE and FALSE so we need
		// to do that remapping here as well.

		lres = (hr == S_OK) ? TRUE : FALSE;

		break;

	case EM_EXLINEFROMCHAR:
	case EM_LINEFROMCHAR:

		// If success, then hr a number. If error, it s/b 0.
		lres = SUCCEEDED(hr) ? (LRESULT) hr : 0;
		break;
			
	case EM_LINEINDEX:

		// If success, then hr a number. If error, it s/b -1.
		lres = SUCCEEDED(hr) ? (LRESULT) hr : -1;
		break;	

	default:
		lres = (LRESULT) hr;		
	}

	return lres;
}


BOOL GetIconic(HWND hwnd) 
{
	while(hwnd)
	{
		if(::IsIconic(hwnd))
			return TRUE;
		hwnd = GetParent(hwnd);
	}
	return FALSE;
}


CTxtWinHost::CTxtWinHost()
{
	ZeroMemory(&pnc, sizeof(CTxtWinHost) - offsetof(CTxtWinHost, pnc));

	cchTextMost = cInitTextMax;
    
	laccelpos = -1;
}

CTxtWinHost::~CTxtWinHost()
{
	// Revoke our drop target
	RevokeDragDrop();

	pserv->OnTxInPlaceDeactivate();

	pserv->Release();
}


////////////////////// Create/Init/Destruct Commands ///////////////////////


HRESULT CreateHost(
	HWND hwnd, 
	const CREATESTRUCT *pcs,
	PNOTIFY_CALL pnc,
	CTxtWinHost **pptec)
{
	HRESULT hr = E_FAIL;
    GdiSetBatchLimit(1);

	CTxtWinHost *phost = new CTxtWinHost();

	if(phost)
	{
		if (phost->Init(hwnd, pcs, pnc))
		{
			*pptec = phost;
			hr = S_OK;
		}
	}

	if (FAILED(hr))
	{
		delete phost;
	}

	return TRUE;
}


/*
 *	CTxtWinHost::Init
 *
 *	Purpose:
 *		Initializes this CTxtWinHost
 *
 *	TODO:	Format Cache logic needs to be cleaned up. 
 */
BOOL CTxtWinHost::Init(
	HWND h_wnd, 
	const CREATESTRUCT *pcs,
	PNOTIFY_CALL p_nc)
{
    HDC hdc;
    HFONT hfontOld;
    TEXTMETRIC tm;
	IUnknown *pUnk;
	HRESULT hr;

	// Initialize Reference count
	cRefs = 1;

	// Set up the notification callback
	pnc = p_nc;	
	
	hwnd = h_wnd;

	// Create and cache CHARFORMAT for this control
	if(FAILED(InitDefaultCharFormat(&cf, NULL)))
		goto err;
		
	// Create and cache PARAFORMAT for this control
	if(FAILED(InitDefaultParaFormat(&pf)))
		goto err;

 	// edit controls created without a window are multiline by default
	// so that paragraph formats can be
	dwStyle = ES_MULTILINE;
	fHidden = TRUE;

	// edit controls are rich by default
	fRich = TRUE;
	
	if(pcs)
	{
		hwndParent = pcs->hwndParent;
		dwExStyle = pcs->dwExStyle;
		dwStyle = pcs->style;

		fBorder = !!(dwStyle & WS_BORDER);

		if(dwStyle & ES_SUNKEN)
		{
			fBorder = TRUE;
		}

		if (!(dwStyle & (ES_AUTOHSCROLL | WS_HSCROLL)))
		{
			fWordWrap = TRUE;
		}
	}

	if(!(dwStyle & ES_LEFT))
	{
		if(dwStyle & ES_CENTER)
			pf.wAlignment = PFA_CENTER;
		else if(dwStyle & ES_RIGHT)
			pf.wAlignment = PFA_RIGHT;
	}

    // Init system metrics
	hdc = GetDC(hwnd);
    if(!hdc)
        goto err;

   	hfontOld = (HFONT)SelectObject(hdc, GetStockObject(SYSTEM_FONT));

	if(!hfontOld)
		goto err;

	GetTextMetrics(hdc, &tm);
	SelectObject(hdc, hfontOld);

	xWidthSys = (INT) tm.tmAveCharWidth;
   yHeightSys = (INT) tm.tmHeight;
	xPerInch = GetDeviceCaps(hdc, LOGPIXELSX); 
	yPerInch =	GetDeviceCaps(hdc, LOGPIXELSY); 

	ReleaseDC(hwnd, hdc);

	// At this point the border flag is set and so is the pixels per inch
	// so we can initalize the inset.
	SetDefaultInset();

	fInplaceActive = TRUE;

	// Create Text Services component
	if(FAILED(CreateTextServices(NULL, this, &pUnk)))
		goto err;

	hr = pUnk->QueryInterface(IID_ITextServices,(void **)&pserv);

	// Whether the previous call succeeded or failed we are done
	// with the private interface.
	pUnk->Release();

	if(FAILED(hr))
	{
		goto err;
	}

	// Set window text
	if(pcs && pcs->lpszName)
	{
#ifdef UNICODE		
		if(FAILED(pserv->TxSetText((TCHAR *)pcs->lpszName)))
			goto err;
#else
		//have to thunk the string to a unicode string.
		WCHAR wsz_name[MAX_PATH] ;
		ZeroMemory(wsz_name, MAX_PATH *sizeof WCHAR) ;
		MultiByteToWideChar(CP_ACP, 0, pcs->lpszName, lstrlen(pcs->lpszName),(LPWSTR)&wsz_name, MAX_PATH) ;
		if(FAILED(pserv->TxSetText((LPWSTR)&wsz_name)))
			goto err;
#endif
	}

	rcClient.left = pcs->x;
	rcClient.top = pcs->y;
	rcClient.right = pcs->x + pcs->cx;
	rcClient.bottom = pcs->y + pcs->cy;

	// The extent matches the full client rectangle in HIMETRIC
	sizelExtent.cx = DXtoHimetricX(pcs->cx - 2 * HOST_BORDER, xPerInch);
	sizelExtent.cy = DYtoHimetricY(pcs->cy - 2 * HOST_BORDER, yPerInch);

	// notify Text Services that we are in place active
	if(FAILED(pserv->OnTxInPlaceActivate(&rcClient)))
		goto err;
	
	// Hide all scrollbars to start
	if(hwnd && !(dwStyle & ES_DISABLENOSCROLL))
	{
		LONG dwStyle = GetWindowLong(hwnd, GWL_STYLE);
		dwStyle &= ~(WS_VSCROLL | WS_HSCROLL);
		SetWindowLong(hwnd, GWL_STYLE, dwStyle);
	}

	if (!(dwStyle & ES_READONLY))
	{
		// This isn't a read only window so we need a drop target.
		RegisterDragDrop();
	}

	return TRUE;

err:
	return FALSE;
}


/////////////////////////////////  IUnknown ////////////////////////////////


HRESULT CTxtWinHost::QueryInterface(REFIID riid, void **ppvObject)
{
	HRESULT hr = E_NOINTERFACE;
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_ITextEditControl))
	{
		*ppvObject = (ITextEditControl *) this;
		AddRef();
		hr = S_OK;
	}
	else if (IsEqualIID(riid, IID_IUnknown) 
		|| IsEqualIID(riid, IID_ITextHost)) 
	{
		AddRef();
		*ppvObject = (ITextHost *) this;
		hr = S_OK;
	}

	return hr;
}

ULONG CTxtWinHost::AddRef(void)
{
	return ++cRefs;
}

ULONG CTxtWinHost::Release(void)
{
	ULONG c_Refs = --cRefs;

	if (c_Refs == 0)
	{
		delete this;
	}

	return c_Refs;
}


//////////////////////////////// Properties ////////////////////////////////


TXTEFFECT CTxtWinHost::TxGetEffects() const
{
	return (dwStyle & ES_SUNKEN) ? TXTEFFECT_SUNKEN : TXTEFFECT_NONE;
}


//////////////////////////// System API wrapper ////////////////////////////



///////////////////////  Windows message dispatch methods  ///////////////////////////////


LRESULT CTxtWinHost::TxWindowProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
{
	LRESULT	lres = 0;
	HRESULT hr;

	switch(msg)
	{
	case WM_NCCALCSIZE:
		// we can't rely on WM_WININICHANGE so we use WM_NCCALCSIZE since
		// changing any of these should trigger a WM_NCCALCSIZE
		GetSysParms();
		break;

	case WM_KEYDOWN:
		lres = OnKeyDown((WORD) wparam, (DWORD) lparam);
		if(lres != 0)
			goto serv;
		break;		   

	case WM_CHAR:
		lres = OnChar((WORD) wparam, (DWORD) lparam);
		if(lres != 0)
			goto serv;
		break;

	case WM_SYSCOLORCHANGE:
		OnSysColorChange();

		// Notify the text services that there has been a change in the 
		// system colors.
		goto serv;

	case WM_GETDLGCODE:
		lres = OnGetDlgCode(wparam, lparam);
		break;

	case EM_HIDESELECTION:
		if((BOOL)lparam)
		{
			DWORD dwPropertyBits = 0;

			if((BOOL)wparam)
			{
				dwStyle &= ~(DWORD) ES_NOHIDESEL;
				dwPropertyBits = TXTBIT_HIDESELECTION;
			}
			else
				dwStyle |= ES_NOHIDESEL;

			// Notify text services of change in status.
			pserv->OnTxPropertyBitsChange(TXTBIT_HIDESELECTION, 
				dwPropertyBits);
		}

		goto serv;



    case EM_LIMITTEXT:

        lparam = wparam;

        // Intentionally fall through. These messages are duplicates.

	case EM_EXLIMITTEXT:

        if (lparam == 0)
        {
            // 0 means set the control to the maximum size. However, because
            // 1.0 set this to 64K will keep this the same value so as not to
            // supprise anyone. Apps are free to set the value to be above 64K.
            lparam = (LPARAM) cResetTextMax;
        }

		cchTextMost = (LONG) lparam;
		pserv->OnTxPropertyBitsChange(TXTBIT_MAXLENGTHCHANGE, 
					TXTBIT_MAXLENGTHCHANGE);

		break;

	case EM_SETREADONLY:
		OnSetReadOnly(BOOL(wparam));
		lres = 1;
		break;

	case EM_GETEVENTMASK:
		lres = OnGetEventMask();
		break;

	case EM_SETEVENTMASK:
		OnSetEventMask((DWORD) lparam);
		goto serv;

	case EM_GETOPTIONS:
		lres = OnGetOptions();
		break;

	case EM_SETOPTIONS:
		OnSetOptions((WORD) wparam, (DWORD) lparam);
		lres = (dwStyle & ECO_STYLES);
		if(fEnableAutoWordSel)
			lres |= ECO_AUTOWORDSELECTION;
		break;

	case WM_SETFONT:
		lres = OnSetFont((HFONT) wparam);
		break;

	case EM_SETRECT:
        OnSetRect((LPRECT)lparam);
        break;
        
	case EM_GETRECT:
        OnGetRect((LPRECT)lparam);
        break;

	case EM_SETBKGNDCOLOR:

		lres = (LRESULT) crBackground;
		fNotSysBkgnd = !wparam;
		crBackground = (COLORREF) lparam;

		if(wparam)
			crBackground = GetSysColor(COLOR_WINDOW);

		// Notify the text services that color has changed
		pserv->TxSendMessage(WM_SYSCOLORCHANGE, 0, 0, 0);

		if(lres != (LRESULT) crBackground)
			TxInvalidateRect(NULL, TRUE);

		break;

	case EM_SETCHARFORMAT:

		if(!FValidCF((CHARFORMAT *) lparam))
		{
			return 0;
		}

		if(wparam & SCF_SELECTION)
			goto serv;								// Change selection format
		OnSetCharFormat((CHARFORMAT *) lparam);		// Change default format
		break;

	case EM_SETPARAFORMAT:
		if(!FValidPF((PARAFORMAT *) lparam))
		{
			return 0;
		}

		// check to see if we're setting the default.
		// either SCF_DEFAULT will be specified *or* there is no
		// no text in the document (richedit1.0 behaviour).
		if (!(wparam & SCF_DEFAULT))
		{
			hr = pserv->TxSendMessage(WM_GETTEXTLENGTH, 0, 0, 0);

			if (hr == 0)
			{
				wparam |= SCF_DEFAULT;
			}
		}

		if(wparam & SCF_DEFAULT)
		{								
			OnSetParaFormat((PARAFORMAT *) lparam);	// Change default format
		}
		else
		{
			goto serv;								// Change selection format
		}
		break;

    case WM_SETTEXT:

        // For RichEdit 1.0, the max text length would be reset by a settext so 
        // we follow pattern here as well.

		hr = pserv->TxSendMessage(msg, wparam, lparam, 0);

        if (SUCCEEDED(hr))
        {
            // Update succeeded.
            LONG cNewText = _tcslen((LPCTSTR) lparam);

            // If the new text is greater than the max set the max to the new
            // text length.
            if (cNewText > cchTextMost)
            {
                cchTextMost = cNewText;                
            }

			lres = 1;
        }

        break;

	case WM_SIZE:
		lres = OnSize(hwnd, wparam, LOWORD(lparam), HIWORD(lparam));
		break;

	case WM_WINDOWPOSCHANGING:
		lres = ::DefWindowProc(hwnd, msg, wparam, lparam);

		if(TxGetEffects() == TXTEFFECT_SUNKEN)
			OnSunkenWindowPosChanging(hwnd, (WINDOWPOS *) lparam);
		break;

	case WM_SETCURSOR:
		//Only set cursor when over us rather than a child; this
		//			helps prevent us from fighting it out with an inplace child
		if((HWND)wparam == hwnd)
		{
			if(!(lres = ::DefWindowProc(hwnd, msg, wparam, lparam)))
			{
				POINT pt;
				GetCursorPos(&pt);
				::ScreenToClient(hwnd, &pt);
				pserv->OnTxSetCursor(
					DVASPECT_CONTENT,	
					-1,
					NULL,
					NULL,
					NULL,
					NULL,
					NULL,			// Client rect - no redraw 
					pt.x, 
					pt.y);
				lres = TRUE;
			}
		}
		break;

	case WM_SHOWWINDOW:
		hr = OnTxVisibleChange((BOOL)wparam);
		break;

	case WM_NCPAINT:

		lres = ::DefWindowProc(hwnd, msg, wparam, lparam);

		if(TxGetEffects() == TXTEFFECT_SUNKEN)
		{
			HDC hdc = GetDC(hwnd);

			if(hdc)
			{
				DrawSunkenBorder(hwnd, hdc);
				ReleaseDC(hwnd, hdc);
			}
		}
		break;

	case WM_PAINT:
		{
			// Put a frame around the control so it can be seen
			FrameRect((HDC) wparam, &rcClient, 
				(HBRUSH) GetStockObject(BLACK_BRUSH));

			RECT rcClient;
			RECT *prc = NULL;
			LONG lViewId = TXTVIEW_ACTIVE;

			if (!fInplaceActive)
			{
				GetControlRect(&rcClient);
				prc = &rcClient;
				lViewId = TXTVIEW_INACTIVE;
			}
			
			// Remember wparam is actually the hdc and lparam is the update
			// rect because this message has been preprocessed by the window.
			pserv->TxDraw(
	    		DVASPECT_CONTENT,  		// Draw Aspect
				/*-1*/0,						// Lindex
				NULL,					// Info for drawing optimazation
				NULL,					// target device information
	        	(HDC) wparam,			// Draw device HDC
	        	NULL, 				   	// Target device HDC
				(RECTL *) prc,			// Bounding client rectangle
				NULL, 					// Clipping rectangle for metafiles
				(RECT *) lparam,		// Update rectangle
				NULL, 	   				// Call back function
				NULL,					// Call back parameter
				lViewId);				// What view of the object				

			if(TxGetEffects() == TXTEFFECT_SUNKEN)
				DrawSunkenBorder(hwnd, (HDC) wparam);
		}

		break;

	default:
serv:
		{
			hr = pserv->TxSendMessage(msg, wparam, lparam, &lres);

			if (hr == S_FALSE)
			{
				lres = ::DefWindowProc(hwnd, msg, wparam, lparam);
			}
		}
	}

	return lres;
}
	

///////////////////////////////  Keyboard Messages  //////////////////////////////////


LRESULT CTxtWinHost::OnKeyDown(WORD vkey, DWORD dwFlags)
{
	switch(vkey)
	{
	case VK_ESCAPE:
		if(fInDialogBox)
		{
			PostMessage(hwndParent, WM_CLOSE, 0, 0);
			return 0;
		}
		break;
	
	case VK_RETURN:
		if(fInDialogBox && !(GetKeyState(VK_CONTROL) & 0x8000) 
				&& !(dwStyle & ES_WANTRETURN))
		{
			// send to default button
			LRESULT id;
			HWND hwndT;

			id = SendMessage(hwndParent, DM_GETDEFID, 0, 0);
			if(LOWORD(id) &&
				(hwndT = GetDlgItem(hwndParent, LOWORD(id))))
			{
				SendMessage(hwndParent, WM_NEXTDLGCTL, (WPARAM) hwndT, (LPARAM) 1);
				if(GetFocus() != hwnd)
					PostMessage(hwndT, WM_KEYDOWN, (WPARAM) VK_RETURN, 0);
			}
			return 0;
		}
		break;

	case VK_TAB:
		if(fInDialogBox) 
		{
			SendMessage(hwndParent, WM_NEXTDLGCTL, 
							!!(GetKeyState(VK_SHIFT) & 0x8000), 0);
			return 0;
		}
		break;
	}
	return 1;
}

#define CTRL(_ch) (_ch - 'A' + 1)

LRESULT CTxtWinHost::OnChar(WORD vkey, DWORD dwFlags)
{
	switch(vkey)
	{
	// Ctrl-Return generates Ctrl-J (LF), treat it as an ordinary return
	case CTRL('J'):
	case VK_RETURN:
		if(fInDialogBox && !(GetKeyState(VK_CONTROL) & 0x8000)
				 && !(dwStyle & ES_WANTRETURN))
			return 0;
		break;

	case VK_TAB:
		if(fInDialogBox && !(GetKeyState(VK_CONTROL) & 0x8000))
			return 0;
	}
	
	return 1;
}


////////////////////////////////////  View rectangle //////////////////////////////////////


void CTxtWinHost::OnGetRect(LPRECT prc)
{
    RECT rcInset;

	// Get view inset (in HIMETRIC)
    TxGetViewInset(&rcInset);

	// Convert the himetric inset to pixels
	rcInset.left = HimetricXtoDX(rcInset.left, xPerInch);
	rcInset.top = HimetricYtoDY(rcInset.top , yPerInch);
	rcInset.right = HimetricXtoDX(rcInset.right, xPerInch);
	rcInset.bottom = HimetricYtoDY(rcInset.bottom, yPerInch);
    
	// Get client rect in pixels
    TxGetClientRect(prc);

	// Modify the client rect by the inset 
    prc->left += rcInset.left;
    prc->top += rcInset.top;
    prc->right -= rcInset.right;
    prc->bottom -= rcInset.bottom;
}

void CTxtWinHost::OnSetRect(LPRECT prc)
{
	RECT rcClient;
	
	if(!prc)
	{
		SetDefaultInset();
	}	
	else	
    {
    	// For screen display, the following intersects new view RECT
    	// with adjusted client area RECT
    	TxGetClientRect(&rcClient);

        // Adjust client rect
        // Factors in space for borders
        if(fBorder)
        {																					  
    	    rcClient.top		+= yHeightSys / 4;
    	    rcClient.bottom 	-= yHeightSys / 4 - 1;
    	    rcClient.left		+= xWidthSys / 2;
    	    rcClient.right	-= xWidthSys / 2;
        }
	
        // Ensure we have minimum width and height
        rcClient.right = max(rcClient.right, rcClient.left + xWidthSys);
        rcClient.bottom = max(rcClient.bottom, rcClient.top + yHeightSys);

        // Intersect the new view rectangle with the 
        // adjusted client area rectangle
        if(!IntersectRect(&rcViewInset, &rcClient, prc))
    	    rcViewInset = rcClient;

        // compute inset in pixels
        rcViewInset.left -= rcClient.left;
        rcViewInset.top -= rcClient.top;
        rcViewInset.right = rcClient.right - rcViewInset.right;
        rcViewInset.bottom = rcClient.bottom - rcViewInset.bottom;

		// Convert the inset to himetric that must be returned to ITextServices
        rcViewInset.left = DXtoHimetricX(rcViewInset.left, xPerInch);
        rcViewInset.top = DYtoHimetricY(rcViewInset.top, yPerInch);
        rcViewInset.right = DXtoHimetricX(rcViewInset.right, xPerInch);
        rcViewInset.bottom = DYtoHimetricY(rcViewInset.bottom, yPerInch);
    }

    pserv->OnTxPropertyBitsChange(TXTBIT_VIEWINSETCHANGE, 
    	TXTBIT_VIEWINSETCHANGE);
}



////////////////////////////////////  System notifications  //////////////////////////////////


void CTxtWinHost::OnSysColorChange()
{
	crAuto = GetSysColor(COLOR_WINDOWTEXT);
	if(!fNotSysBkgnd)
		crBackground = GetSysColor(COLOR_WINDOW);
	TxInvalidateRect(NULL, TRUE);
}

LRESULT CTxtWinHost::OnGetDlgCode(WPARAM wparam, LPARAM lparam)
{
	LRESULT lres = DLGC_WANTCHARS | DLGC_WANTARROWS | DLGC_WANTTAB;

	if(dwStyle & ES_MULTILINE)
		lres |= DLGC_WANTALLKEYS;

	if(!(dwStyle & ES_SAVESEL))
		lres |= DLGC_HASSETSEL;

	if(lparam)
		fInDialogBox = TRUE;

	if(lparam &&
		((WORD) wparam == VK_BACK))
	{
		lres |= DLGC_WANTMESSAGE;
	}

	return lres;
}


/////////////////////////////////  Other messages  //////////////////////////////////////


LRESULT CTxtWinHost::OnGetOptions() const
{
	LRESULT lres = (dwStyle & ECO_STYLES);

	if(fEnableAutoWordSel)
		lres |= ECO_AUTOWORDSELECTION;
	
	return lres;
}

void CTxtWinHost::OnSetOptions(WORD wOp, DWORD eco)
{
	const BOOL fAutoWordSel = !!(eco & ECO_AUTOWORDSELECTION);
	DWORD dwStyleNew = dwStyle;
	DWORD dw_Style = 0 ;

	DWORD dwChangeMask = 0;

	// single line controls can't have a selection bar
	// or do vertical writing
	if(!(dw_Style & ES_MULTILINE))
	{
#ifdef DBCS
		eco &= ~(ECO_SELECTIONBAR | ECO_VERTICAL);
#else
		eco &= ~ECO_SELECTIONBAR;
#endif
	}
	dw_Style = (eco & ECO_STYLES);

	switch(wOp)
	{
	case ECOOP_SET:
		dwStyleNew = ((dwStyleNew) & ~ECO_STYLES) | dwStyle;
		fEnableAutoWordSel = fAutoWordSel;
		break;

	case ECOOP_OR:
		dwStyleNew |= dw_Style;
		if(fAutoWordSel)
			fEnableAutoWordSel = TRUE;
		break;

	case ECOOP_AND:
		dwStyleNew &= (dw_Style | ~ECO_STYLES);
		if(fEnableAutoWordSel && !fAutoWordSel)
			fEnableAutoWordSel = FALSE;
		break;

	case ECOOP_XOR:
		dwStyleNew ^= dw_Style;
		fEnableAutoWordSel = (!fEnableAutoWordSel != !fAutoWordSel);
		break;
	}

	if(fEnableAutoWordSel != (unsigned)fAutoWordSel)
	{
		dwChangeMask |= TXTBIT_AUTOWORDSEL; 
	}

	if(dwStyleNew != dw_Style)
	{
		DWORD dwChange = dwStyleNew ^ dw_Style;
#ifdef DBCS
		USHORT	usMode;
#endif

		dwStyle = dwStyleNew;
		SetWindowLong(hwnd, GWL_STYLE, dwStyleNew);

		if(dwChange & ES_NOHIDESEL)	
		{
			dwChangeMask |= TXTBIT_HIDESELECTION;
		}

		if(dwChange & ES_READONLY)
		{
			dwChangeMask |= TXTBIT_READONLY;

			// Change drop target state as appropriate.
			if (dwStyleNew & ES_READONLY)
			{
				RevokeDragDrop();
			}
			else
			{
				RegisterDragDrop();
			}
		}

		if(dwChange & ES_VERTICAL)
		{
			dwChangeMask |= TXTBIT_VERTICAL;
		}

		// no action require for ES_WANTRETURN
		// no action require for ES_SAVESEL
		// do this last
		if(dwChange & ES_SELECTIONBAR)
		{
			lSelBarWidth = 212;
			dwChangeMask |= TXTBIT_SELBARCHANGE;
		}
	}

	if (dwChangeMask)
	{
		DWORD dwProp = 0;
		TxGetPropertyBits(dwChangeMask, &dwProp);
		pserv->OnTxPropertyBitsChange(dwChangeMask, dwProp);
	}
}

void CTxtWinHost::OnSetReadOnly(BOOL fReadOnly)
{
	DWORD dwUpdatedBits = 0;

	if(fReadOnly)
	{
		dwStyle |= ES_READONLY;

		// Turn off Drag Drop 
		RevokeDragDrop();
		dwUpdatedBits |= TXTBIT_READONLY;
	}
	else
	{
		dwStyle &= ~(DWORD) ES_READONLY;

		// Turn drag drop back on
		RegisterDragDrop();	
	}

	pserv->OnTxPropertyBitsChange(TXTBIT_READONLY, dwUpdatedBits);
}

void CTxtWinHost::OnSetEventMask(DWORD mask)
{
	LRESULT lres = (LRESULT) dwEventMask;
	dwEventMask = (DWORD) mask;

}


LRESULT CTxtWinHost::OnGetEventMask() const
{
	return (LRESULT) dwEventMask;
}

/*
 *	CTxtWinHost::OnSetFont(hfont)
 *
 *	Purpose:	
 *		Set new font from hfont
 *
 *	Arguments:
 *		hfont	new font (NULL for system font)
 */
BOOL CTxtWinHost::OnSetFont(HFONT hfont)
{
	if(SUCCEEDED(InitDefaultCharFormat(&cf, hfont)))
	{
		pserv->OnTxPropertyBitsChange(TXTBIT_CHARFORMATCHANGE, 
			TXTBIT_CHARFORMATCHANGE);
		return TRUE;
	}

	return FALSE;
}

/*
 *	CTxtWinHost::OnSetCharFormat(pcf)
 *
 *	Purpose:	
 *		Set new default CharFormat
 *
 *	Arguments:
 *		pch		ptr to new CHARFORMAT
 */
BOOL CTxtWinHost::OnSetCharFormat(CHARFORMAT *pcf)
{
	pserv->OnTxPropertyBitsChange(TXTBIT_CHARFORMATCHANGE, 
		TXTBIT_CHARFORMATCHANGE);

	return TRUE;
}

/*
 *	CTxtWinHost::OnSetParaFormat(ppf)
 *
 *	Purpose:	
 *		Set new default ParaFormat
 *
 *	Arguments:
 *		pch		ptr to new PARAFORMAT
 */
BOOL CTxtWinHost::OnSetParaFormat(PARAFORMAT *pPF)
{
	pf = *pPF;									// Copy it

	pserv->OnTxPropertyBitsChange(TXTBIT_PARAFORMATCHANGE, 
		TXTBIT_PARAFORMATCHANGE);

	return TRUE;
}



////////////////////////////  Event firing  /////////////////////////////////



void * CTxtWinHost::CreateNmhdr(UINT uiCode, LONG cb)
{
	NMHDR *pnmhdr;

	pnmhdr = (NMHDR*) new char[cb];
	if(!pnmhdr)
		return NULL;

	memset(pnmhdr, 0, cb);

	pnmhdr->hwndFrom = hwnd;
	pnmhdr->idFrom = GetWindowLong(hwnd, GWL_ID);
	pnmhdr->code = uiCode;

	return (VOID *) pnmhdr;
}


////////////////////////////////////  Helpers  /////////////////////////////////////////
void CTxtWinHost::SetDefaultInset()
{
    // Generate default view rect from client rect.
    if(fBorder)
    {
        // Factors in space for borders
  	    rcViewInset.top = DYtoHimetricY(yHeightSys / 4, yPerInch);
   	 rcViewInset.bottom	= DYtoHimetricY(yHeightSys / 4 - 1, yPerInch);
   	 rcViewInset.left = DXtoHimetricX(xWidthSys / 2, xPerInch);
   	 rcViewInset.right = DXtoHimetricX(xWidthSys / 2, xPerInch);
    }
    else
    {
		rcViewInset.top = rcViewInset.left =
		rcViewInset.bottom = rcViewInset.right = 0;
	}
}


/////////////////////////////////  Far East Support  //////////////////////////////////////

HIMC CTxtWinHost::TxImmGetContext(void)
{
#if 0
	HIMC himc;

	himc = ImmGetContext( hwnd );

	return himc;
#endif // 0

	return NULL;
}

void CTxtWinHost::TxImmReleaseContext(HIMC himc)
{
	ImmReleaseContext( hwnd, himc );
}

void CTxtWinHost::RevokeDragDrop(void)
{
	if (fRegisteredForDrop)
	{
		::RevokeDragDrop(hwnd);
		fRegisteredForDrop = FALSE;
	}
}

void CTxtWinHost::RegisterDragDrop(void)
{
	IDropTarget *pdt;

	if(!fRegisteredForDrop && pserv->TxGetDropTarget(&pdt) == NOERROR)
	{
		HRESULT hr = ::RegisterDragDrop(hwnd, pdt);

		if(hr == NOERROR)
		{	
			fRegisteredForDrop = TRUE;
		}

		pdt->Release();
	}
}

VOID DrawRectFn(
	HDC hdc, 
	RECT *prc, 
	INT icrTL, 
	INT icrBR,
	BOOL fBot, 
	BOOL fRght)
{
	COLORREF cr;
	COLORREF crSave;
	RECT rc;

	cr = GetSysColor(icrTL);
	crSave = SetBkColor(hdc, cr);

	// top
	rc = *prc;
	rc.bottom = rc.top + 1;
	ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);

	// left
	rc.bottom = prc->bottom;
	rc.right = rc.left + 1;
	ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);

	if(icrTL != icrBR)
	{
		cr = GetSysColor(icrBR);
		SetBkColor(hdc, cr);
	}

	// right
	rc.right = prc->right;
	rc.left = rc.right - 1;
	if(!fBot)
		rc.bottom -= cyHScroll;
	if(fRght)
	{
		ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
	}

	// bottom
	if(fBot)
	{
		rc.left = prc->left;
		rc.top = rc.bottom - 1;
		if(!fRght)
			rc.right -= cxVScroll;
		ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
	}
	SetBkColor(hdc, crSave);
}

#define cmultBorder 1

VOID CTxtWinHost::OnSunkenWindowPosChanging(HWND hwnd, WINDOWPOS *pwndpos)
{
	if(fVisible)
	{
		RECT rc;
		HWND hwndParent;

		GetWindowRect(hwnd, &rc);
		InflateRect(&rc, cxBorder * cmultBorder, cyBorder * cmultBorder);
		hwndParent = GetParent(hwnd);
		MapWindowPoints(HWND_DESKTOP, hwndParent, (POINT *) &rc, 2);
		InvalidateRect(hwndParent, &rc, FALSE);
	}
}


VOID CTxtWinHost::DrawSunkenBorder(HWND hwnd, HDC hdc)
{
	RECT rc;
	RECT rcParent;
	DWORD dwScrollBars;
	HWND hwndParent;

	GetWindowRect(hwnd, &rc);
    hwndParent = GetParent(hwnd);
	rcParent = rc;
	MapWindowPoints(HWND_DESKTOP, hwndParent, (POINT *)&rcParent, 2);
	InflateRect(&rcParent, cxBorder, cyBorder);
	OffsetRect(&rc, -rc.left, -rc.top);

	// draw inner rect
	TxGetScrollBars(&dwScrollBars);
	DrawRectFn(hdc, &rc, COLOR_WINDOWFRAME, COLOR_BTNFACE,
		!(dwScrollBars & WS_HSCROLL), !(dwScrollBars & WS_VSCROLL));

	// draw outer rect
	hwndParent = GetParent(hwnd);
	hdc = GetDC(hwndParent);
	DrawRectFn(hdc, &rcParent, COLOR_BTNSHADOW, COLOR_BTNHIGHLIGHT,
		TRUE, TRUE);
	ReleaseDC(hwndParent, hdc);
}

LRESULT CTxtWinHost::OnSize(HWND hwnd, WORD fwSizeType, int nWidth, int nHeight)
{
	// Update our client rectangle
	rcClient.right = rcClient.left + nWidth;
	rcClient.bottom = rcClient.top + nHeight;

	if(!fVisible)
	{
		fIconic = GetIconic(hwnd);
		if(!fIconic)
			fResized = TRUE;
	}
	else
	{
		if(GetIconic(hwnd))
		{
			fIconic = TRUE;
		}
		else
		{
			pserv->OnTxPropertyBitsChange(TXTBIT_CLIENTRECTCHANGE, 
				TXTBIT_CLIENTRECTCHANGE);

			if(fIconic)
			{
				InvalidateRect(hwnd, NULL, FALSE);
				fIconic = FALSE;
			}
			
			if(TxGetEffects() == TXTEFFECT_SUNKEN)	// Draw borders
				DrawSunkenBorder(hwnd, NULL);
		}
	}
	return 0;
}

HRESULT CTxtWinHost::OnTxVisibleChange(BOOL fVisible)
{
	fVisible = fVisible;

	if(!fVisible && fResized)
	{
		RECT rc;
		// Control was resized while hidden,
		// need to really resize now
		TxGetClientRect(&rc);
		fResized = FALSE;
		pserv->OnTxPropertyBitsChange(TXTBIT_CLIENTRECTCHANGE, 
			TXTBIT_CLIENTRECTCHANGE);
	}

	return S_OK;
}



//////////////////////////// ITextHost Interface  ////////////////////////////

HDC CTxtWinHost::TxGetDC()
{
	return ::GetDC(hwnd);
}


int CTxtWinHost::TxReleaseDC(HDC hdc)
{
	return ::ReleaseDC (hwnd, hdc);
}


BOOL CTxtWinHost::TxShowScrollBar(INT fnBar,	BOOL fShow)
{
	return ::ShowScrollBar(hwnd, fnBar, fShow);
}

BOOL CTxtWinHost::TxEnableScrollBar (INT fuSBFlags, INT fuArrowflags)
{
	return ::EnableScrollBar(hwnd, fuSBFlags, fuArrowflags) ;//SB_HORZ, ESB_DISABLE_BOTH);
}


BOOL CTxtWinHost::TxSetScrollRange(INT fnBar, LONG nMinPos, INT nMaxPos, BOOL fRedraw)
{
	return ::SetScrollRange(hwnd, fnBar, nMinPos, nMaxPos, fRedraw);
}


BOOL CTxtWinHost::TxSetScrollPos (INT fnBar, INT nPos, BOOL fRedraw)
{
	return ::SetScrollPos(hwnd, fnBar, nPos, fRedraw);
}

void CTxtWinHost::TxInvalidateRect(LPCRECT prc, BOOL fMode)
{
	::InvalidateRect(hwnd, prc, fMode);
}

void CTxtWinHost::TxViewChange(BOOL fUpdate) 
{
	::UpdateWindow (hwnd);
}


BOOL CTxtWinHost::TxCreateCaret(HBITMAP hbmp, INT xWidth, INT yHeight)
{
	return ::CreateCaret (hwnd, hbmp, xWidth, yHeight);
}


BOOL CTxtWinHost::TxShowCaret(BOOL fShow)
{
	if(fShow)
		return ::ShowCaret(hwnd);
	else
		return ::HideCaret(hwnd);
}

BOOL CTxtWinHost::TxSetCaretPos(INT x, INT y)
{
	return ::SetCaretPos(x, y);
}


BOOL CTxtWinHost::TxSetTimer(UINT idTimer, UINT uTimeout)
{
	fTimer = TRUE;
	return ::SetTimer(hwnd, idTimer, uTimeout, NULL);
}


void CTxtWinHost::TxKillTimer(UINT idTimer)
{
	::KillTimer(hwnd, idTimer);
	fTimer = FALSE;
}

void CTxtWinHost::TxScrollWindowEx (INT dx, INT dy, LPCRECT lprcScroll,	LPCRECT lprcClip,	HRGN hrgnUpdate, LPRECT lprcUpdate,	UINT fuScroll)	
{
	::ScrollWindowEx(hwnd, dx, dy, lprcScroll, lprcClip, hrgnUpdate, lprcUpdate, fuScroll);
}

void CTxtWinHost::TxSetCapture(BOOL fCapture)
{
	if (fCapture)
		::SetCapture(hwnd);
	else
		::ReleaseCapture();
}

void CTxtWinHost::TxSetFocus()
{
	::SetFocus(hwnd);
}

void CTxtWinHost::TxSetCursor(HCURSOR hcur,	BOOL fText)
{
	::SetCursor(hcur);
}

BOOL CTxtWinHost::TxScreenToClient(LPPOINT lppt)
{
	return ::ScreenToClient(hwnd, lppt);	
}

BOOL CTxtWinHost::TxClientToScreen(LPPOINT lppt)
{
	return ::ClientToScreen(hwnd, lppt);
}

HRESULT CTxtWinHost::TxActivate(LONG *plOldState)
{
    return S_OK;
}

HRESULT CTxtWinHost::TxDeactivate(LONG lNewState)
{
    return S_OK;
}
    

HRESULT CTxtWinHost::TxGetClientRect(LPRECT prc)
{
	*prc = rcClient;

	GetControlRect(prc);

	return NOERROR;
}


HRESULT CTxtWinHost::TxGetViewInset(LPRECT prc) 
{

    *prc = rcViewInset;
    
    return NOERROR;	
}

HRESULT CTxtWinHost::TxGetCharFormat(const CHARFORMATW **ppCF)
{
	*ppCF = &cf;
	return NOERROR;
}

HRESULT CTxtWinHost::TxGetParaFormat(const PARAFORMAT **ppPF)
{
	*ppPF = &pf;
	return NOERROR;
}


COLORREF CTxtWinHost::TxGetSysColor(int nIndex) 
{
	if (nIndex == COLOR_WINDOW)
	{
		if(!fNotSysBkgnd)
			return GetSysColor(COLOR_WINDOW);
		return crBackground;
	}

	return GetSysColor(nIndex);
}



HRESULT CTxtWinHost::TxGetBackStyle(TXTBACKSTYLE *pstyle)
{
	*pstyle = !fTransparent ? TXTBACK_OPAQUE : TXTBACK_TRANSPARENT;
	return NOERROR;
}


HRESULT CTxtWinHost::TxGetMaxLength(DWORD *pLength)
{
	*pLength = cchTextMost;
	return NOERROR;
}



HRESULT CTxtWinHost::TxGetScrollBars(DWORD *pdwScrollBar)
{
	*pdwScrollBar =  dwStyle & (WS_VSCROLL | WS_HSCROLL | ES_AUTOVSCROLL | 
						ES_AUTOHSCROLL | ES_DISABLENOSCROLL);

	return NOERROR;
}


HRESULT CTxtWinHost::TxGetPasswordChar(TCHAR *pch)
{
#ifdef UNICODE
	*pch = chPasswordChar;
#else
	WideCharToMultiByte(CP_ACP, 0, &chPasswordChar, 1, pch, 1, NULL, NULL) ;
#endif
	return NOERROR;
}

HRESULT CTxtWinHost::TxGetAcceleratorPos(LONG *pcp)
{
	*pcp = laccelpos;
	return S_OK;
} 										   

HRESULT CTxtWinHost::OnTxCharFormatChange(const CHARFORMATW *pcf)
{
	return S_OK;
}


HRESULT CTxtWinHost::OnTxParaFormatChange(const PARAFORMAT *ppf)
{
	pf = *ppf;
	return S_OK;
}


HRESULT CTxtWinHost::TxGetPropertyBits(DWORD dwMask, DWORD *pdwBits) 
{
	DWORD dwProperties = 0;

	if (fRich)
	{
		dwProperties = TXTBIT_RICHTEXT;
	}

	if (dwStyle & ES_MULTILINE)
	{
		dwProperties |= TXTBIT_MULTILINE;
	}

	if (dwStyle & ES_READONLY)
	{
		dwProperties |= TXTBIT_READONLY;
	}


	if (dwStyle & ES_PASSWORD)
	{
		dwProperties |= TXTBIT_USEPASSWORD;
	}

	if (!(dwStyle & ES_NOHIDESEL))
	{
		dwProperties |= TXTBIT_HIDESELECTION;
	}

	if (fEnableAutoWordSel)
	{
		dwProperties |= TXTBIT_AUTOWORDSEL;
	}

	if (fVertical)
	{
		dwProperties |= TXTBIT_VERTICAL;
	}
					
	if (fWordWrap)
	{
		dwProperties |= TXTBIT_WORDWRAP;
	}

	if (fAllowBeep)
	{
		dwProperties |= TXTBIT_ALLOWBEEP;
	}

	if (fSaveSelection)
	{
		dwProperties |= TXTBIT_SAVESELECTION;
	}

	*pdwBits = dwProperties & dwMask; 
	return NOERROR;
}


HRESULT CTxtWinHost::TxNotify(DWORD iNotify, void *pv)
{
	if( iNotify == EN_REQUESTRESIZE )
	{
		RECT rc;
		REQRESIZE *preqsz = (REQRESIZE *)pv;
		
		GetControlRect(&rc);
		rc.bottom = rc.top + preqsz->rc.bottom + HOST_BORDER;
		rc.right  = rc.left + preqsz->rc.right + HOST_BORDER;
		rc.top -= HOST_BORDER;
		rc.left -= HOST_BORDER;
		
		SetClientRect(&rc, TRUE);
		
		return S_OK;
	} 

	// Forward this to the container
	if (pnc)
	{
		(*pnc)(iNotify);
	}

	return S_OK;
}



HRESULT CTxtWinHost::TxGetExtent(LPSIZEL lpExtent)
{

	// Calculate the length & convert to himetric
	*lpExtent = sizelExtent;

	return S_OK;
}

HRESULT	CTxtWinHost::TxGetSelectionBarWidth (LONG *plSelBarWidth)
{
	*plSelBarWidth = lSelBarWidth;
	return S_OK;
}


BOOL CTxtWinHost::GetReadOnly()
{
	return (dwStyle & ES_READONLY) != 0;
}

void CTxtWinHost::SetReadOnly(BOOL fReadOnly)
{
	if (fReadOnly)
	{
		dwStyle |= ES_READONLY;
	}
	else
	{
		dwStyle &= ~ES_READONLY;
	}

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_READONLY, 
		fReadOnly ? TXTBIT_READONLY : 0);
}

BOOL CTxtWinHost::GetAllowBeep()
{
	return fAllowBeep;
}

void CTxtWinHost::SetAllowBeep(BOOL fAllowBeep)
{
	fAllowBeep = fAllowBeep;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_ALLOWBEEP, 
		fAllowBeep ? TXTBIT_ALLOWBEEP : 0);
}

void CTxtWinHost::SetViewInset(RECT *prc)
{
	rcViewInset = *prc;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_VIEWINSETCHANGE, 0);
}

WORD CTxtWinHost::GetDefaultAlign()
{
	return pf.wAlignment;
}


void CTxtWinHost::SetDefaultAlign(WORD wNewAlign)
{
	pf.wAlignment = wNewAlign;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_PARAFORMATCHANGE, 0);
}

BOOL CTxtWinHost::GetRichTextFlag()
{
	return fRich;
}

void CTxtWinHost::SetRichTextFlag(BOOL fNew)
{
	fRich = fNew;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_RICHTEXT, 
		fNew ? TXTBIT_RICHTEXT : 0);
}

LONG CTxtWinHost::GetDefaultLeftIndent()
{
	return pf.dxOffset;
}


void CTxtWinHost::SetDefaultLeftIndent(LONG lNewIndent)
{
	pf.dxOffset = lNewIndent;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_PARAFORMATCHANGE, 0);
}

void CTxtWinHost::SetClientRect(RECT *prc, BOOL fUpdateExtent) 
{
	// If the extent matches the client rect then we assume the extent should follow
	// the client rect.
	LONG lTestExt = DYtoHimetricY(
		(rcClient.bottom - rcClient.top)  - 2 * HOST_BORDER, yPerInch);

	if (fUpdateExtent 
		&& (sizelExtent.cy == lTestExt))
	{
		sizelExtent.cy = DXtoHimetricX((prc->bottom - prc->top) - 2 * HOST_BORDER, 
			xPerInch);
		sizelExtent.cx = DYtoHimetricY((prc->right - prc->left) - 2 * HOST_BORDER,
			yPerInch);
	}

	rcClient = *prc; 
}

BOOL CTxtWinHost::SetSaveSelection(BOOL f_SaveSelection)
{
	BOOL fResult = f_SaveSelection;

	fSaveSelection = f_SaveSelection;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_SAVESELECTION, 
		fSaveSelection ? TXTBIT_SAVESELECTION : 0);

	return fResult;		
}

HRESULT	CTxtWinHost::OnTxInPlaceDeactivate()
{
	HRESULT hr = pserv->OnTxInPlaceDeactivate();

	if (SUCCEEDED(hr))
	{
		fInplaceActive = FALSE;
	}

	return hr;
}

HRESULT	CTxtWinHost::OnTxInPlaceActivate(LPCRECT prcClient)
{
	fInplaceActive = TRUE;

	HRESULT hr = pserv->OnTxInPlaceActivate(prcClient);

	if (FAILED(hr))
	{
		fInplaceActive = FALSE;
	}

	return hr;
}

BOOL CTxtWinHost::DoSetCursor(RECT *prc, POINT *pt)
{
	RECT rc = prc ? *prc : rcClient;

	// Give some space for our border
	rc.top += HOST_BORDER;
	rc.bottom -= HOST_BORDER;
	rc.left += HOST_BORDER;
	rc.right -= HOST_BORDER;

	// Is this in our rectangle?
	if (PtInRect(&rc, *pt))
	{
		RECT *prcClient = (!fInplaceActive || prc) ? &rc : NULL;

		HDC hdc = GetDC(hwnd);

		pserv->OnTxSetCursor(
			DVASPECT_CONTENT,	
			-1,
			NULL,
			NULL,
			hdc,
			NULL,
			prcClient,
			pt->x, 
			pt->y);

		ReleaseDC(hwnd, hdc);

		return TRUE;
	}

	return FALSE;
}

void CTxtWinHost::GetControlRect(
	LPRECT prc			//@parm	Where to put client coordinates
)
{
	// Give some space for our border
	prc->top = rcClient.top + HOST_BORDER;
	prc->bottom = rcClient.bottom - HOST_BORDER;
	prc->left = rcClient.left + HOST_BORDER;
	prc->right = rcClient.right - HOST_BORDER;
}

void CTxtWinHost::SetTransparent(BOOL f_Transparent)
{
	fTransparent = f_Transparent;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_BACKSTYLECHANGE, 0);
}

LONG CTxtWinHost::SetAccelPos(LONG l_accelpos)
{
	LONG laccelposOld = l_accelpos;

	laccelpos = l_accelpos;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_SHOWACCELERATOR, 0);

	return laccelposOld;
}

WCHAR CTxtWinHost::SetPasswordChar(WCHAR ch_PasswordChar)
{
	WCHAR chOldPasswordChar = chPasswordChar;

	chPasswordChar = ch_PasswordChar;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_USEPASSWORD, 
		(chPasswordChar != 0) ? TXTBIT_USEPASSWORD : 0);

	return chOldPasswordChar;
}

void CTxtWinHost::SetDisabled(BOOL fOn)
{
	cf.dwMask	  |= CFM_COLOR	   | CFM_DISABLED;
	cf.dwEffects |= CFE_AUTOCOLOR | CFE_DISABLED;

	if( !fOn )
	{
		cf.dwEffects &= ~CFE_DISABLED;
	}
	
	pserv->OnTxPropertyBitsChange(TXTBIT_CHARFORMATCHANGE, 
		TXTBIT_CHARFORMATCHANGE);
}

LONG CTxtWinHost::SetSelBarWidth(LONG l_SelBarWidth)
{
	LONG lOldSelBarWidth = lSelBarWidth;

	lSelBarWidth = l_SelBarWidth;

	if (lSelBarWidth)
	{
		dwStyle |= ES_SELECTIONBAR;
	}
	else
	{
		dwStyle &= (~ES_SELECTIONBAR);
	}

	pserv->OnTxPropertyBitsChange(TXTBIT_SELBARCHANGE, TXTBIT_SELBARCHANGE);

	return lOldSelBarWidth;
}

BOOL CTxtWinHost::GetTimerState()
{
	return fTimer;
}