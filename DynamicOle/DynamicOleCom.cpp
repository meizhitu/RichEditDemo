// DynamicOleCom.cpp : Implementation of CDynamicOleCom

#include "stdafx.h"
#include "DynamicOleCom.h"


// CDynamicOleCom

STDMETHODIMP CDynamicOleCom::InsertGif(void)
{
	return S_OK;
}

HRESULT CDynamicOleCom::OnDraw(ATL_DRAWINFO& di)
{
	RECT& rc = *(RECT*)di.prcBounds;
	Rectangle(di.hdcDraw, rc.left, rc.top, rc.right, rc.bottom);
	DrawText(di.hdcDraw, _T("ATL 2.0"), -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
	return S_OK;
}