// DynamicOleCom.cpp : Implementation of CDynamicOleCom

#include "stdafx.h"
#include "DynamicOleCom.h"


// CDynamicOleCom

STDMETHODIMP CDynamicOleCom::InsertGif(LONG img)
{
	return S_OK;
}

HRESULT CDynamicOleCom::OnDraw(ATL_DRAWINFO& di)
{
	RECT& rc = *(RECT*)di.prcBounds;
	//Rectangle(di.hdcDraw, rc.left+1, rc.top+1, rc.right-1, rc.bottom-1);
	DrawText(di.hdcDraw, _T("ATL 2.0"), -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
	return S_OK;
}