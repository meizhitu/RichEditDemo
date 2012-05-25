// DynamicOleCom.cpp : Implementation of CDynamicOleCom

#include "stdafx.h"
#include "DynamicOleCom.h"


// CDynamicOleCom

STDMETHODIMP CDynamicOleCom::InsertGif(LONG img)
{
	m_gifImg = (Image*)img;
	m_objectType = 2;
	::SetTimer( m_mockWindow.m_hWnd, (UINT_PTR)this , 100 , OnGifTimer );
	return S_OK;
}

STDMETHODIMP CDynamicOleCom::SetHostWindow(LONG hWnd)
{
	m_hostWnd = (HWND)hWnd;
	return S_OK;
}

HRESULT CDynamicOleCom::OnDraw(ATL_DRAWINFO& di)
{
	switch(m_objectType)
	{
	case 2:
		{
			Bitmap bmp(di.prcBounds->right - di.prcBounds->left, di.prcBounds->bottom - di.prcBounds->top);
			Graphics *pMemG = Graphics::FromImage(&bmp);
			pMemG->DrawImage(m_gifImg,0, 0);
			Graphics  g(di.hdcDraw);
			RectF destRect(di.prcBounds->left, di.prcBounds->top, di.prcBounds->right - di.prcBounds->left, di.prcBounds->bottom - di.prcBounds->top);
			g.DrawImage(&bmp, destRect);
			delete pMemG;	
		}
		break;
	}

	return S_OK;
}

void CDynamicOleCom::RedrawHostWindow()
{
	//如果用FireViewChange，会有滚动条的问题
	if (m_hostWnd && ::IsWindow(m_hostWnd))
	{
		::RedrawWindow(m_hostWnd,NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
	}
}


VOID CDynamicOleCom::OnGifTimer( HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime )
{
	::KillTimer( hwnd, idEvent );
	CDynamicOleCom* dynOleCom = ( CDynamicOleCom* )idEvent;

	GUID pageGUID = Gdiplus::FrameDimensionTime;
	
	Image* img = dynOleCom->m_gifImg;
	img->SelectActiveFrame( &pageGUID , dynOleCom->m_currentFrame );
	dynOleCom->m_currentFrame = (dynOleCom->m_currentFrame+1)%48;
	::SetTimer( hwnd, idEvent , 100 , OnGifTimer );
	//dynOleCom->RedrawHostWindow();
	//GIF解析
	/*int nFrameDimensionsCount = img->GetFrameDimensionsCount();
	if( nFrameDimensionsCount > 0 )
	{
		GUID *pDimensionIDs = new GUID[ nFrameDimensionsCount ];
		img->GetFrameDimensionsList( pDimensionIDs , nFrameDimensionsCount );
		int nFrameCount = img->GetFrameCount( &pDimensionIDs[0] );
		delete []pDimensionIDs;
		if( nFrameCount > 1 )
		{
			int nPropSize = img->GetPropertyItemSize( PropertyTagFrameDelay );
			if( nPropSize > 0 )
			{
				Gdiplus::PropertyItem *propItem = (Gdiplus::PropertyItem *)malloc( nPropSize ) ;
				if( propItem != NULL )
				{
					img->GetPropertyItem( PropertyTagFrameDelay , nPropSize , propItem );
					long lPause = ((long*) propItem->value)[dynOleCom->m_currentFrame] * 10;
					if( lPause < 100 ) lPause = 100;
					dynOleCom->m_currentFrame = (dynOleCom->m_currentFrame+1)%nFrameCount;
					::SetTimer( hwnd, idEvent , lPause , OnGifTimer );
					free( propItem );
				}			
			}			
		}
	}*/

}


LRESULT CWinHiddenMock::OnRefreshOle(UINT uMsg, WPARAM wParam, 
									 LPARAM lParam, BOOL& bHandled)
{
	if (m_dynamicOleCom != NULL)
		m_dynamicOleCom->FireViewChange();//调用控件对象的函数，在这个函数中实现我们要的功能
	return 0;
}

void CWinHiddenMock::AttachOle(CDynamicOleCom* ctl)
{
	m_dynamicOleCom = ctl;
}