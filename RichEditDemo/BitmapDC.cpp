#include "stdafx.h"
#include "BitmapDC.h"

CBitmapDC::CBitmapDC()
{
    m_pBits = NULL;
    m_hBmp = NULL;
	m_hDC = NULL;
	m_DcRect.SetRectEmpty();
}

CBitmapDC::~CBitmapDC()
{
   Delete();
}

void CBitmapDC::Create(int nWidth, int nHeight)
{
	if (nWidth <= 0 || nHeight<= 0)
		return;

	if (nWidth != m_DcRect.Width() || nHeight != m_DcRect.Height())
	{
		Delete();

		BITMAPINFOHEADER bih;
		memset(&bih, 0, sizeof(BITMAPINFOHEADER));
		bih.biSize = sizeof(BITMAPINFOHEADER);
		bih.biWidth = nWidth;
		bih.biHeight = nHeight;
		bih.biPlanes = 1;
		bih.biBitCount = 32;
		bih.biCompression = BI_RGB;

		m_DcRect.right = bih.biWidth;
		m_DcRect.bottom = bih.biHeight;

		m_hDC = CreateCompatibleDC(NULL);
		if (m_hDC != NULL)
		{
			m_hBmp	= ::CreateDIBSection(GetSafeHdc(), (BITMAPINFO*)&bih,
				DIB_RGB_COLORS, (void**)(&m_pBits), NULL, 0);

			if (m_hBmp != NULL && m_pBits != NULL)
				SelectObject(m_hDC, m_hBmp);
			else
				Delete();
		}
	}
	else
	{
		if (m_pBits != NULL)
			memset(m_pBits, 0, m_DcRect.Width() * m_DcRect.Height() * 4);
	}
}

void CBitmapDC::Delete()
{
	if (m_hBmp != NULL)
	{
		DeleteObject(m_hBmp);
		m_hBmp = NULL;
	}

	if (m_hDC != NULL)
	{
		DeleteDC(m_hDC);
		m_hDC = NULL;
	}

	m_pBits = NULL;
	m_DcRect.SetRectEmpty();
}

bool CBitmapDC::IsReady()
{
	return ((m_hBmp != NULL) && (m_hDC != NULL) && (m_pBits != NULL));
}
