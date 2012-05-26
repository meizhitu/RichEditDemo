
#pragma once

class CBitmapDC
{
public:

    CBitmapDC();
    virtual ~CBitmapDC();

	void Delete();
	void Create(int nWidth, int nHeight);

	HDC GetSafeHdc(void) { return m_hDC; };
	HBITMAP GetBmpHandle(void) { return m_hBmp; };
	DWORD* GetBits(void) { return (DWORD *)m_pBits; };
	bool IsReady();
	CRect GetDcRect() { return m_DcRect; };

protected:

	HBITMAP m_hBmp;
	HDC m_hDC;
	unsigned char* m_pBits;
	CRect m_DcRect;
};
