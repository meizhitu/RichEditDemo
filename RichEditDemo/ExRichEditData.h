#pragma once


enum ExDataType
{
	IMAGE = 1,
	GIF = 2
};

class CExRichEditData
{
public:
	ExDataType m_dataType;
	void* m_data;
};