
// ExcelAssembly.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelAssemblyApp:
// �йش����ʵ�֣������ ExcelAssembly.cpp
//

class CExcelAssemblyApp : public CWinApp
{
public:
	CExcelAssemblyApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelAssemblyApp theApp;