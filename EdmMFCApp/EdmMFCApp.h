
// EdmMFCApp.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CEdmMFCAppApp:
// �йش����ʵ�֣������ EdmMFCApp.cpp
//

class CEdmMFCAppApp : public CWinApp
{
public:
	CEdmMFCAppApp();

// ��д
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CEdmMFCAppApp theApp;