// dllmain.cpp : DllMain 的实现。

#include "stdafx.h"
#include "resource.h"
#include "WiseExportExcel_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CWiseExportExcelModule _AtlModule;

class CWiseExportExcelApp : public CWinApp
{
public:

// 重写
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CWiseExportExcelApp, CWinApp)
END_MESSAGE_MAP()

CWiseExportExcelApp theApp;

BOOL CWiseExportExcelApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	return CWinApp::InitInstance();
}

int CWiseExportExcelApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
