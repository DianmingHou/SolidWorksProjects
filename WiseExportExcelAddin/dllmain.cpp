// dllmain.cpp: DllMain ��ʵ�֡�

#include "stdafx.h"
#include "resource.h"
#include "WiseExportExcelAddin_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CWiseExportExcelAddinModule _AtlModule;

class CWiseExportExcelAddinApp : public CWinApp
{
public:

// ��д
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CWiseExportExcelAddinApp, CWinApp)
END_MESSAGE_MAP()

CWiseExportExcelAddinApp theApp;

BOOL CWiseExportExcelAddinApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	//AfxOleInit();
	return CWinApp::InitInstance();
}

int CWiseExportExcelAddinApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
