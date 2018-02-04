// dllmain.cpp : DllMain 的实现。

#include "stdafx.h"
#include "resource.h"
#include "SWPDMATL_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CSWPDMATLModule _AtlModule;

class CSWPDMATLApp : public CWinApp
{
public:

// 重写
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CSWPDMATLApp, CWinApp)
END_MESSAGE_MAP()

CSWPDMATLApp theApp;

BOOL CSWPDMATLApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	return CWinApp::InitInstance();
}

int CSWPDMATLApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
