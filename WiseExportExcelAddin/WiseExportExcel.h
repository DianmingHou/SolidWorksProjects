// WiseExportExcel.h : CWiseExportExcel 的声明

#pragma once
#include "resource.h"       // 主符号



#include "WiseExportExcelAddin_i.h"

#include "WiseExportExcelDlg.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;


// CWiseExportExcel

class ATL_NO_VTABLE CWiseExportExcel :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CWiseExportExcel, &CLSID_WiseExportExcel>,
	public IWiseExportExcel,
	public IEdmAddIn5
{
public:
	CWiseExportExcel()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_WISEEXPORTEXCEL)


	BEGIN_COM_MAP(CWiseExportExcel)
		COM_INTERFACE_ENTRY(IWiseExportExcel)
		COM_INTERFACE_ENTRY(IEdmAddIn5)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:




	// IEdmAddIn5 Methods
public:
	STDMETHOD(GetAddInInfo)(EdmAddInInfo * poInfo, IEdmVault5 * poVault, IEdmCmdMgr5 * poCmdMgr);
	STDMETHOD(OnCmd)(EdmCmd * poCmd, SAFEARRAY * * ppoData);
};

OBJECT_ENTRY_AUTO(__uuidof(WiseExportExcel), CWiseExportExcel)
