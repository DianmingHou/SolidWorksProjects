// WiseExportExcel.h : CWiseExportExcel ������

#pragma once
#include "resource.h"       // ������



#include "WiseExportExcelAddin_i.h"

#include "WiseExportExcelDlg.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
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
