// WiseExportExcelAddin.h : CWiseExportExcelAddin ������

#pragma once
#include "resource.h"       // ������



#include "WiseExportExcel_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
#endif

using namespace ATL;


// CWiseExportExcelAddin

class ATL_NO_VTABLE CWiseExportExcelAddin :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CWiseExportExcelAddin, &CLSID_WiseExportExcelAddin>,
	public IDispatchImpl<IWiseExportExcelAddin, &IID_IWiseExportExcelAddin, &LIBID_WiseExportExcelLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IEdmAddIn5
{
public:
	CWiseExportExcelAddin()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_WISEEXPORTEXCELADDIN)


	BEGIN_COM_MAP(CWiseExportExcelAddin)
		COM_INTERFACE_ENTRY(IWiseExportExcelAddin)
		COM_INTERFACE_ENTRY(IDispatch)
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

OBJECT_ENTRY_AUTO(__uuidof(WiseExportExcelAddin), CWiseExportExcelAddin)
