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
	STDMETHOD(GetAddInInfo)(EdmAddInInfo * poInfo, IEdmVault5 * poVault, IEdmCmdMgr5 * poCmdMgr)
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState());

		if (poInfo == NULL || poCmdMgr == NULL)
			return E_POINTER;

		//Return some information to the Properties dialog box
		poInfo->mbsAddInName = SysAllocString(L"WiseExportExcelAddin");
		poInfo->mbsCompany = SysAllocString(L"WISE Corp.");
		poInfo->mbsDescription = SysAllocString(L"This Add-in export Bill Material to an excel file.");
		poInfo->mlAddInVersion = 1;

		//SOLIDWORKS PDM Professional 5.2 is required by this add-in
		poInfo->mlRequiredVersionMajor = 5;
		poInfo->mlRequiredVersionMinor = 2;

		//Add hooks and menu commands to SOLIDWORKS PDM Professional
		//Below is a menu command that appears in the Tools
		//and context-sensitive menus of a vault in Windows Explorer 
		poCmdMgr->AddCmd(1001, bstr_t("导出到Excel"), EdmMenu_OnlyFiles + EdmMenu_OnlySingleSelection + EdmMenu_ShowInMenuBarTools, bstr_t(""), bstr_t(""), 0, 0);

		return S_OK;
	}
	STDMETHOD(OnCmd)(EdmCmd * poCmd, SAFEARRAY * * ppoData)
	{
		//The AFX_MANAGE_STATE macro is needed for MFC applications, but should not 
		//be used for applications that are MFC-free
		AFX_MANAGE_STATE(AfxGetStaticModuleState());

		if (poCmd == NULL || ppoData == NULL)
			return E_POINTER;
		long parentFoldId = poCmd->mlCurrentFolderID;
		IEdmVault5Ptr poVault = poCmd->mpoVault;
		if (SafeArrayGetDim(*ppoData) != 1)
			return E_INVALIDARG;

		EdmCmdData *poElements = NULL;
		HRESULT hRes = SafeArrayAccessData(*ppoData, (void**)&poElements);
		if (FAILED(hRes))
			return hRes;

		//Append the paths of all files to a string
		CString oMessage = L"The following files were affected by a hook:\n";
		int iCount = (*ppoData)->rgsabound->cElements;
		//for (int i = 0; i < iCount; ++i)
		//{
		//	oMessage += (LPCTSTR)bstr_t(poElements->mbsStrData1);
		//	oMessage += "\n";
		//	++poElements;
		//}
		//MessageBox((HWND)poCmd->mlParentWnd, oMessage, L"SOLIDWORKS PDM Professional", MB_OK);

		if (poCmd->meCmdType== EdmCmd_Menu) {
			switch (poCmd->mlCmdID) {
			case 1001: {
				if (iCount > 1) {
					break;
				}
				IEdmFolder5Ptr foldPtr = NULL; 
				poVault->GetObject(EdmObject_Folder, parentFoldId,(IEdmObject5 * *)&foldPtr);
				BSTR localPath;
				foldPtr->get_LocalPath(&localPath);
				IEdmFile5Ptr fildPtr = NULL;
				CString foldPath = bstr_t(localPath);
				CString fileName =  bstr_t(poElements->mbsStrData1);
				foldPtr->GetFile(fileName.AllocSysString(),&fildPtr);
				fileName = foldPath +L"\\" +fileName;
				CWiseExportExcelDlg myDlg;
				int id = 00;
				myDlg.SetVault(poVault);
				myDlg.SetVaultFile(fildPtr,fileName);
				//myDlg.ShowWindow(SW_SHOW);
				myDlg.DoModal();
				}break;
			};
		}


		//Release the array data and display a message to the user
		SafeArrayUnaccessData(*ppoData);
		return S_OK;
	}
};

OBJECT_ENTRY_AUTO(__uuidof(WiseExportExcel), CWiseExportExcel)
