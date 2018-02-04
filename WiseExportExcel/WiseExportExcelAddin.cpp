// WiseExportExcelAddin.cpp : CWiseExportExcelAddin 的实现

#include "stdafx.h"
#include "WiseExportExcelAddin.h"


// CWiseExportExcelAddin

// CWiseExportExcel

HRESULT _stdcall CWiseExportExcel::GetAddInInfo(EdmAddInInfo * poInfo, IEdmVault5 * poVault, IEdmCmdMgr5 * poCmdMgr)
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

HRESULT _stdcall CWiseExportExcel::OnCmd(EdmCmd * poCmd, SAFEARRAY * * ppoData)
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