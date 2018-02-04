// MyAddin.h : CMyAddin 的声明

#pragma once
#include "resource.h"       // 主符号



#include "SWPDMATL_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;


// CMyAddin

class ATL_NO_VTABLE CMyAddin :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMyAddin, &CLSID_MyAddin>,
	public IMyAddin,
	public IEdmAddIn5
{
public:
	CMyAddin()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_MYADDIN)


	BEGIN_COM_MAP(CMyAddin)
		COM_INTERFACE_ENTRY(IMyAddin)
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
		//The AFX_MANAGE_STATE macro is needed for MFC applications but should not 
		//be used for applications that are MFC-free
		AFX_MANAGE_STATE(AfxGetStaticModuleState());

		if (poInfo == NULL || poCmdMgr == NULL )
			return E_POINTER;

		//Return some information to the Properties dialog box
		poInfo->mbsAddInName= SysAllocString( L"MyButton" );
		poInfo->mbsCompany = SysAllocString( L"WISE" );
		poInfo->mbsDescription= SysAllocString( L"This is a very nice add-in." );
		poInfo->mlAddInVersion = 2;

		//SOLIDWORKS PDM Professional 5.2 is required by this add-in
		poInfo->mlRequiredVersionMajor = 5;
		poInfo->mlRequiredVersionMinor= 2;

		//Add hooks and menu commands to SOLIDWORKS PDM Professional
		//Below is a menu command that appears in the Tools
		//and context-sensitive menus of a vault in Windows Explorer 
		poCmdMgr->AddCmd( 1, bstr_t("menu-command"), EdmMenu_Nothing, bstr_t(""), bstr_t(""), 0, 0 );
		//Notify the add-in when a file data card button is clicked
		IDispatch * poReserve = NULL;
		poCmdMgr->AddHook(EdmCmd_CardButton,poReserve);
		return S_OK;

	}
	STDMETHOD(OnCmd)(EdmCmd * poCmd, SAFEARRAY * * ppoData)
	{
		//The AFX_MANAGE_STATE macro is needed for MFC applications, but should not 
		//be used for applications that are MFC-free
		AFX_MANAGE_STATE(AfxGetStaticModuleState());

		CStdioFile file(L"D:\\a.txt",CFile::modeWrite|CFile::modeCreate);

		if (poCmd == NULL ||ppoData == NULL)
			return E_POINTER;
		//MessageBox((HWND)poCmd->mlParentWnd, L"Hello World!", L"SOLIDWORKS PDM Professional", MB_OK );
		file.WriteString(L"test begin\n");
		//return S_OK;

		//Respond only to a specific button command
		//The button command to respond to begins with "MyButton:" and ends with the name of the 
		//variable to update in the card 
		CString msbComment = poCmd->mbsComment;
		file.WriteString(msbComment);
		if (msbComment.Left(9) == "MyButton:")
		{
			//Get the name of the variable to update. 
			CString VarName = L"Select File for ";

			VarName =VarName+ msbComment.Right(9);//Strings.Right(poCmd.mbsComment, Strings.Len(poCmd.mbsComment) - 9);

			//Let the user select the file whose path will be copied to the card variable
			
			IEdmVault5 * vault = (IEdmVault5*)poCmd->mpoVault;//EdmVault5 default(EdmVault5);
			//vault = (EdmVault5)poCmd.mpoVault;
			IEdmStrLst5* PathList = NULL;
			vault->BrowseForFile(poCmd->mlParentWnd,EdmBws_ForOpen+EdmBws_PermitVaultFiles,L"",L"",L"",L"",VarName.AllocSysString(),&PathList);
			if ((PathList != NULL))
			{
				BSTR path = NULL;
				IEdmPos5 * headPos = NULL;
				PathList->GetHeadPosition(&headPos);
				PathList->GetNext(headPos,&path);

				//Store the path in the card variable 
				//IEdmEnumeratorVariable5 vars = default(IEdmEnumeratorVariable5);
				//vars = (IEdmEnumeratorVariable5)poCmd.mpoExtra;
				//object VariantPath = null;
				//VariantPath = path;
				//vars.SetVar(VarName, "", VariantPath);
			}
		}
		file.Flush();
		file.Close();
		return S_OK;


	}
};

OBJECT_ENTRY_AUTO(__uuidof(MyAddin), CMyAddin)
