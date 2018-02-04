// PDMCons.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include "PDMCons.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 唯一的应用程序对象

CWinApp theApp;

using namespace std;

int main()
{
    int nRetCode = 0;

    HMODULE hModule = ::GetModuleHandle(nullptr);
	CoInitialize(0);

    if (hModule != nullptr)
    {
        // 初始化 MFC 并在失败时显示错误
        if (!AfxWinInit(hModule, nullptr, ::GetCommandLine(), 0))
        {
            // TODO: 更改错误代码以符合您的需要
            wprintf(L"错误: MFC 初始化失败\n");
            nRetCode = 1;
        }
        else
        {
            // TODO: 在此处为应用程序的行为编写代码。
        }
    }
    else
    {
        // TODO: 更改错误代码以符合您的需要
        wprintf(L"错误: GetModuleHandle 失败\n");
        nRetCode = 1;
    }
	



	IEdmVault5Ptr poVault;
	HRESULT hRes = poVault.CreateInstance(__uuidof(EdmVault5), NULL);
	try
	{
		//Create a vault interface.
		if (FAILED(hRes))
			_com_issue_error(hRes);

		//Log in on the vault.
		poVault->LoginAuto("wisevault0", (long)m_hWnd);

		//Get a pointer to the root folder.
		IEdmFolder5Ptr poFolder;
		poFolder = poVault->RootFolder;

		//Get position of first file in the folder.
		IEdmPos5Ptr poPos;
		poPos = poFolder->GetFirstFilePosition();

		CString oMessage;
		if (poPos->IsNull)
			oMessage = "The root folder of your vault does not contain any files.";
		else
			oMessage = "The root folder of your vault contains these files:\n";

		while (poPos->IsNull == VARIANT_FALSE)
		{
			IEdmFile5Ptr poFile = poFolder->GetNextFile(poPos);
			oMessage += (LPCTSTR)poFile->GetName();
			oMessage += "\n";
		}

		AfxMessageBox(oMessage);
	}
	catch (const _com_error &roError)
	{
		//We get here if one of the methods above failed.
		if (poVault == NULL)
		{
			AfxMessageBox(_T("Could not create vault interface."));
		}
		else
		{
			BSTR bsName = NULL;
			BSTR bsDesc = NULL;
			poVault->GetErrorString((long)roError.Error(), &bsName, &bsDesc);
			bstr_t oName(bsName, false);
			bstr_t oDesc(bsDesc, false);

			bstr_t oMsg = "Something went wrong.\n";
			oMsg += oName;
			oMsg += "\n";
			oMsg += oDesc;
			AfxMessageBox(oMsg);
		}
	}





	CoUninitialize();
    return nRetCode;
}
