// PDMCons.cpp : �������̨Ӧ�ó������ڵ㡣
//

#include "stdafx.h"
#include "PDMCons.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Ψһ��Ӧ�ó������

CWinApp theApp;

using namespace std;

int main()
{
    int nRetCode = 0;

    HMODULE hModule = ::GetModuleHandle(nullptr);
	CoInitialize(0);

    if (hModule != nullptr)
    {
        // ��ʼ�� MFC ����ʧ��ʱ��ʾ����
        if (!AfxWinInit(hModule, nullptr, ::GetCommandLine(), 0))
        {
            // TODO: ���Ĵ�������Է���������Ҫ
            wprintf(L"����: MFC ��ʼ��ʧ��\n");
            nRetCode = 1;
        }
        else
        {
            // TODO: �ڴ˴�ΪӦ�ó������Ϊ��д���롣
        }
    }
    else
    {
        // TODO: ���Ĵ�������Է���������Ҫ
        wprintf(L"����: GetModuleHandle ʧ��\n");
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
