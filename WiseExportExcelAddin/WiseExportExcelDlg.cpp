// WiseExportExcelDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "WiseExportExcelDlg.h"
#include "afxdialogex.h"


// CWiseExportExcelDlg �Ի���

IMPLEMENT_DYNAMIC(CWiseExportExcelDlg, CDialogEx)

CWiseExportExcelDlg::CWiseExportExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG_EXPORTEXCEL, pParent)
	, m_strFilePath(_T(""))
	, m_strFileSelect(_T(""))
{

}

CWiseExportExcelDlg::~CWiseExportExcelDlg()
{
}

void CWiseExportExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_FILE_PATH, m_strFilePath);
	DDX_Text(pDX, IDC_EDIT_FILE_SELECT, m_strFileSelect);
}


BEGIN_MESSAGE_MAP(CWiseExportExcelDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CWiseExportExcelDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_FILE_PATH, &CWiseExportExcelDlg::OnBnClickedButtonFilePath)
END_MESSAGE_MAP()


// CWiseExportExcelDlg ��Ϣ�������


void CWiseExportExcelDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	IEdmEnumeratorVariable5Ptr enumVarPtr = NULL;
	m_poFile->GetEnumeratorVariable(m_strFileSelect.AllocSysString(), &enumVarPtr);
	VARIANT vParam;
	CString varStr = L"";
	char *aar[5] = { "Number","Weight","Material","Author","У��" };
	for (int i = 0; i < 5; i++) {
		VariantInit(&vParam);//��һ�������
		VARIANT_BOOL isSuccess = VARIANT_FALSE;
		enumVarPtr->GetVar(_bstr_t(aar[i]), L"@", &vParam, &isSuccess);
		varStr += aar[i];
		varStr += L"  == ";
		varStr += vParam.bstrVal;
		varStr += "\n";
		VariantClear(&vParam);
	}
	
	AfxMessageBox(varStr);


	CDialogEx::OnOK();
}


void CWiseExportExcelDlg::OnBnClickedButtonFilePath()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Describe Files (*.cfg)|*.cfg|All Files (*.*)|*.*||"), NULL);

	if (dlgFile.DoModal())
	{
		//AfxMessageBox(dlgFile.GetPathName() + L"   " + dlgFile.GetFolderPath() + L"  " + dlgFile.GetFileName());
		m_strFilePath = dlgFile.GetPathName();
		UpdateData(false);
	}

}


void CWiseExportExcelDlg::SetVault(IEdmVault5Ptr poVault)
{
	m_poVault = poVault;
}


IEdmVault5Ptr CWiseExportExcelDlg::GetVault()
{
	return m_poVault;
}


void CWiseExportExcelDlg::SetSelectedFile(CString iFileName)
{
	m_strFileSelect = iFileName;
}


BOOL CWiseExportExcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
				  // �쳣: OCX ����ҳӦ���� FALSE
}


// �������ݿ��ļ�
void CWiseExportExcelDlg::SetVaultFile(IEdmFile5Ptr piVault, CString iFileName)
{
	m_poFile = piVault;
	m_strFileSelect = iFileName;
}
