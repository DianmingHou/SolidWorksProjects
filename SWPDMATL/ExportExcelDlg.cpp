// ExportExcelDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ExportExcelDlg.h"
#include "afxdialogex.h"


// ExportExcelDlg �Ի���

IMPLEMENT_DYNAMIC(ExportExcelDlg, CDialogEx)

ExportExcelDlg::ExportExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG_EXPORT_EXCEL, pParent)
	, m_strFileName(_T(""))
	, m_strExportFileName(_T(""))
{

}

ExportExcelDlg::~ExportExcelDlg()
{
}

void ExportExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_SELECT_FILE, m_strFileName);
	DDX_Text(pDX, IDC_EDIT_EXPORT_FILE, m_strExportFileName);
}


BEGIN_MESSAGE_MAP(ExportExcelDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_PATH, &ExportExcelDlg::OnBnClickedButtonPath)
	ON_BN_CLICKED(IDOK, &ExportExcelDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// ExportExcelDlg ��Ϣ�������


void ExportExcelDlg::OnBnClickedButtonPath()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Describe Files (*.xls,xlsx)|*.xls|*.xlsx|All Files (*.*)|*.*||"), NULL);

	if (dlgFile.DoModal())
	{
		m_strExportFileName = dlgFile.GetPathName();
	}
}


void ExportExcelDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CDialogEx::OnOK();
}


// ����vault
void ExportExcelDlg::SetVault(IEdmVault5Ptr poVault)
{
	m_poVault = poVault;
}


// ��ȡVault
IEdmVault5Ptr ExportExcelDlg::GetVault()
{
	return m_poVault;
}
