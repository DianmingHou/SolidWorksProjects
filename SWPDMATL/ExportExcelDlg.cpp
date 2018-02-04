// ExportExcelDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExportExcelDlg.h"
#include "afxdialogex.h"


// ExportExcelDlg 对话框

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


// ExportExcelDlg 消息处理程序


void ExportExcelDlg::OnBnClickedButtonPath()
{
	// TODO: 在此添加控件通知处理程序代码
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Describe Files (*.xls,xlsx)|*.xls|*.xlsx|All Files (*.*)|*.*||"), NULL);

	if (dlgFile.DoModal())
	{
		m_strExportFileName = dlgFile.GetPathName();
	}
}


void ExportExcelDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();
}


// 设置vault
void ExportExcelDlg::SetVault(IEdmVault5Ptr poVault)
{
	m_poVault = poVault;
}


// 获取Vault
IEdmVault5Ptr ExportExcelDlg::GetVault()
{
	return m_poVault;
}
