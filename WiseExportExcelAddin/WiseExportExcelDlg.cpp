// WiseExportExcelDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "WiseExportExcelDlg.h"
#include "afxdialogex.h"

#include "WiseUtility.h"
// CWiseExportExcelDlg 对话框
#include "CWorkbook.h"


IMPLEMENT_DYNAMIC(CWiseExportExcelDlg, CDialogEx)

CWiseExportExcelDlg::CWiseExportExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CWiseExportExcelDlg::IDD, pParent)
	, m_strFileSelect(_T(""))
	, m_strFilePath(_T(""))
{

}

CWiseExportExcelDlg::~CWiseExportExcelDlg()
{
}

void CWiseExportExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_SELECT_FILE, m_strFileSelect);
	DDX_Text(pDX, IDC_EDIT_OUTPUT_FILE, m_strFilePath);
}


BEGIN_MESSAGE_MAP(CWiseExportExcelDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_OUTPUT_FILE, &CWiseExportExcelDlg::OnBnClickedButtonOutputFile)
	ON_BN_CLICKED(IDOK, &CWiseExportExcelDlg::OnBnClickedOk)
	ON_EN_CHANGE(IDC_EDIT_OUTPUT_FILE, &CWiseExportExcelDlg::OnEnChangeEditOutputFile)
END_MESSAGE_MAP()


// CWiseExportExcelDlg 消息处理程序



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

	// TODO:  在此添加额外的初始化
	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


// 设置数据库文件
void CWiseExportExcelDlg::SetVaultFile(IEdmFile5Ptr piVault, CString iFileName)
{
	m_poFile = piVault;
	m_strFileSelect = iFileName;
}


void CWiseExportExcelDlg::OnBnClickedButtonOutputFile()
{
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Worksheet Files (*.xls;*.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*||"), NULL);

	if (dlgFile.DoModal())
	{
		//AfxMessageBox(dlgFile.GetPathName() + L"   " + dlgFile.GetFolderPath() + L"  " + dlgFile.GetFileName());
		m_strFilePath = dlgFile.GetPathName();
		UpdateData(false);
	}
}


void CWiseExportExcelDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码

	//获取PDMclient安装路径以及模板文件的路径
	BSTR valueData = NULL;
	WiseUtility::GetRegistryValueStr(HKEY_LOCAL_MACHINE,L"SOFTWARE\\solidworks\\Applications\\PDMWorks Enterprise\\",L"Location",&valueData);
	CString pdmLocation = valueData;
	pdmLocation+="ExportTemplate\\template-list.xlsx";
	SldAsm asmPrd;
	if (m_strFilePath=="")
	{
		AfxMessageBox(L"请填写导出文件名");
		return;
	}
	WiseUtility::GetAsmInfoFromFile(m_poVault,m_strFileSelect,asmPrd);

	WiseUtility::GetReferencedFiles(m_poVault,m_strFileSelect,NULL,0,asmPrd);


	if(true){
		//间隔，在目录列表和真是列表之间插入空行
		SldPrt * tempPrt = new SldPrt();
		tempPrt->name = "";
		asmPrd.sldPrtList.AddHead(tempPrt);
	}
	CWorkbook workbook;
	WiseUtility::GetExcelWorkbook(pdmLocation,workbook);

	//插入标记，自制件多少页，标准件多少页(28条每页)，外购件多少页(28条每页)
	WiseUtility::SaveBuyPrdToWorkbook(workbook,asmPrd);
	WiseUtility::SaveStdPrdToWorkbook(workbook,asmPrd);
	WiseUtility::SavePrdToWorkbook(workbook,asmPrd);

	WiseUtility::SaveWorkbookAsExcel(m_strFilePath,workbook);

	CDialogEx::OnOK();
}


void CWiseExportExcelDlg::OnEnChangeEditOutputFile()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	UpdateData(true);
}
