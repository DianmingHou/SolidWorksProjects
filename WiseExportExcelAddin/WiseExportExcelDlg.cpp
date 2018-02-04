// WiseExportExcelDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "WiseExportExcelDlg.h"
#include "afxdialogex.h"

#include "WiseUtility.h"
// CWiseExportExcelDlg �Ի���
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


// CWiseExportExcelDlg ��Ϣ�������



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


void CWiseExportExcelDlg::OnBnClickedButtonOutputFile()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	//��ȡPDMclient��װ·���Լ�ģ���ļ���·��
	BSTR valueData = NULL;
	WiseUtility::GetRegistryValueStr(HKEY_LOCAL_MACHINE,L"SOFTWARE\\solidworks\\Applications\\PDMWorks Enterprise\\",L"Location",&valueData);
	CString pdmLocation = valueData;
	pdmLocation+="ExportTemplate\\template-list.xlsx";
	SldAsm asmPrd;
	if (m_strFilePath=="")
	{
		AfxMessageBox(L"����д�����ļ���");
		return;
	}
	WiseUtility::GetAsmInfoFromFile(m_poVault,m_strFileSelect,asmPrd);

	WiseUtility::GetReferencedFiles(m_poVault,m_strFileSelect,NULL,0,asmPrd);


	if(true){
		//�������Ŀ¼�б�������б�֮��������
		SldPrt * tempPrt = new SldPrt();
		tempPrt->name = "";
		asmPrd.sldPrtList.AddHead(tempPrt);
	}
	CWorkbook workbook;
	WiseUtility::GetExcelWorkbook(pdmLocation,workbook);

	//�����ǣ����Ƽ�����ҳ����׼������ҳ(28��ÿҳ)���⹺������ҳ(28��ÿҳ)
	WiseUtility::SaveBuyPrdToWorkbook(workbook,asmPrd);
	WiseUtility::SaveStdPrdToWorkbook(workbook,asmPrd);
	WiseUtility::SavePrdToWorkbook(workbook,asmPrd);

	WiseUtility::SaveWorkbookAsExcel(m_strFilePath,workbook);

	CDialogEx::OnOK();
}


void CWiseExportExcelDlg::OnEnChangeEditOutputFile()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	UpdateData(true);
}
