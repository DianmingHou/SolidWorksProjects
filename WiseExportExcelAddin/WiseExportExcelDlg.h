#pragma once


// CWiseExportExcelDlg 对话框

class CWiseExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CWiseExportExcelDlg)

public:
	CWiseExportExcelDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CWiseExportExcelDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_EXPORTEXCEL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	// 文件路径
	CString m_strFilePath;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedButtonFilePath();
	void SetVault(IEdmVault5Ptr poVault);
	IEdmVault5Ptr GetVault();
private:
	IEdmVault5Ptr m_poVault;
public:
	CString m_strFileSelect;
	void SetSelectedFile(CString iFileName);
	virtual BOOL OnInitDialog();
private:
	IEdmFile5Ptr m_poFile;
public:
	// 设置数据库文件
	void SetVaultFile(IEdmFile5Ptr piVault, CString iFileName);
};
