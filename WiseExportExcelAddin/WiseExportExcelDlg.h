#pragma once


// CWiseExportExcelDlg 对话框

class CWiseExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CWiseExportExcelDlg)

public:
	CWiseExportExcelDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CWiseExportExcelDlg();

// 对话框数据
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
private:
	CString m_strFileSelect;
	CString m_strFilePath;
	IEdmVault5Ptr m_poVault;
	IEdmFile5Ptr m_poFile;
public:
	// 设置数据库文件
	void SetVaultFile(IEdmFile5Ptr piVault, CString iFileName);
	void SetVault(IEdmVault5Ptr poVault);
	IEdmVault5Ptr GetVault();
	void SetSelectedFile(CString iFileName);
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedButtonOutputFile();
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeEditOutputFile();
};
