#pragma once


// ExportExcelDlg 对话框

class ExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ExportExcelDlg)

public:
	ExportExcelDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~ExportExcelDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_EXPORT_EXCEL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
private:
	// 选中文件名
	CString m_strFileName;
	// 导出文件路径和名称
	CString m_strExportFileName;
public:
	afx_msg void OnBnClickedButtonPath();
	afx_msg void OnBnClickedOk();
private:
	// vault仓库
	IEdmVault5Ptr m_poVault;
public:
	// 设置vault
	void SetVault(IEdmVault5Ptr poVault);
	// 获取Vault
	IEdmVault5Ptr GetVault();
};
