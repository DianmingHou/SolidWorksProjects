#pragma once


// CWiseExportExcelDlg �Ի���

class CWiseExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CWiseExportExcelDlg)

public:
	CWiseExportExcelDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CWiseExportExcelDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_EXPORTEXCEL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	// �ļ�·��
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
	// �������ݿ��ļ�
	void SetVaultFile(IEdmFile5Ptr piVault, CString iFileName);
};
