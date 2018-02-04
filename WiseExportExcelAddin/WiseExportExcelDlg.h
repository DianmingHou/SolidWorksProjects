#pragma once


// CWiseExportExcelDlg �Ի���

class CWiseExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CWiseExportExcelDlg)

public:
	CWiseExportExcelDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CWiseExportExcelDlg();

// �Ի�������
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
private:
	CString m_strFileSelect;
	CString m_strFilePath;
	IEdmVault5Ptr m_poVault;
	IEdmFile5Ptr m_poFile;
public:
	// �������ݿ��ļ�
	void SetVaultFile(IEdmFile5Ptr piVault, CString iFileName);
	void SetVault(IEdmVault5Ptr poVault);
	IEdmVault5Ptr GetVault();
	void SetSelectedFile(CString iFileName);
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedButtonOutputFile();
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeEditOutputFile();
};
