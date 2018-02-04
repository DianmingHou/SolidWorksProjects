#pragma once


// ExportExcelDlg �Ի���

class ExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ExportExcelDlg)

public:
	ExportExcelDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~ExportExcelDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_EXPORT_EXCEL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
private:
	// ѡ���ļ���
	CString m_strFileName;
	// �����ļ�·��������
	CString m_strExportFileName;
public:
	afx_msg void OnBnClickedButtonPath();
	afx_msg void OnBnClickedOk();
private:
	// vault�ֿ�
	IEdmVault5Ptr m_poVault;
public:
	// ����vault
	void SetVault(IEdmVault5Ptr poVault);
	// ��ȡVault
	IEdmVault5Ptr GetVault();
};
