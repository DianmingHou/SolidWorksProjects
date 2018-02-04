
// EdmMFCAppDlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"


// CEdmMFCAppDlg �Ի���
class CEdmMFCAppDlg : public CDialogEx
{
// ����
public:
	CEdmMFCAppDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EDMMFCAPP_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonLogin();
	afx_msg void OnBnClickedButtonReset();
	CString vaultStr;
	CString userName;
	CString userPwd;
	CString loginState;
	CTreeCtrl m_FileTree;
	afx_msg void OnTvnSelchangedFiletree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnTvnGetInfoTipFiletree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnTvnItemexpandingFiletree(NMHDR *pNMHDR, LRESULT *pResult);
private:
	IEdmVault5Ptr m_VaultPtr;
public:
	// ��ȡָ��vaultָ��
	IEdmVault5Ptr GetVaultPtr(CString sVaultName = L"");
	afx_msg void OnNMDblclkFiletree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonVarvalue();
	afx_msg void OnBnClickedButtonBominfo();
	afx_msg void OnBnClickedButtonTest();
};
