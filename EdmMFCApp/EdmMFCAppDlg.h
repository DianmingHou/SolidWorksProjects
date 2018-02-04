
// EdmMFCAppDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"


// CEdmMFCAppDlg 对话框
class CEdmMFCAppDlg : public CDialogEx
{
// 构造
public:
	CEdmMFCAppDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EDMMFCAPP_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
	// 获取指定vault指针
	IEdmVault5Ptr GetVaultPtr(CString sVaultName = L"");
	afx_msg void OnNMDblclkFiletree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonVarvalue();
	afx_msg void OnBnClickedButtonBominfo();
	afx_msg void OnBnClickedButtonTest();
};
