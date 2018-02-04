#pragma once


// CTestModelDlg 对话框

class CTestModelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CTestModelDlg)

public:
	CTestModelDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CTestModelDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
};
