#pragma once


// CTestModelDlg �Ի���

class CTestModelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CTestModelDlg)

public:
	CTestModelDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CTestModelDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
};
