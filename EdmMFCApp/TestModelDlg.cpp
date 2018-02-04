// TestModelDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "EdmMFCApp.h"
#include "TestModelDlg.h"
#include "afxdialogex.h"


// CTestModelDlg 对话框

IMPLEMENT_DYNAMIC(CTestModelDlg, CDialogEx)

CTestModelDlg::CTestModelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG1, pParent)
{

}

CTestModelDlg::~CTestModelDlg()
{
}

void CTestModelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CTestModelDlg, CDialogEx)
END_MESSAGE_MAP()


// CTestModelDlg 消息处理程序
