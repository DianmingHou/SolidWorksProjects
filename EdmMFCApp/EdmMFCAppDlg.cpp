
// EdmMFCAppDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "EdmMFCApp.h"
#include "EdmMFCAppDlg.h"
#include "afxdialogex.h"
#include "TestModelDlg.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CEdmMFCAppDlg �Ի���



CEdmMFCAppDlg::CEdmMFCAppDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CEdmMFCAppDlg::IDD, pParent)
	, vaultStr(_T(""))
	, userName(_T(""))
	, userPwd(_T(""))
	, loginState(_T(""))
	, m_VaultPtr(NULL)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CEdmMFCAppDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_VAULT, vaultStr);
	DDX_Text(pDX, IDC_EDIT_USERNAME, userName);
	DDX_Text(pDX, IDC_EDITPASSWORD, userPwd);
	DDX_Control(pDX, IDC_FILETREE, m_FileTree);
}

BEGIN_MESSAGE_MAP(CEdmMFCAppDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_LOGIN, &CEdmMFCAppDlg::OnBnClickedButtonLogin)
	ON_NOTIFY(TVN_SELCHANGED, IDC_FILETREE, &CEdmMFCAppDlg::OnTvnSelchangedFiletree)
	ON_NOTIFY(TVN_GETINFOTIP, IDC_FILETREE, &CEdmMFCAppDlg::OnTvnGetInfoTipFiletree)
	ON_NOTIFY(TVN_ITEMEXPANDING, IDC_FILETREE, &CEdmMFCAppDlg::OnTvnItemexpandingFiletree)
	ON_NOTIFY(NM_DBLCLK, IDC_FILETREE, &CEdmMFCAppDlg::OnNMDblclkFiletree)
	ON_BN_CLICKED(IDC_BUTTON_VARVALUE, &CEdmMFCAppDlg::OnBnClickedButtonVarvalue)
	ON_BN_CLICKED(IDC_BUTTON_BOMINFO, &CEdmMFCAppDlg::OnBnClickedButtonBominfo)
	ON_BN_CLICKED(IDC_BUTTON_TEST, &CEdmMFCAppDlg::OnBnClickedButtonTest)
END_MESSAGE_MAP()


// CEdmMFCAppDlg ��Ϣ�������

BOOL CEdmMFCAppDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			//pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CEdmMFCAppDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	//if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	//{
	//	CAboutDlg dlgAbout;
	//	dlgAbout.DoModal();
	//}
	//else
	//{
		CDialogEx::OnSysCommand(nID, lParam);
	//}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CEdmMFCAppDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CEdmMFCAppDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CEdmMFCAppDlg::OnBnClickedButtonLogin()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	HRESULT rc = S_OK;
	IEdmVault5Ptr poVault = GetVaultPtr();
	if (poVault != NULL) {
		BSTR rootPathptr = NULL;
		poVault->get_RootFolderPath(&rootPathptr);
		//AfxMessageBox(rootPathptr);
		IEdmFolder5Ptr poFolder;
		rc = poVault->get_RootFolder(&poFolder);//->RootFolder;
		if (FAILED(rc) || poFolder == NULL)
		{
			AfxMessageBox(L"��ȡ��Ŀ¼ʧ�ܣ���");
		}

		HTREEITEM rootItem = m_FileTree.InsertItem(rootPathptr, TVI_ROOT);
		//m_FileTree.InsertItem(L"wisevault0", rootItem);
		//m_FileTree.Expand(rootItem, TVE_COLLAPSE);// TVE_TOGGLE);
		//m_FileTree.SetItemState(rootItem, 22, 2);
		return;
	}
	else {
		AfxMessageBox(L"��ȡ�͵�¼�ļ���ʧ�ܣ���");
		return;
	}
	rc = poVault.CreateInstance( __uuidof(EdmVault5), NULL );

	try
	{
		//Create a vault interface.
		if( FAILED(rc) )
			_com_issue_error(rc);

		//Log in on the vault.
		rc = poVault->LoginAuto( ::SysAllocString(L"wisevault0"), (long)m_hWnd );
		if(FAILED(rc)){
			loginState = "��¼״̬����¼ʧ��";
		}else{
			loginState = "��¼״̬����¼�ɹ�";
		}


		IEdmVariableMgr5Ptr varMgr = (IEdmVariableMgr5Ptr)poVault;

		//rc = poVault->Login(userName.AllocSysString(),userPwd.AllocSysString(),vaultStr.AllocSysString());
		//if(FAILED(rc)){
		//	loginState = "��¼״̬����¼ʧ��";
		//}else{
		//	loginState = "��¼״̬����¼�ɹ�";
		//}
		//
		UpdateData(true);
		//Get a pointer to the root folder.
		BSTR rootPathptr = NULL;
		poVault->get_RootFolderPath(&rootPathptr);
		AfxMessageBox(rootPathptr );
		IEdmFolder5Ptr poFolder;
		rc = poVault->get_RootFolder(&poFolder);//->RootFolder;
		if (FAILED(rc)||poFolder==NULL)
		{
			AfxMessageBox( L"��ȡ��Ŀ¼ʧ��" );
		}

		m_FileTree.InsertItem(rootPathptr, 0, 0, TVI_ROOT);
		
		
		//Get position of first file in the folder.
		IEdmPos5Ptr poPos;
		VARIANT_BOOL isNull = VARIANT_TRUE;
		CString oMessage;
		for (poFolder->GetFirstSubFolderPosition(&poPos); SUCCEEDED(poPos->get_IsNull(&isNull)) && isNull == VARIANT_FALSE;) {
			if (varMgr != NULL) {
				//
			}
			IEdmFolder5Ptr subFolderPtr = NULL;
			poFolder->GetNextSubFolder(poPos, &subFolderPtr);
			BSTR nameptr = NULL;
			subFolderPtr->get_Name(&nameptr);
			oMessage += (LPCTSTR)nameptr;
			oMessage += L" card: ";
			IEdmCard5Ptr cardPtr = NULL;
			subFolderPtr->GetCard(L".", &cardPtr);
			IEdmPos5Ptr cardContPosPtr = NULL;
			cardPtr->GetFirstControlPosition(&cardContPosPtr);
			VARIANT_BOOL isCardNull = VARIANT_TRUE;
			for (; SUCCEEDED(cardContPosPtr->get_IsNull(&isCardNull)) && isCardNull == VARIANT_FALSE;) {
				IEdmCardControl5Ptr cardContPtr = NULL;
				cardPtr->GetNextControl(cardContPosPtr, &cardContPtr);
				IEdmCardControl6Ptr cn6ptr = (IEdmCardControl6Ptr)cardContPtr;
				BSTR contName = NULL;
				cn6ptr->get_Name(&contName);
				oMessage += L" name: ";
				oMessage += contName ;
				
			}
			cardPtr->get_Name(&nameptr);
			oMessage += (LPCTSTR)nameptr;
			oMessage += L" ; ";

			IEdmEnumeratorVariable5Ptr enumVarPtr = (IEdmEnumeratorVariable5Ptr)subFolderPtr;
			VARIANT* varValue = NULL;
			enumVarPtr->GetVar(L"Project number", NULL, varValue, &isCardNull);
			
			oMessage += "\n";
		}
		AfxMessageBox(oMessage);

		poFolder->GetFirstFilePosition(&poPos);
		poPos->get_IsNull(&isNull);
		if( isNull )
			oMessage = "The root folder of your vault does not contain any files.";
		else
			oMessage = "The root folder of your vault contains these files:\n";
		
		while( isNull == VARIANT_FALSE )
		{
			IEdmFile5Ptr poFile = NULL;
			rc = poFolder->GetNextFile( poPos, &poFile );
			//poFile->GetEnumeratorVariable()
			BSTR nameptr = NULL;
			poFile->get_Name(&nameptr);
			oMessage += (LPCTSTR)nameptr;
			oMessage += "\n";
			poPos->get_IsNull(&isNull);
		}

		AfxMessageBox( oMessage );
	}
	catch( const _com_error &roError )
	{
		//We get here if one of the methods above failed.
		if( poVault == NULL )
		{
			AfxMessageBox( _T("Could not create vault interface." ));
		}
		else
		{
			BSTR bsName = NULL;
			BSTR bsDesc = NULL;
			poVault->GetErrorString( (long)roError.Error(), &bsName, &bsDesc );
			bstr_t oName( bsName, false );
			bstr_t oDesc( bsDesc, false );

			bstr_t oMsg = "Something went wrong.\n";
			oMsg += oName;
			oMsg += "\n";
			oMsg += oDesc;
			AfxMessageBox( oMsg );
		}
	}

}


void CEdmMFCAppDlg::OnBnClickedButtonReset()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
}


void CEdmMFCAppDlg::OnTvnSelchangedFiletree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;

	HTREEITEM hItem = m_FileTree.GetSelectedItem();
	CString strText = m_FileTree.GetItemText(hItem);
	SetDlgItemText(IDC_EDIT_USERNAME, strText);
}


void CEdmMFCAppDlg::OnTvnGetInfoTipFiletree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTVGETINFOTIP pGetInfoTip = reinterpret_cast<LPNMTVGETINFOTIP>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;
	NMTVGETINFOTIP* pTVTipInfo = (NMTVGETINFOTIP*)pNMHDR;   // �������pNMHDRת��ΪNMTVGETINFOTIPָ������
	HTREEITEM hRoot = m_FileTree.GetRootItem();      // ��ȡ���ĸ��ڵ�
	CString strText;     // ÿ�����ڵ����ʾ��Ϣ
	if (pTVTipInfo->hItem == hRoot) {
		// �����껮���Ľڵ��Ǹ��ڵ㣬����ʾ��ϢΪ��   
		strText = _T("");
	}
	else
	{
		// �����껮���Ľڵ㲻�Ǹ��ڵ㣬�򽫸ýڵ�ĸ���32λ���ݸ�ʽ��Ϊ�ַ���   
		strText.Format(_T("%d"), pTVTipInfo->lParam);
	}
	// ��strText�ַ���������pTVTipInfo�ṹ�������pszText��Ա�У�����������ʾ����ΪstrText����ʾ��Ϣ   
	wcscpy(pTVTipInfo->pszText, strText);
}


void CEdmMFCAppDlg::OnTvnItemexpandingFiletree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;
	if (pNMTreeView->action == TVE_COLLAPSE) {
		return;
	}
	HRESULT rc = S_OK;
	HTREEITEM hcur = pNMTreeView->itemNew.hItem;
	CString curFileName = m_FileTree.GetItemText(hcur);
	bool isFolder = true;
	if (curFileName.ReverseFind('.') > 0) {
		//�ļ�
		isFolder = false;
	}
	if (isFolder&&!m_FileTree.ItemHasChildren(hcur)){
		//û�� ��ȡ
		HTREEITEM hRoot = m_FileTree.GetRootItem();      // ��ȡ���ĸ��ڵ�
		HTREEITEM recTreeItem = hcur;
		CString curPath = L"";
		while(recTreeItem != NULL) {
			CString strText = m_FileTree.GetItemText(recTreeItem);
			if (curPath.IsEmpty())
				curPath = strText;
			else
				curPath = strText + L"\\" + curPath;
			if (recTreeItem == hRoot) {
				break;
			}
			recTreeItem = m_FileTree.GetParentItem(recTreeItem);
		}
		AfxMessageBox(curPath);
		IEdmVault5Ptr poVault = GetVaultPtr();
		IEdmFolder5Ptr poFolder;
		poVault->GetFolderFromPath(curPath.AllocSysString(), &poFolder);
		IEdmPos5Ptr poPos = NULL;
		VARIANT_BOOL isNull = VARIANT_TRUE;
		//�����ļ���
		for (poFolder->GetFirstSubFolderPosition(&poPos); SUCCEEDED(poPos->get_IsNull(&isNull)) && isNull == VARIANT_FALSE;) {
			IEdmFolder5Ptr subFolderPtr = NULL;
			rc = poFolder->GetNextSubFolder(poPos, &subFolderPtr);
			BSTR nameptr = NULL;
			subFolderPtr->get_Name(&nameptr);
			m_FileTree.InsertItem(nameptr, hcur);
		}
		//�����ļ�
		for (poFolder->GetFirstFilePosition(&poPos); SUCCEEDED(poPos->get_IsNull(&isNull)) && isNull == VARIANT_FALSE;) {
			IEdmFile5Ptr poFile = NULL;
			rc = poFolder->GetNextFile(poPos, &poFile);
			//poFile->GetEnumeratorVariable()
			BSTR nameptr = NULL;
			poFile->get_Name(&nameptr);
			m_FileTree.InsertItem(nameptr, hcur);
		}


	}
}


// ��ȡָ��vault��ָ��
//IEdmVault5Ptr CAboutDlg::GetVaultPtr(CString sVaultName)
//{
//	if (sVaultName.IsEmpty()) {
//		sVaultName = L"wisevault0";
//	}
//	if()
//	return IEdmVault5Ptr();
//}


// ��ȡָ��vaultָ��
IEdmVault5Ptr CEdmMFCAppDlg::GetVaultPtr(CString sVaultName)
{
	HRESULT rc = S_OK;
	if (sVaultName.IsEmpty()) {
		sVaultName = L"wisevault0";
	}
//	if (m_VaultPtr == nullptr) {
		rc = m_VaultPtr.CreateInstance(__uuidof(EdmVault5), NULL);
		if (FAILED(rc) || m_VaultPtr == NULL)
			return NULL;
//	}
	VARIANT_BOOL isLogin = VARIANT_FALSE;
	m_VaultPtr->get_IsLoggedIn(&isLogin);
	if (isLogin == VARIANT_FALSE) {
		rc = m_VaultPtr->LoginAuto(sVaultName.AllocSysString(), (long)m_hWnd);
		if (FAILED(rc))
			return NULL;
	}
	return m_VaultPtr;
}


void CEdmMFCAppDlg::OnNMDblclkFiletree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;
	HRESULT rc = S_OK;
	HTREEITEM hcur = m_FileTree.GetSelectedItem();// pNMTreeView->itemNew.hItem;
	CString curFileName = m_FileTree.GetItemText(hcur);
	bool isFolder = true;
	if (curFileName.ReverseFind('.') > 0) {
		//�ļ�
		isFolder = false;
	}
	if (isFolder&&!m_FileTree.ItemHasChildren(hcur)) {
		//û�� ��ȡ
		HTREEITEM hRoot = m_FileTree.GetRootItem();      // ��ȡ���ĸ��ڵ�
		HTREEITEM recTreeItem = hcur;
		CString curPath = L"";
		while (recTreeItem != NULL) {
			CString strText = m_FileTree.GetItemText(recTreeItem);
			if (curPath.IsEmpty())
				curPath = strText;
			else
				curPath = strText + L"\\" + curPath;
			if (recTreeItem == hRoot) {
				break;
			}
			recTreeItem = m_FileTree.GetParentItem(recTreeItem);
		}
		IEdmVault5Ptr poVault = GetVaultPtr();
		IEdmFolder5Ptr poFolder;
		poVault->GetFolderFromPath(curPath.AllocSysString(), &poFolder);
		IEdmPos5Ptr poPos = NULL;
		VARIANT_BOOL isNull = VARIANT_TRUE;
		//�����ļ���
		for (poFolder->GetFirstSubFolderPosition(&poPos); SUCCEEDED(poPos->get_IsNull(&isNull)) && isNull == VARIANT_FALSE;) {
			IEdmFolder5Ptr subFolderPtr = NULL;
			rc = poFolder->GetNextSubFolder(poPos, &subFolderPtr);
			BSTR nameptr = NULL;
			subFolderPtr->get_Name(&nameptr);
			m_FileTree.InsertItem(nameptr, hcur);
		}
		//�����ļ�
		for (poFolder->GetFirstFilePosition(&poPos); SUCCEEDED(poPos->get_IsNull(&isNull)) && isNull == VARIANT_FALSE;) {
			IEdmFile5Ptr poFile = NULL;
			rc = poFolder->GetNextFile(poPos, &poFile);
			//poFile->GetEnumeratorVariable()
			BSTR nameptr = NULL;
			poFile->get_Name(&nameptr);
			m_FileTree.InsertItem(nameptr, hcur);
		}


	}
}


void CEdmMFCAppDlg::OnBnClickedButtonVarvalue()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
//	if (userPwd.IsEmpty()) {
		GetDlgItemText(IDC_EDITPASSWORD, userPwd);
		if (userPwd.IsEmpty()) {
			AfxMessageBox(L"������Ϊ��");
			return;
		}
//	}

	HTREEITEM hcur = m_FileTree.GetSelectedItem();// pNMTreeView->itemNew.hItem;
	CString curFileName = m_FileTree.GetItemText(hcur);
	bool isFolder = true;
	if (curFileName.ReverseFind('.') > 0) {
		//�ļ�
		isFolder = false;
	}
	HTREEITEM hRoot = m_FileTree.GetRootItem();      // ��ȡ���ĸ��ڵ�
	HTREEITEM recTreeItem = hcur;
	CString curPath = L"";
	while (recTreeItem != NULL) {
		CString strText = m_FileTree.GetItemText(recTreeItem);
		if (curPath.IsEmpty())
			curPath = strText;
		else
			curPath = strText + L"\\" + curPath;
		if (recTreeItem == hRoot) {
			break;
		}
		recTreeItem = m_FileTree.GetParentItem(recTreeItem);
	}
	IEdmVault5Ptr poVault = GetVaultPtr();
	IEdmFolder5Ptr poFolder;
	IEdmFile5Ptr poFile;
	IEdmEnumeratorVariable5Ptr enumVarPtr = NULL;
	if (isFolder) {
		//û�� ��ȡ
		poVault->GetFolderFromPath(curPath.AllocSysString(), &poFolder);
		enumVarPtr = (IEdmEnumeratorVariable5Ptr)poFolder;
	}
	else {
		poVault->GetFileFromPath(curPath.AllocSysString(), &poFolder, &poFile);
		poFile->GetEnumeratorVariable(curPath.AllocSysString(), &enumVarPtr);
	}
	VARIANT_BOOL isSuccess = VARIANT_FALSE;
	VARIANT vParam;
	VariantInit(&vParam);//��һ�������
	enumVarPtr->GetVar(userPwd.AllocSysString(), L"@", &vParam, &isSuccess);
	AfxMessageBox(vParam.bstrVal);
	VariantClear(&vParam);

}


void CEdmMFCAppDlg::OnBnClickedButtonBominfo()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	
	HTREEITEM hcur = m_FileTree.GetSelectedItem();// pNMTreeView->itemNew.hItem;
	CString curFileName = m_FileTree.GetItemText(hcur);
	bool isFolder = true;
	if (curFileName.ReverseFind('.') <0) {
		//�ļ��з���
		AfxMessageBox(L"�ļ���û��BOM����");
		return;
	}
	HTREEITEM hRoot = m_FileTree.GetRootItem();      // ��ȡ���ĸ��ڵ�
	HTREEITEM recTreeItem = hcur;
	CString curPath = L"";
	while (recTreeItem != NULL) {
		CString strText = m_FileTree.GetItemText(recTreeItem);
		if (curPath.IsEmpty())
			curPath = strText;
		else
			curPath = strText + L"\\" + curPath;
		if (recTreeItem == hRoot) {
			break;
		}
		recTreeItem = m_FileTree.GetParentItem(recTreeItem);
	}
	IEdmVault5Ptr poVault = GetVaultPtr();
	IEdmFolder5Ptr poFolder;
	IEdmFile5Ptr poFile;
	IEdmEnumeratorVariable5Ptr enumVarPtr = NULL;
	poVault->GetFileFromPath(curPath.AllocSysString(), &poFolder, &poFile);
	IEdmFile7Ptr file7Ptr = (IEdmFile7Ptr)poFile;
	SAFEARRAY* pArray = NULL;
	SafeArrayAllocData(pArray);
	file7Ptr->GetDerivedBOMs(&pArray);
	EdmBomInfo * pData;
	long LBound;
	long UBound;
	SafeArrayAccessData(pArray, (void HUGEP **)&pData);
	SafeArrayGetLBound(pArray, 1, &LBound);
	SafeArrayGetUBound(pArray, 1, &UBound);
	int cElements = UBound - LBound + 1;
	for (int i = 0; i < cElements - 1; i++){
		EdmBomInfo Item = pData[i];
		long id = Item.mlBomID;
		IEdmObject5Ptr object5Ptr = NULL;
		poVault->GetObject(EdmObject_BOM, id, &object5Ptr);
		//IEdmBom *bom = (IEdmBom *)object5Ptr;
	}


	IEdmVault7Ptr poVault7 = (IEdmVault7Ptr)poVault;
	IEdmBomMgrPtr bomMgrPtr = NULL;
	poVault7->CreateUtility(EdmUtil_BomMgr, (IDispatch**)&bomMgrPtr);
	SAFEARRAY* pLayoutArray = NULL;
	bomMgrPtr->GetBomLayouts(&pLayoutArray);
	EdmBomLayout *pLayoutData = NULL;
	SafeArrayAccessData(pLayoutArray, (void HUGEP **)&pLayoutData);
	SafeArrayGetLBound(pLayoutArray, 1, &LBound);
	SafeArrayGetUBound(pLayoutArray, 1, &UBound);
	cElements = UBound - LBound + 1;
	//��ȡBOM��VIEW
	IEdmBomViewPtr iBomViewPtr = NULL;
	for (int i = 0; i < cElements - 1; i++) {
		EdmBomLayout sBomLayout = pLayoutData[i];
		CString curBOMLayName = sBomLayout.mbsLayoutName;
		if (curBOMLayName.Compare(L"BOM")==0) {
			VARIANT vParam;
			VariantInit(&vParam);//��һ�������
			vParam.vt = VT_BSTR;
			vParam.bstrVal = sBomLayout.mbsLayoutName;
			file7Ptr->GetComputedBOM(vParam, 0, L"@", EdmBf_AsBuilt, &iBomViewPtr);
		}
	}
	//��ȡ��
	SAFEARRAY* pCellArray = NULL;
	iBomViewPtr->GetColumns(&pCellArray);
	EdmBomColumn *pColData = NULL;
	SafeArrayAccessData(pCellArray, (void HUGEP **)&pColData);
	long colLBound;
	long colUBound;
	SafeArrayGetLBound(pCellArray, 1, &colLBound);
	SafeArrayGetUBound(pCellArray, 1, &colUBound);
	long colElements = colUBound - colLBound + 1;

	SAFEARRAY* pRowArray = NULL;
	iBomViewPtr->GetRows(&pRowArray);
	VARIANT *pRowData = NULL;
	long rowLBound;
	long rowUBound;
	SafeArrayAccessData(pRowArray, (void HUGEP **)&pRowData);
	SafeArrayGetLBound(pRowArray, 1, &rowLBound);
	SafeArrayGetUBound(pRowArray, 1, &rowUBound);
	long rowElements = rowUBound - rowLBound + 1;
	for (int j = 0; j < rowElements - 1; j++) {
		VARIANT sBomRow = pRowData[j];
		IEdmBomCellPtr pRowPtr = NULL;
		if (sBomRow.vt == VT_UNKNOWN) {
			pRowPtr = sBomRow.punkVal;
		}
		long itemId = 0;
		pRowPtr->GetItemID(&itemId);
		BSTR pathStr = NULL;
		pRowPtr->GetPathName(&pathStr);
		long treeLevel = 0;
		pRowPtr->GetTreeLevel(&treeLevel);
		for (int i = 0; i < colElements - 1; i++) {
			EdmBomColumn pColPtr = pColData[i];
			VARIANT sBomRowValue;
			VariantInit(&sBomRowValue);//��һ�������
			VARIANT sBomRowCompValue ;
			VariantInit(&sBomRowCompValue);//��һ�������
			BSTR config = NULL;
			VARIANT_BOOL pReadOnly = VARIANT_TRUE;
			pRowPtr->GetVar(pColPtr.mlVariableID, pColPtr.meType, &sBomRowValue, &sBomRowCompValue, &config, &pReadOnly);
			VariantClear(&sBomRowValue);
			VariantClear(&sBomRowCompValue);
			int df = 99;
			int ds = 0;
		}
	}
	

	//��ȡ��
	
	
	
	////��ȡ��
	//SAFEARRAY* pCellArray = NULL;
	//iBomViewPtr->GetColumns(&pCellArray);
	//EdmBomColumn *pColData = NULL;
	//SafeArrayAccessData(pCellArray, (void HUGEP **)&pColData);
	//SafeArrayGetLBound(pCellArray, 1, &LBound);
	//SafeArrayGetUBound(pCellArray, 1, &UBound);
	//cElements = UBound - LBound + 1;
	//for (int i = 0; i < cElements - 1; i++) {
	//	EdmBomColumn pColPtr = pColData[i];
	//	int j = 0;
	//	long itemId = 0;
	//	//pColPtr->GetItemID(&itemId);
	//	BSTR pathStr = NULL;
	//	//pColPtr->GetPathName(&pathStr);
	//	long treeLevel = 0;
	//	//pColPtr->GetTreeLevel(&treeLevel);
	//}

	////��ȡ��
	//SAFEARRAY* pRowArray = NULL;
	//iBomViewPtr->GetRows(&pRowArray);
	//VARIANT *pRowData = NULL;
	//SafeArrayAccessData(pRowArray, (void HUGEP **)&pRowData);
	//SafeArrayGetLBound(pRowArray, 1, &LBound);
	//SafeArrayGetUBound(pRowArray, 1, &UBound);
	//cElements = UBound - LBound + 1;
	//for (int i = 0; i < cElements - 1; i++) {
	//	VARIANT sBomRow = pRowData[i];
	//	IEdmBomCellPtr pRowPtr = NULL;
	//	if (sBomRow.vt == VT_UNKNOWN) {
	//		pRowPtr = sBomRow.punkVal;
	//	}else if (sBomRow.vt == (VT_BYREF | VT_UNKNOWN)) {
	//		pRowPtr = *sBomRow.ppunkVal;
	//	}else if (sBomRow.vt == (VT_BYREF | VT_UNKNOWN)) {
	//		pRowPtr = *sBomRow.ppdispVal;
	//	}else {
	//	}
	//	long itemId = 0;
	//	pRowPtr->GetItemID(&itemId);
	//	BSTR pathStr = NULL;
	//	pRowPtr->GetPathName(&pathStr);
	//	long treeLevel = 0;
	//	pRowPtr->GetTreeLevel(&treeLevel);
	//	//pRowPtr->GetVar()
	//}
	
}


void CEdmMFCAppDlg::OnBnClickedButtonTest()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CTestModelDlg dlg(this);
	//CFileDialog    dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY, _T("Describe Files (*.txt)|*.txt|All Files (*.*)|*.*||"), NULL);
	CString m_strFilePath = L"";
	if (dlg.DoModal())
	{
		//m_strFilePath = dlgFile.GetPathName();
		//AfxMessageBox(m_strFilePath);
	}
}
