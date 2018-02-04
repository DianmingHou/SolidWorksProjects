#include "stdafx.h"
#include "WiseUtility.h"
#include <winreg.h>
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRanges.h"
#include "CRange.h"
#include <iostream>
using namespace std;
WiseUtility::WiseUtility(void)
{
}


WiseUtility::~WiseUtility(void)
{
}


HRESULT WiseUtility::GetRegistryValueStr(HKEY hKey, LPCWSTR iSubKey, LPCWSTR iValue, BSTR* iData)
{
	HKEY hKEY;
	LONG ret = RegOpenKeyEx( hKey,iSubKey,NULL, KEY_READ, &hKEY);
	if (ret != ERROR_SUCCESS)
	{
		return E_FAIL;  
	}
	DWORD dataType;  
	DWORD dataSize;
	char data[100] = {0};
	ret = RegQueryValueEx(hKEY, iValue, NULL, &dataType, (LPBYTE)data, &dataSize );
	if (ret != ERROR_SUCCESS)
	{
		return E_FAIL;  
	}
	*iData = ::SysAllocStringByteLen(data,dataSize);
	RegCloseKey(hKEY);
	return S_OK;
}

HRESULT WiseUtility::GetRegistryValue(HKEY hKey, LPCWSTR iSubKey, LPCWSTR iValue, VARIANT *iData)
{
	HKEY hKEY;
	LONG ret = RegOpenKeyEx( hKey,iSubKey,NULL, KEY_READ, &hKEY);
	if (ret != ERROR_SUCCESS)
	{
		return E_FAIL;  
	}
	DWORD dataType;  
	DWORD dataSize;
	char data[100] = {0};
	ret = RegQueryValueEx(hKEY, iValue, NULL, &dataType, (LPBYTE)data, &dataSize );
	if (ret != ERROR_SUCCESS)
	{
		return E_FAIL;  
	}
	VariantInit(iData);
	switch (dataType)
	{
	case REG_SZ:
		iData->vt = VT_BSTR;
		iData->bstrVal = _bstr_t(data);
		break;
	case REG_BINARY:
		break;
	case REG_DWORD:
		break;
	default:
		;
	};
	RegCloseKey(hKEY);
	return S_OK;
}

HRESULT WiseUtility::GetAsmStructFromFile(IEdmVault5Ptr poVault,CString filePath,SldAsm &asmPrd){
	IEdmFolder5Ptr folderPtr = NULL;
	IEdmFile5Ptr poFile = NULL;
	poVault->GetFileFromPath(filePath.AllocSysString(),&folderPtr,&poFile);

	IEdmFile7Ptr file7Ptr = (IEdmFile7Ptr)poFile;
	IEdmBomViewPtr iBomViewPtr = NULL;
	VARIANT vParam;
	VariantInit(&vParam);
	vParam.vt = VT_I4;
	vParam.lVal = 1;
	//vParam.bstrVal = L"BOM";
	file7Ptr->GetComputedBOM(vParam, 0, L"@", EdmBf_AsBuilt, &iBomViewPtr);
	//VariantClear(&vParam);
	if (iBomViewPtr==NULL)
	{
		return E_FAIL;
	}
	//IEdmPos5Ptr folderPosPtr = NULL;
	//poFile->GetFirstFolderPosition(&folderPosPtr);
	//VARIANT_BOOL isNull = VARIANT_FALSE;
	//while (SUCCEEDED(folderPosPtr->get_IsNull(&isNull))&&isNull==VARIANT_FALSE)
	//{
	//	//�ɹ���ȡ���Ҳ�Ϊ��
	//	IEdmFolder5Ptr folderPtr = NULL;
	//	poFile->GetNextFolder(folderPosPtr,&folderPtr);
	//	if (folderPtr==nullptr)
	//		continue;
	//	long folderId = 0;
	//	folderPtr->get_ID(&folderId);
	//	IEdmReference5Ptr poRetRoot = NULL;
	//	poFile->GetReferenceTree(folderId,0,&poRetRoot);
	//	if(poRetRoot==nullptr)
	//		continue;
	//	IEdmReference10Ptr ref10=(IEdmReference10Ptr)poRetRoot;
	//	//ref10->GetFirstChildPosition3()
	//}

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
	for (int j = 0; j < rowElements; j++) {
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
		IEdmFolder5Ptr childFolder = NULL;
		IEdmFile5Ptr childFile = NULL;
		poVault->GetFileFromPath(pathStr,&childFolder,&childFile);
		SldBsp *basicPrd = NULL;
		WiseUtility::GetSldPrdInfoFromFile(poVault,pathStr,&basicPrd);
		if(basicPrd==NULL){
			return E_FAIL;
		}
		WiseUtility::InsertSldAsmNoRepeat(asmPrd,basicPrd);
		//for (int i = 0; i < colElements - 1; i++) {
		//	EdmBomColumn pColPtr = pColData[i];
		//	VARIANT sBomRowValue;
		//	VariantInit(&sBomRowValue);//��һ�������
		//	VARIANT sBomRowCompValue ;
		//	VariantInit(&sBomRowCompValue);//��һ�������
		//	BSTR config = NULL;
		//	VARIANT_BOOL pReadOnly = VARIANT_TRUE;
		//	pRowPtr->GetVar(pColPtr.mlVariableID, pColPtr.meType, &sBomRowValue, &sBomRowCompValue, &config, &pReadOnly);
		//	VariantClear(&sBomRowValue);
		//	VariantClear(&sBomRowCompValue);
		//	int df = 99;
		//	int ds = 0;
		//}
	}
	return S_OK;
}

HRESULT  WiseUtility::GetReferencedFiles(IEdmVault5Ptr poVault,CString filePath,IEdmReference10Ptr ref,int iLevel,SldAsm &asmPrd){
	//
	bool isTop = false;
	if (ref==nullptr)
	{
		isTop = true;
		IEdmFile5Ptr poFile = nullptr;
		IEdmFolder5Ptr poFolder = nullptr;
		poVault->GetFileFromPath(filePath.AllocSysString(),&poFolder,&poFile);

		if (poFile==nullptr||poFolder==nullptr)
			return E_FAIL;
		long folderId = 0;
		poFolder->get_ID(&folderId);
		IEdmReference5Ptr poRetRoot = NULL;
		poFile->GetReferenceTree(folderId,0,&poRetRoot);
		if(poRetRoot==nullptr)
			return E_FAIL;
		IEdmReference10Ptr ref10=(IEdmReference10Ptr)poRetRoot;
		BSTR projectName = asmPrd.projName.AllocSysString();
		WiseUtility::GetReferencedFiles(poVault,filePath,ref10,1,asmPrd);
		
	} 
	else
	{
		IEdmPos5Ptr refPos = nullptr;
		BSTR projectName = asmPrd.projName.AllocSysString();
		ref->GetFirstChildPosition3(&projectName,isTop,true,EdmRef_File,L"",0,&refPos);
		VARIANT_BOOL isNull = VARIANT_FALSE;
		while (SUCCEEDED(refPos->get_IsNull(&isNull))&&isNull==VARIANT_FALSE)
		{
			IEdmReference5Ptr iRefChild = nullptr;
			ref->GetNextChild(refPos,&iRefChild);
			BSTR childPath = NULL;
			iRefChild->get_FoundPath(&childPath);

			cout<<childPath<<endl;
			long refCount = 0;
			IEdmReference10Ptr ref10Ptr = (IEdmReference10Ptr)iRefChild;
			if(ref10Ptr==nullptr){
				cout<<"ref10ptr is null"<<endl;
				continue;
			}
			ref10Ptr->get_RefCount(&refCount);
			//��ȡ�ļ��������ļ�
			SldBsp *basicPrd = NULL;
			WiseUtility::GetSldPrdInfoFromFile(poVault,childPath,&basicPrd);
			if(basicPrd==NULL){
				continue;
			}
			basicPrd->amount = refCount;
			basicPrd->totalWeight = refCount*basicPrd->weight;
			WiseUtility::InsertSldAsmNoRepeat(asmPrd,basicPrd);
			GetReferencedFiles(poVault,childPath,ref10Ptr,iLevel+1,asmPrd);
		}
	}
	return S_OK;
}

HRESULT WiseUtility::InsertSldAsmNoRepeat(SldAsm &asmPrd,SldBsp *basicPrd){
	if(basicPrd->type==0){
		//���������ж��ظ���
		int listSize = asmPrd.sldPrtList.GetSize();
		POSITION pos = asmPrd.sldPrtList.GetHeadPosition();
		bool hasFind = false;
		for(int i=0;i<listSize;i++){
			SldPrt *prt = asmPrd.sldPrtList.GetNext(pos);
			if(prt->number==basicPrd->number){
				//���ҵ�
				hasFind = true;
				prt->amount+=basicPrd->amount;
				prt->totalWeight = prt->amount*prt->weight;
			}
		}
		if (hasFind==false){
			//δ�ҵ�������
			asmPrd.sldPrtList.AddTail((SldPrt *)basicPrd);
		}
	}else if (basicPrd->type==1)
	{
		//�⹺��
		int listSize = asmPrd.SldBuyPrdList.GetSize();
		POSITION pos = asmPrd.SldBuyPrdList.GetHeadPosition();
		bool hasFind = false;
		for(int i=0;i<listSize;i++){
			SldBuyPrd *prt = asmPrd.SldBuyPrdList.GetNext(pos);
			if(prt->number==basicPrd->number){
				//���ҵ�
				hasFind = true;
				prt->amount+=basicPrd->amount;
				prt->totalWeight = prt->amount*prt->weight;
			}
		}
		if (hasFind==false){
			//δ�ҵ�������
			asmPrd.SldBuyPrdList.AddTail((SldBuyPrd *)basicPrd);
		}
	}else if(basicPrd->type==2)
	{
		//�⹺��
		int listSize = asmPrd.sldStdList.GetSize();
		POSITION pos = asmPrd.sldStdList.GetHeadPosition();
		bool hasFind = false;
		for(int i=0;i<listSize;i++){
			SldStd *prt = asmPrd.sldStdList.GetNext(pos);
			if(prt->number==basicPrd->number){
				//���ҵ�
				hasFind = true;
				prt->amount+=basicPrd->amount;
				prt->totalWeight = prt->amount*prt->weight;
			}
		}
		if (hasFind==false){
			//δ�ҵ�������
			asmPrd.sldStdList.AddTail((SldStd *)basicPrd);
		}
	}
	return S_OK;
}
HRESULT WiseUtility::GetSldPrdInfoFromFile(IEdmVault5Ptr poVault,CString filePath,SldBsp **basicPrd){
	IEdmFolder5Ptr folderPtr = NULL;
	IEdmFile5Ptr poFile = NULL;
	poVault->GetFileFromPath(filePath.AllocSysString(),&folderPtr,&poFile);
	IEdmEnumeratorVariable5Ptr enumVarPtr = NULL;
	VARIANT vParam;
	VARIANT_BOOL isSuccess = VARIANT_FALSE;
	VariantInit(&vParam);
	poFile->GetEnumeratorVariable(filePath.AllocSysString(), &enumVarPtr);
	//��ȡ����
	CString partType = L"";
	enumVarPtr->GetVar(L"Part Type", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		partType =vParam.bstrVal;
	if(partType.GetLength()==0){
		VariantInit(&vParam);
		enumVarPtr->GetVar(L"�������", L"@", &vParam, &isSuccess);
		if(vParam.vt==VT_BSTR)partType =vParam.bstrVal;
		VariantClear(&vParam);
	}
	VariantClear(&vParam);

	if (partType.Compare(L"���Ƽ�")==0){
		(*basicPrd) = new  SldPrt();
		(*basicPrd)->type = 0;
	} else if (partType.Compare(L"�⹺��")==0){
		(*basicPrd) = new  SldBuyPrd();
		(*basicPrd)->type = 1;
	} else if (partType.Compare(L"��׼��")==0){ 
		(*basicPrd) = new  SldStd();
		(*basicPrd)->type = 2;
	} else { 
		return S_FALSE;
	}
	(*basicPrd)->path = filePath;
	if(enumVarPtr==NULL)
		return E_FAIL;
	//��ȡ��������
	(*basicPrd)->amount = 1;
	(*basicPrd)->weight = 0;
	enumVarPtr->GetVar(L"����", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		(*basicPrd)->number = vParam.bstrVal;
	VariantClear(&vParam);
	VariantInit(&vParam);
	enumVarPtr->GetVar(L"����", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		(*basicPrd)->name = vParam.bstrVal;
	VariantClear(&vParam);
	VariantInit(&vParam);
	//����
	enumVarPtr->GetVar(L"����", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		(*basicPrd)->name = vParam.bstrVal;
	VariantClear(&vParam);
	VariantInit(&vParam);
	//����
	enumVarPtr->GetVar(L"Weight", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR){
		char *weightStr = _com_util::ConvertBSTRToString(vParam.bstrVal);
		if(weightStr!=NULL)sscanf(weightStr,"%lf",&((*basicPrd)->weight));
	}else if (vParam.vt==VT_R8){
		(*basicPrd)->weight = vParam.dblVal;
	}
	//����
	(*basicPrd)->totalWeight =(*basicPrd)->weight * ((*basicPrd)->amount);
	VariantClear(&vParam);
	VariantInit(&vParam);

	if ((*basicPrd)->type==1)
	{
		//����·��
		SldPrt * prt = (SldPrt *)(*basicPrd);
		enumVarPtr->GetVar(L"����·��", L"@", &vParam, &isSuccess);
		if(vParam.vt==VT_BSTR)
			prt->route = vParam.bstrVal;
		prt->rawWeight = 0;
		prt->materialNumber = L"710."+prt->number;
	}

	VariantClear(&vParam);
	return S_OK;

}
HRESULT WiseUtility::GetAsmInfoFromFile(IEdmVault5Ptr poVault,CString filePath,SldAsm &asmPrd){
	//
	IEdmFolder5Ptr folderPtr = NULL;
	IEdmFile5Ptr poFile = NULL;
	poVault->GetFileFromPath(filePath.AllocSysString(),&folderPtr,&poFile);

	IEdmEnumeratorVariable5Ptr enumVarPtr = NULL;
	poFile->GetEnumeratorVariable(filePath.AllocSysString(), &enumVarPtr);

	if(filePath!="")
	asmPrd.path = filePath;
	
	if(enumVarPtr==NULL)
		return E_FAIL;
	VARIANT vParam;
	//��ȡ��������
	asmPrd.amount = 1;
	asmPrd.weight = 0;
	VariantInit(&vParam);
	VARIANT_BOOL isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"����", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.number = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);
	//��ͼ������������
	if (asmPrd.number!="")
	{
		asmPrd.ztdm = asmPrd.number.Left(asmPrd.number.Find(L".",0));
		asmPrd.zjdm = asmPrd.number.Right(asmPrd.number.GetLength()-asmPrd.number.Find(L".",0)-1);
	}
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"����", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.name = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);

	//��ȡ�����˺ͱ�������
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"���", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.bzr = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"�������", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_DATE){
		DATE vDate = vParam.date;
		COleDateTime odt = COleDateTime(vDate);
		asmPrd.bzsj = odt.Format(L"%Y/%m/%d");
	}else if (vParam.vt==VT_BSTR)
	{
		asmPrd.bzsj =vParam.bstrVal;// _com_util::ConvertBSTRToString(vParam.bstrVal);
	}
		
	VariantClear(&vParam);
	//��ȡ��׼�˺���׼����
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"��׼", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.pzr = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"��׼����", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_DATE){
		DATE vDate = vParam.date;
		COleDateTime odt = COleDateTime(vDate);
		asmPrd.pzsj = odt.Format(L"%Y/%m/%d");
	}else if (vParam.vt==VT_BSTR)
	{
		asmPrd.pzsj =vParam.bstrVal;// _com_util::ConvertBSTRToString(vParam.bstrVal);
	}
	VariantClear(&vParam);
	//��ȡ�׶α��
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"�׶α��", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.jdbj = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);
	//��ȡ�豸�ͺ�
	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"�豸�ͺ�", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.sbxh = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);


	VariantInit(&vParam);
	isSuccess = VARIANT_FALSE;
	enumVarPtr->GetVar(L"Project Name", L"@", &vParam, &isSuccess);
	if(vParam.vt==VT_BSTR)
		asmPrd.projName = _com_util::ConvertBSTRToString(vParam.bstrVal);
	VariantClear(&vParam);

	//
	return S_OK;
}

HRESULT WiseUtility::GetExcelWorkbook(CString filePath,CWorkbook &workbook){
	CApplication ExcelApp;
	CWorkbooks books;
	//����Excel ������(����Excel)
	if(!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))
	{
		AfxMessageBox(_T("����Excel������ʧ��!"));
		return E_FAIL;
	}
	ExcelApp.put_Visible(FALSE);
	ExcelApp.put_UserControl(FALSE);
	/*�õ�����������*/
	books.AttachDispatch(ExcelApp.get_Workbooks());
	/*��һ��������*/
	LPDISPATCH lpDisp = books.Open(filePath, 
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
		vtMissing, vtMissing, vtMissing, vtMissing);
	workbook.AttachDispatch(lpDisp);
	return S_OK;
}

HRESULT WiseUtility::SaveWorkbookAsExcel(CString filePath,CWorkbook &workbook){
	
	CApplication ExcelApp = workbook.get_Application();
	VARIANT path;
	VariantInit(&path);
	path.vt = VT_BSTR;
	path.bstrVal = filePath.AllocSysString();
	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	workbook.SaveAs(path,covOptional,covOptional,covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional,covOptional);
	VariantClear(&path);
	workbook.ReleaseDispatch();
	ExcelApp.Quit();
	ExcelApp.ReleaseDispatch();
	return S_OK;
}

HRESULT WiseUtility::SetSingleCellValue(CWorksheet &worksheet,CString iCoord,VARIANT &value){
	CRange range = worksheet.get_Range(_variant_t(iCoord),vtMissing);
	range.put_Value2(value);
	range.ReleaseDispatch();
	return S_OK;

}
HRESULT WiseUtility::SetTableCellValue(CWorksheet &worksheet,CString iStartCoord,CString iEndCoord,VARIANT &value){

	CRange range = worksheet.get_Range(variant_t(iStartCoord),variant_t(iEndCoord));
	range.put_Value2(value);
	range.ReleaseDispatch();
	return S_OK;

}

HRESULT WiseUtility::SaveBuyPrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd){


	/************************************************************************/
	/*         ���޸ı䶯��                                                   */
	/************************************************************************/
	long count = asmPrd.SldBuyPrdList.GetCount();
	long perPageSize = 32;//ÿҳ�Ĵ�С
	long perColSize = 8;//ÿһ��������
	long rowStartIndex = 4;//��ϸ���ݿ�ʼ��EXCEL����
	char * pageBaseName = "�⹺��Ŀ¼";

	int  pageSize = count/perPageSize + (count%perPageSize)>0?1:0;
	CWorksheets worksheets;
	CWorksheet curWorkSheet;
	LPDISPATCH lDisp = workbook.get_Worksheets();
	worksheets.AttachDispatch(lDisp);
	LPDISPATCH lSheetDisp = worksheets.get_Item(_variant_t(pageBaseName));
	curWorkSheet.AttachDispatch(lSheetDisp);
	if (count<=0){
		//ɾ���⹺��Ŀ¼��
		curWorkSheet.Delete();
		curWorkSheet.ReleaseDispatch();
		worksheets.ReleaseDispatch();
		return S_OK;
	}else{
		if(true){
			//���һ�б�ʶ�⹺��ҳ����
			SldPrt * tempPrt = new SldPrt();
			tempPrt->name = pageBaseName;
			tempPrt->amount = pageSize;
			asmPrd.sldPrtList.AddHead(tempPrt);
		}

		CString totalPageStr;
		totalPageStr.Format(L"%4d",pageSize);
		CWorksheet aftWorkSheet;
		POSITION firstPos = asmPrd.SldBuyPrdList.GetHeadPosition();
		for (int pageIndex=0;pageIndex<pageSize;pageIndex++)
		{
			if (pageIndex<pageSize-1)//���һ�������ٸ�����
			{
				curWorkSheet.Copy(vtMissing,_variant_t(curWorkSheet));
				LPDISPATCH lpNewDisp = curWorkSheet.get_Next();
				aftWorkSheet.AttachDispatch(lpNewDisp);
			}
			CString curPageStr ;
			curPageStr.Format(L"%4d",pageIndex+1);


			/************************************************************************/
			/*         ���޸ı䶯��                                                   */
			/************************************************************************/
			//������ǰ�ģ�дֵ
			WiseUtility::SetSingleCellValue(curWorkSheet,L"A38",variant_t(asmPrd.ztdm));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"A44",variant_t(asmPrd.zjdm));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"E43",variant_t(asmPrd.bzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"H43",variant_t(asmPrd.bzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"E44",variant_t(asmPrd.pzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"H44",variant_t(asmPrd.pzsj));
			if (asmPrd.jdbj=="S")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"K42",variant_t("S"));
			else if (asmPrd.jdbj=="A")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"L42",variant_t("A"));
			else if(asmPrd.jdbj=="B")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"M42",variant_t("B"));
			else
				WiseUtility::SetSingleCellValue(curWorkSheet,L"N42",variant_t(""));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"K44",variant_t(L"��"+totalPageStr+L"ҳ"));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"P44",variant_t(L"��"+curPageStr+L"ҳ"));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"S44",variant_t(asmPrd.sbxh));
			//д��ϸ
			/**д����ϸ����**/
			CString colStrs[8] = {L"B",L"E",L"J",L"O",L"Q",L"S",L"T",L"U"};
			CString rowStr;
			for (int rowindex=rowStartIndex;rowindex<(rowStartIndex+perPageSize);rowindex++)
			{
				rowStr.Format(L"%d",rowindex);
				int hadInsSize = pageIndex*perPageSize+rowindex-rowStartIndex;//Ҫд��list�Ķ��ٸ�Ԫ��
				if(hadInsSize>=count)
					break;
				SldBuyPrd *prt = asmPrd.SldBuyPrdList.GetNext(firstPos);
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[0]+rowStr,variant_t(prt->materialNumber));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[1]+rowStr,variant_t(prt->number));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[2]+rowStr,variant_t(prt->name));
				if(prt->amount==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[3]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[3]+rowStr,variant_t(prt->amount));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[4]+rowStr,variant_t(prt->material));
				if(prt->weight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[5]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[5]+rowStr,variant_t(prt->weight));
				if(prt->totalWeight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[6]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[6]+rowStr,variant_t(prt->totalWeight));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[7]+rowStr,variant_t(prt->remark));
			}

			/************************************************************************/
			/*         ���޸ı䶯��                                                   */
			/************************************************************************/

			//�������֮��
			if (aftWorkSheet.m_lpDispatch>0)
			{
				curWorkSheet.ReleaseDispatch();
				curWorkSheet.AttachDispatch(aftWorkSheet);
				aftWorkSheet.ReleaseDispatch();
			}
		}
	}

	curWorkSheet.ReleaseDispatch();
	worksheets.ReleaseDispatch();
	return S_OK;
}
HRESULT WiseUtility::SaveStdPrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd){
	long count = asmPrd.sldStdList.GetCount();
	long perPageSize = 32;//ÿҳ�Ĵ�С
	long perColSize = 8;//ÿһ��������
	long rowStartIndex = 4;//��ϸ���ݿ�ʼ��EXCEL����
	char * pageBaseName = "��׼��Ŀ¼";

	int  pageSize = count/perPageSize + (count%perPageSize)>0?1:0;
	CWorksheets worksheets;
	CWorksheet curWorkSheet;
	LPDISPATCH lDisp = workbook.get_Worksheets();
	worksheets.AttachDispatch(lDisp);
	LPDISPATCH lSheetDisp = worksheets.get_Item(_variant_t(pageBaseName));
	curWorkSheet.AttachDispatch(lSheetDisp);
	if (count<=0){
		//ɾ���⹺��Ŀ¼��
		curWorkSheet.Delete();
		curWorkSheet.ReleaseDispatch();
		worksheets.ReleaseDispatch();
		return S_OK;
	}else{
		if(true){
			//���һ�б�ʶ�⹺��ҳ����
			SldPrt * tempPrt = new SldPrt();
			tempPrt->name = pageBaseName;
			tempPrt->amount = pageSize;
			asmPrd.sldPrtList.AddHead(tempPrt);
		}

		CString totalPageStr;
		totalPageStr.Format(L"%4d",pageSize);
		CWorksheet aftWorkSheet;
		POSITION firstPos = asmPrd.sldStdList.GetHeadPosition();
		for (int pageIndex=0;pageIndex<pageSize;pageIndex++)
		{
			if (pageIndex<pageSize-1)//���һҳ�����ٸ�����
			{
				curWorkSheet.Copy(vtMissing,_variant_t(curWorkSheet));
				LPDISPATCH lpNewDisp = curWorkSheet.get_Next();
				aftWorkSheet.AttachDispatch(lpNewDisp);
			}
			CString curPageStr ;
			curPageStr.Format(L"%4d",pageIndex+1);

			/************************************************************************/
			/*         ���޸ı䶯��                                                   */
			/************************************************************************/
			//������ǰ�ģ�дֵ
			WiseUtility::SetSingleCellValue(curWorkSheet,L"A38",variant_t(asmPrd.ztdm));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"A44",variant_t(asmPrd.zjdm));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"E43",variant_t(asmPrd.bzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"H43",variant_t(asmPrd.bzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"E44",variant_t(asmPrd.pzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"H44",variant_t(asmPrd.pzsj));
			if (asmPrd.jdbj=="S")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"K42",variant_t("S"));
			else if (asmPrd.jdbj=="A")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"L42",variant_t("A"));
			else if(asmPrd.jdbj=="B")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"M42",variant_t("B"));
			else
				WiseUtility::SetSingleCellValue(curWorkSheet,L"N42",variant_t(""));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"K44",variant_t(L"��"+totalPageStr+L"ҳ"));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"P44",variant_t(L"��"+curPageStr+L"ҳ"));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"S44",variant_t(asmPrd.sbxh));
			//д��ϸ
			/**д����ϸ����**/
			CString colStrs[8] = {L"B",L"E",L"J",L"O",L"Q",L"S",L"T",L"U"};
			CString rowStr;
			for (int rowindex=rowStartIndex;rowindex<(rowStartIndex+perPageSize);rowindex++)
			{
				rowStr.Format(L"%d",rowindex);
				int hadInsSize = pageIndex*perPageSize+rowindex-rowStartIndex;//Ҫд��list�Ķ��ٸ�Ԫ��
				if(hadInsSize>=count)
					break;
				SldStd *prt = asmPrd.sldStdList.GetNext(firstPos);
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[0]+rowStr,variant_t(prt->materialNumber));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[1]+rowStr,variant_t(prt->number));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[2]+rowStr,variant_t(prt->name));
				if(prt->amount==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[3]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[3]+rowStr,variant_t(prt->amount));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[4]+rowStr,variant_t(prt->material));
				if(prt->weight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[5]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[5]+rowStr,variant_t(prt->weight));
				if(prt->totalWeight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[6]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[6]+rowStr,variant_t(prt->totalWeight));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[7]+rowStr,variant_t(prt->remark));
			}

			/************************************************************************/
			/*                                                                      */
			/************************************************************************/

			//�������֮��
			if (aftWorkSheet.m_lpDispatch>0)
			{
				curWorkSheet.ReleaseDispatch();
				curWorkSheet.AttachDispatch(aftWorkSheet);
				aftWorkSheet.ReleaseDispatch();
			}
		}
	}

	curWorkSheet.ReleaseDispatch();
	worksheets.ReleaseDispatch();
	return S_OK;
}
HRESULT WiseUtility::SavePrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd){

	long count = asmPrd.sldPrtList.GetCount();//  [8/25/2016 houdm]
	long perPageSize = 19;//  [8/25/2016 houdm]
	long perColSize = 10;//
	long rowStartIndex = 4;
	char * pageBaseName = "������Ŀ¼";
	
	int  pageSize = count/perPageSize + (count%perPageSize)>0?1:0;
	CWorksheets worksheets;
	CWorksheet curWorkSheet;
	LPDISPATCH lDisp = workbook.get_Worksheets();
	worksheets.AttachDispatch(lDisp);
	LPDISPATCH lSheetDisp = worksheets.get_Item(_variant_t(pageBaseName));
	curWorkSheet.AttachDispatch(lSheetDisp);
	if (count<=0){
		curWorkSheet.Delete();
		curWorkSheet.ReleaseDispatch();
		worksheets.ReleaseDispatch();
		return S_OK;
	}else{
		if(true){
			SldPrt * tempPrt = new SldPrt();
			tempPrt->name = pageBaseName;
			tempPrt->amount = pageSize;
			asmPrd.sldPrtList.AddHead(tempPrt);
			count++;
		}

		CString totalPageStr;
		totalPageStr.Format(L"%4d",pageSize);
		CWorksheet aftWorkSheet;
		POSITION firstPos = asmPrd.sldPrtList.GetHeadPosition();
		for (int pageIndex=0;pageIndex<pageSize;pageIndex++)
		{
			if (pageIndex<pageSize-1)//���һ�������ٸ�����
			{
				curWorkSheet.Copy(vtMissing,_variant_t(curWorkSheet));
				LPDISPATCH lpNewDisp = curWorkSheet.get_Next();
				aftWorkSheet.AttachDispatch(lpNewDisp);
			}
			CString curPageStr ;
			curPageStr.Format(L"%4d",pageIndex+1);


			/************************************************************************/
			/*         ���޸ı䶯��                                                   */
			/************************************************************************/
			//������ǰ�ģ�д��ͷֵ
			WiseUtility::SetSingleCellValue(curWorkSheet,L"A25",variant_t(asmPrd.ztdm));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"A31",variant_t(asmPrd.zjdm));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"E30",variant_t(asmPrd.bzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"H30",variant_t(asmPrd.bzsj));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"E31",variant_t(asmPrd.pzr));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"H31",variant_t(asmPrd.pzsj));
			if (asmPrd.jdbj=="S")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"K29",variant_t("S"));
			else if (asmPrd.jdbj=="A")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"M29",variant_t("A"));
			else if(asmPrd.jdbj=="B")
				WiseUtility::SetSingleCellValue(curWorkSheet,L"O29",variant_t("B"));
			else
				WiseUtility::SetSingleCellValue(curWorkSheet,L"R29",variant_t(""));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"K31",variant_t(L"��"+totalPageStr+L"  ҳ"));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"S31",variant_t(L"��"+curPageStr+L"  ҳ"));
			WiseUtility::SetSingleCellValue(curWorkSheet,L"W31",variant_t(asmPrd.sbxh));

			/**д����ϸ����**/
			CString colStrs[10] = {L"B",L"E",L"J",L"O",L"Q",L"S",L"T",L"U",L"V",L"AC"};
			CString rowStr;
			for (int rowindex=rowStartIndex;rowindex<(rowStartIndex+perPageSize);rowindex++)
			{
				rowStr.Format(L"%d",rowindex);
				int hadInsSize = pageIndex*perPageSize+rowindex-rowStartIndex;//Ҫд��list�Ķ��ٸ�Ԫ��
				if(hadInsSize>=count)
					break;
				SldPrt *prt = asmPrd.sldPrtList.GetNext(firstPos);
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[0]+rowStr,variant_t(prt->materialNumber));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[1]+rowStr,variant_t(prt->number));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[2]+rowStr,variant_t(prt->name));
				if(prt->amount==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[3]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[3]+rowStr,variant_t(prt->amount));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[4]+rowStr,variant_t(prt->material));
				if(prt->weight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[5]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[5]+rowStr,variant_t(prt->weight));
				if(prt->totalWeight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[6]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[6]+rowStr,variant_t(prt->totalWeight));
				if(prt->rawWeight==0)
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[7]+rowStr,variant_t(""));
				else
					WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[7]+rowStr,variant_t(prt->rawWeight));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[8]+rowStr,variant_t(prt->route));
				WiseUtility::SetSingleCellValue(curWorkSheet,colStrs[9]+rowStr,variant_t(prt->remark));
			}

			/************************************************************************/
			/*                                                                      */
			/************************************************************************/



			//�������֮��
			if (aftWorkSheet.m_lpDispatch>0)
			{
				curWorkSheet.ReleaseDispatch();
				curWorkSheet.AttachDispatch(aftWorkSheet);
				aftWorkSheet.ReleaseDispatch();
			}
		}
	}

	curWorkSheet.ReleaseDispatch();
	worksheets.ReleaseDispatch();
	return S_OK;
}