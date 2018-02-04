#pragma once
#include "stdafx.h"

#define PRD_PER_PAGE_SIZE 19;
#define PRD_PAGE_COL_SIZE  10
#define PRD_START_ROW_INDEX 4

#define STD_PER_PAGE_SIZE 28
#define STD_PAGE_COL_SIZE  7
#define STD_START_ROW_INDEX 4

#define BUY_PER_PAGE_SIZE 28
#define BUY_PAGE_COL_SIZE 7
#define BUY_START_ROW_INDEX  4


class CWorkbook;
class CWorksheet;
class SldBsp{ //����������
public:
	CString path;//�ļ�·��
	CString parentFile;//����·��
	CString materialNumber;//���ϴ���
	CString number;//����
	CString name;//����
	int amount;//����
	CString material;//����
	double weight;//����
	double totalWeight;//�ܼ�
	CString remark;//��ע
	int type;
	SldBsp(){
		path = parentFile = materialNumber = number = name = material = remark = L"";
		amount = type = 0;
		weight = totalWeight = 0;
	}
};
class SldPrt:public SldBsp{
public:
	double rawWeight;//ë������
	CString route;//����·��
	SldPrt(){
		rawWeight = 0;
		route = "";
	}
};
class SldStd :public SldBsp{
};
class SldBuyPrd:public SldBsp{
};
class SldAsm:public SldBsp{
public:
	CString ztdm;//��ͼ���룬���ŵ�һλ
	CString zjdm;//������룬�ڶ�λ����
	CString bzr;//������
	CString bzsj;//����ʱ��
	CString pzr;//��׼��
	CString pzsj;//��׼ʱ��
	CString jdbj;//�׶α��
	CString sbxh;//�豸�ͺ�
	CString projName;
	CList<SldPrt*> sldPrtList;
	CList<SldStd*> sldStdList;
	CList<SldBuyPrd*> SldBuyPrdList;
	SldAsm(){
		ztdm = zjdm = bzr = bzsj = pzr = pzsj = jdbj = sbxh = projName = "";
	}
};

class WiseUtility
{
public:
	WiseUtility(void);
	~WiseUtility(void);
	/************************************************************************/
	/* ��ȡע���ֵ                                                         */
	/************************************************************************/
	static HRESULT GetRegistryValue(HKEY hKey, LPCWSTR iSubKey, LPCWSTR iValue, VARIANT *iData);
	static HRESULT GetRegistryValueStr(HKEY hKey, LPCWSTR iSubKey, LPCWSTR iValue, BSTR* iData);

	/************************************************************************/
	/* ��ȡָ���ļ������ù�ϵ                                               */
	/************************************************************************/
	static HRESULT GetReferencedFiles(IEdmVault5Ptr poVault,CString filePath,IEdmReference10Ptr ref,int iLevel,SldAsm &asmPrd);
	/************************************************************************/
	/* �������(����������׼�����⹺��)���ظ��Ĳ��뵽װ��ṹ��             */
	/************************************************************************/
	static HRESULT InsertSldAsmNoRepeat(SldAsm &asmPrd,SldBsp *basicPrd);
	/************************************************************************/
	/* ��ȡ����װ����Ϣ                                                     */
	/************************************************************************/
	static HRESULT GetAsmInfoFromFile(IEdmVault5Ptr poVault,CString filePath,SldAsm &asmPrd);
	/************************************************************************/
	/* ��������������Ϣ                                                 */
	/************************************************************************/
	static HRESULT GetSldPrdInfoFromFile(IEdmVault5Ptr poVault,CString filePath,SldBsp **basicPrd);


	static HRESULT GetAsmStructFromFile(IEdmVault5Ptr poVault,CString filePath,SldAsm &asmPrd);


	/************************************************************************/
	/* EXCEL��������workbook                                              */
	/************************************************************************/
	static HRESULT GetExcelWorkbook(CString filePath,CWorkbook &workbook);
	static HRESULT SaveWorkbookAsExcel(CString filePath,CWorkbook &workbook);
	static HRESULT SetTableCellValue(CWorksheet &worksheet,CString iStartCoord,CString iEndCoord,VARIANT &value);
	static HRESULT SetSingleCellValue(CWorksheet &worksheet,CString iCoord,VARIANT &value);

	static HRESULT SaveBuyPrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd);
	static HRESULT SaveStdPrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd);
	static HRESULT SavePrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd);
};