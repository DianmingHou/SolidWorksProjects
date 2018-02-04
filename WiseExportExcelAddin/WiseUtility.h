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
class SldBsp{ //基本类属性
public:
	CString path;//文件路径
	CString parentFile;//父级路径
	CString materialNumber;//物料代码
	CString number;//代号
	CString name;//名称
	int amount;//数量
	CString material;//材料
	double weight;//重量
	double totalWeight;//总计
	CString remark;//备注
	int type;
	SldBsp(){
		path = parentFile = materialNumber = number = name = material = remark = L"";
		amount = type = 0;
		weight = totalWeight = 0;
	}
};
class SldPrt:public SldBsp{
public:
	double rawWeight;//毛坯重量
	CString route;//工艺路线
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
	CString ztdm;//总图代码，代号第一位
	CString zjdm;//组件代码，第二位代号
	CString bzr;//编制人
	CString bzsj;//编制时间
	CString pzr;//批准人
	CString pzsj;//批准时间
	CString jdbj;//阶段标记
	CString sbxh;//设备型号
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
	/* 获取注册表值                                                         */
	/************************************************************************/
	static HRESULT GetRegistryValue(HKEY hKey, LPCWSTR iSubKey, LPCWSTR iValue, VARIANT *iData);
	static HRESULT GetRegistryValueStr(HKEY hKey, LPCWSTR iSubKey, LPCWSTR iValue, BSTR* iData);

	/************************************************************************/
	/* 获取指定文件的引用关系                                               */
	/************************************************************************/
	static HRESULT GetReferencedFiles(IEdmVault5Ptr poVault,CString filePath,IEdmReference10Ptr ref,int iLevel,SldAsm &asmPrd);
	/************************************************************************/
	/* 将零组件(基本件、标准件、外购件)非重复的插入到装配结构下             */
	/************************************************************************/
	static HRESULT InsertSldAsmNoRepeat(SldAsm &asmPrd,SldBsp *basicPrd);
	/************************************************************************/
	/* 获取顶层装配信息                                                     */
	/************************************************************************/
	static HRESULT GetAsmInfoFromFile(IEdmVault5Ptr poVault,CString filePath,SldAsm &asmPrd);
	/************************************************************************/
	/* 按类别处理零组件信息                                                 */
	/************************************************************************/
	static HRESULT GetSldPrdInfoFromFile(IEdmVault5Ptr poVault,CString filePath,SldBsp **basicPrd);


	static HRESULT GetAsmStructFromFile(IEdmVault5Ptr poVault,CString filePath,SldAsm &asmPrd);


	/************************************************************************/
	/* EXCEL操作，打开workbook                                              */
	/************************************************************************/
	static HRESULT GetExcelWorkbook(CString filePath,CWorkbook &workbook);
	static HRESULT SaveWorkbookAsExcel(CString filePath,CWorkbook &workbook);
	static HRESULT SetTableCellValue(CWorksheet &worksheet,CString iStartCoord,CString iEndCoord,VARIANT &value);
	static HRESULT SetSingleCellValue(CWorksheet &worksheet,CString iCoord,VARIANT &value);

	static HRESULT SaveBuyPrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd);
	static HRESULT SaveStdPrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd);
	static HRESULT SavePrdToWorkbook(CWorkbook &workbook,SldAsm &asmPrd);
};