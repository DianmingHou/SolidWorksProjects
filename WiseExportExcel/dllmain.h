// dllmain.h : 模块类的声明。

class CWiseExportExcelModule : public ATL::CAtlDllModuleT< CWiseExportExcelModule >
{
public :
	DECLARE_LIBID(LIBID_WiseExportExcelLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WISEEXPORTEXCEL, "{CE3957F2-06D0-4979-9EE4-6AB650972D1D}")
};

extern class CWiseExportExcelModule _AtlModule;
