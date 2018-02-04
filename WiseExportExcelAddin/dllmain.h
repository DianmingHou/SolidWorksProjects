// dllmain.h : 模块类的声明。

class CWiseExportExcelAddinModule : public ATL::CAtlDllModuleT< CWiseExportExcelAddinModule >
{
public :
	DECLARE_LIBID(LIBID_WiseExportExcelAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WISEEXPORTEXCELADDIN, "{CC001DA5-B6A1-4927-80ED-E5F7C71E2982}")
};

extern class CWiseExportExcelAddinModule _AtlModule;
