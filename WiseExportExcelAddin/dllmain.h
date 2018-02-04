// dllmain.h: 模块类的声明。

class CWiseExportExcelAddinModule : public ATL::CAtlDllModuleT< CWiseExportExcelAddinModule >
{
public :
	DECLARE_LIBID(LIBID_WiseExportExcelAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WISEEXPORTEXCELADDIN, "{129EABE7-E7A9-45A8-8656-BA1199CC48C2}")
};

extern class CWiseExportExcelAddinModule _AtlModule;
