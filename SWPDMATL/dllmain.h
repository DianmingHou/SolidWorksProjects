// dllmain.h : ģ�����������

class CSWPDMATLModule : public ATL::CAtlDllModuleT< CSWPDMATLModule >
{
public :
	DECLARE_LIBID(LIBID_SWPDMATLLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_SWPDMATL, "{7DA8DFC9-1A44-42FB-8B3D-338470A34E14}")
};

extern class CSWPDMATLModule _AtlModule;
