// dllmain.h : Declaration of module class.

class CDynamicOleModule : public CAtlDllModuleT< CDynamicOleModule >
{
public :
	DECLARE_LIBID(LIBID_DynamicOleLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DYNAMICOLE, "{1BE9E9F5-464A-4842-8F4A-D0A5DB6E4C12}")
};

extern class CDynamicOleModule _AtlModule;
