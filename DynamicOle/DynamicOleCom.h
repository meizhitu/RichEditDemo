// DynamicOleCom.h : Declaration of the CDynamicOleCom

#pragma once
#include "resource.h"       // main symbols

#include "DynamicOle_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CDynamicOleCom

class ATL_NO_VTABLE CDynamicOleCom :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CDynamicOleCom, &CLSID_DynamicOleCom>,
	public IDispatchImpl<IDynamicOleCom, &IID_IDynamicOleCom, &LIBID_DynamicOleLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IOleControlImpl<CDynamicOleCom>,
	public IOleObjectImpl<CDynamicOleCom>,
	public IOleInPlaceActiveObjectImpl<CDynamicOleCom>,
	public IOleInPlaceObjectWindowlessImpl<CDynamicOleCom>,
	public IViewObjectExImpl<CDynamicOleCom>,
	public CComControl<CDynamicOleCom>
	
{
public:
	CDynamicOleCom()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_DYNAMICOLECOM)


BEGIN_COM_MAP(CDynamicOleCom)
	COM_INTERFACE_ENTRY(IDynamicOleCom)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IViewObjectEx)  
	COM_INTERFACE_ENTRY(IViewObject2)  
	COM_INTERFACE_ENTRY(IViewObject)  
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)  
	COM_INTERFACE_ENTRY(IOleInPlaceObject)  
	COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)  
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)  
	COM_INTERFACE_ENTRY(IOleControl)  
	COM_INTERFACE_ENTRY(IOleObject)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	HRESULT OnDraw(ATL_DRAWINFO& di);
	STDMETHOD (InsertGif)(void);
};

OBJECT_ENTRY_AUTO(__uuidof(DynamicOleCom), CDynamicOleCom)
