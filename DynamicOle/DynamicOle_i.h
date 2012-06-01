

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Fri Jun 01 11:42:23 2012
 */
/* Compiler settings for .\DynamicOle.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __DynamicOle_i_h__
#define __DynamicOle_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IDynamicOleCom_FWD_DEFINED__
#define __IDynamicOleCom_FWD_DEFINED__
typedef interface IDynamicOleCom IDynamicOleCom;
#endif 	/* __IDynamicOleCom_FWD_DEFINED__ */


#ifndef __DynamicOleCom_FWD_DEFINED__
#define __DynamicOleCom_FWD_DEFINED__

#ifdef __cplusplus
typedef class DynamicOleCom DynamicOleCom;
#else
typedef struct DynamicOleCom DynamicOleCom;
#endif /* __cplusplus */

#endif 	/* __DynamicOleCom_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IDynamicOleCom_INTERFACE_DEFINED__
#define __IDynamicOleCom_INTERFACE_DEFINED__

/* interface IDynamicOleCom */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IDynamicOleCom;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0624AC93-AB85-420E-9667-D93AFF14997B")
    IDynamicOleCom : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetHostWindow( 
            /* [in] */ LONG hWnd) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE InsertGif( 
            /* [in] */ LONG img) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IDynamicOleComVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IDynamicOleCom * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IDynamicOleCom * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IDynamicOleCom * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IDynamicOleCom * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IDynamicOleCom * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IDynamicOleCom * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IDynamicOleCom * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SetHostWindow )( 
            IDynamicOleCom * This,
            /* [in] */ LONG hWnd);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *InsertGif )( 
            IDynamicOleCom * This,
            /* [in] */ LONG img);
        
        END_INTERFACE
    } IDynamicOleComVtbl;

    interface IDynamicOleCom
    {
        CONST_VTBL struct IDynamicOleComVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDynamicOleCom_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IDynamicOleCom_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IDynamicOleCom_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IDynamicOleCom_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IDynamicOleCom_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IDynamicOleCom_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IDynamicOleCom_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IDynamicOleCom_SetHostWindow(This,hWnd)	\
    ( (This)->lpVtbl -> SetHostWindow(This,hWnd) ) 

#define IDynamicOleCom_InsertGif(This,img)	\
    ( (This)->lpVtbl -> InsertGif(This,img) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IDynamicOleCom_INTERFACE_DEFINED__ */



#ifndef __DynamicOleLib_LIBRARY_DEFINED__
#define __DynamicOleLib_LIBRARY_DEFINED__

/* library DynamicOleLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_DynamicOleLib;

EXTERN_C const CLSID CLSID_DynamicOleCom;

#ifdef __cplusplus

class DECLSPEC_UUID("5F03E670-D93D-483A-A514-C1BB01A43713")
DynamicOleCom;
#endif
#endif /* __DynamicOleLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


