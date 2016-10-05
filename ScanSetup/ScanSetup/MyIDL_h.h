

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Wed Oct 05 09:15:15 2016
 */
/* Compiler settings for MyIDL.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

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


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IScanSetup_FWD_DEFINED__
#define __IScanSetup_FWD_DEFINED__
typedef interface IScanSetup IScanSetup;

#endif 	/* __IScanSetup_FWD_DEFINED__ */


#ifndef ___ScanSetup_FWD_DEFINED__
#define ___ScanSetup_FWD_DEFINED__
typedef interface _ScanSetup _ScanSetup;

#endif 	/* ___ScanSetup_FWD_DEFINED__ */


#ifndef __ScanSetup_FWD_DEFINED__
#define __ScanSetup_FWD_DEFINED__

#ifdef __cplusplus
typedef class ScanSetup ScanSetup;
#else
typedef struct ScanSetup ScanSetup;
#endif /* __cplusplus */

#endif 	/* __ScanSetup_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __ScanSetup_LIBRARY_DEFINED__
#define __ScanSetup_LIBRARY_DEFINED__

/* library ScanSetup */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_ScanSetup;

#ifndef __IScanSetup_DISPINTERFACE_DEFINED__
#define __IScanSetup_DISPINTERFACE_DEFINED__

/* dispinterface IScanSetup */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IScanSetup;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("DE99963D-3650-4F3A-9236-176BAE1FED09")
    IScanSetup : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IScanSetupVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IScanSetup * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IScanSetup * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IScanSetup * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IScanSetup * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IScanSetup * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IScanSetup * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IScanSetup * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IScanSetupVtbl;

    interface IScanSetup
    {
        CONST_VTBL struct IScanSetupVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IScanSetup_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IScanSetup_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IScanSetup_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IScanSetup_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IScanSetup_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IScanSetup_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IScanSetup_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IScanSetup_DISPINTERFACE_DEFINED__ */


#ifndef ___ScanSetup_DISPINTERFACE_DEFINED__
#define ___ScanSetup_DISPINTERFACE_DEFINED__

/* dispinterface _ScanSetup */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__ScanSetup;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("42234C68-AD7A-491A-ACCB-E6E439D561A3")
    _ScanSetup : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _ScanSetupVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _ScanSetup * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _ScanSetup * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _ScanSetup * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _ScanSetup * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _ScanSetup * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _ScanSetup * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _ScanSetup * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _ScanSetupVtbl;

    interface _ScanSetup
    {
        CONST_VTBL struct _ScanSetupVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _ScanSetup_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _ScanSetup_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _ScanSetup_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _ScanSetup_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _ScanSetup_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _ScanSetup_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _ScanSetup_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___ScanSetup_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ScanSetup;

#ifdef __cplusplus

class DECLSPEC_UUID("14D4697D-30DF-49CD-B32C-0460DFA0EAFE")
ScanSetup;
#endif
#endif /* __ScanSetup_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


