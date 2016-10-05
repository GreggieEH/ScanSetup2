#include "stdafx.h"
#include "DispObject.h"

CDispObject::CDispObject(void) :
	m_pdisp(NULL),
	// custom sink
	m_iidSink(IID_NULL),
	m_dwCookie(0),
	// property notify sink
	m_dwPropNotify(0)
{
}

CDispObject::~CDispObject(void)
{
//	Utils_RELEASE_INTERFACE(this->m_pdisp);
	this->CloseUp();
}

// close up
void CDispObject::CloseUp()
{
	if (NULL != this->m_pdisp)
	{
		if (0 != this->m_dwCookie) Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, this->m_iidSink, FALSE, &this->m_dwCookie);
		if (0 != this->m_dwPropNotify) Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, IID_IPropertyNotifySink, FALSE, &this->m_dwPropNotify);
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
}

BOOL CDispObject::doInit(
						LPCTSTR			szProgID)
{
	HRESULT			hr;
	CLSID			clsid;
	BOOL			fSuccess	= FALSE;
	IDispatch	*	pdisp = NULL;

	hr = CLSIDFromProgID(szProgID, &clsid);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch,
			(LPVOID*) &pdisp);
	}
	if (SUCCEEDED(hr))
	{
		fSuccess = this->doInit(pdisp);
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CDispObject::doInit(
						IDispatch	*	pdisp)
{
//	Utils_RELEASE_INTERFACE(this->m_pdisp);
	this->CloseUp();
	if (NULL != pdisp)
	{
		this->m_pdisp	= pdisp;
		this->m_pdisp->AddRef();
		return TRUE;
	}
	return FALSE;
}

BOOL CDispObject::getMyObject(
						IDispatch	**	ppdisp)
{
	if (NULL != this->m_pdisp)
	{
		*ppdisp = this->m_pdisp;
		this->m_pdisp->AddRef();
		return TRUE;
	}
	else
	{
		*ppdisp	= NULL;
		return FALSE;
	}
}

// properties
HRESULT CDispObject::getProperty(
						LPCTSTR			szProp,
						VARIANT		*	Value)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, szProp, &dispid);
	return Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, Value);
}

HRESULT CDispObject::setProperty(
						LPCTSTR			szProp,
						VARIANTARG	*	Value)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, szProp, &dispid);
	return Utils_InvokePropertyPut(this->m_pdisp, dispid, Value, 1);
}

BOOL CDispObject::getBoolProperty(
						LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			bValue		= FALSE;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &bValue);
	return bValue;
}

void CDispObject::setBoolProperty(
						LPCTSTR			szProp,
						BOOL			bValue)
{
	VARIANTARG			varg;
	InitVariantFromBoolean(bValue, &varg);
	this->setProperty(szProp, &varg);
}

long CDispObject::getLongProperty(
						LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	long			lval		= 0;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult,&lval);
	return lval;
}

void CDispObject::setLongProperty(
						LPCTSTR			szProp,
						long			lval)
{
	VARIANTARG		varg;
	InitVariantFromInt32(lval, &varg);
	this->setProperty(szProp, &varg);
}

short int CDispObject::getShortProperty(
	LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	short int			sval = 0;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToInt16(varResult, &sval);
	return sval;
}

void CDispObject::setShortProperty(
	LPCTSTR			szProp,
	short int		sval)
{
	VARIANTARG		varg;
	InitVariantFromInt16(sval, &varg);
	this->setProperty(szProp, &varg);
}

double CDispObject::getDoubleProperty(
						LPCTSTR			szProp)
{
	HRESULT			hr;
	VARIANT			varResult;
	double			dval		= 0.0;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &dval);
	return dval;
}

void CDispObject::setDoubleProperty(
						LPCTSTR			szProp,
						double			dval)
{
	VARIANTARG			varg;
	InitVariantFromDouble(dval, &varg);
	this->setProperty(szProp, &varg);
}

BOOL CDispObject::getStringProperty(
						LPCTSTR			szProp,
						LPTSTR			szString,
						UINT			nBufferSize)
{
	BOOL		fSuccess	= FALSE;
	LPTSTR		szTemp		= NULL;

	szString[0]	= '\0';
	if (this->getStringProperty(szProp, &szTemp))
	{
		StringCchCopy(szString, nBufferSize, szTemp);
		fSuccess = TRUE;
		CoTaskMemFree((LPVOID) szTemp);
	}
	return fSuccess;
}

BOOL CDispObject::getStringProperty(
						LPCTSTR			szProp,
						LPTSTR		*	szString)
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fSuccess	= FALSE;
	*szString		= NULL;
	hr = this->getProperty(szProp, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringAlloc(varResult, szString);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CDispObject::setStringProperty(
						LPCTSTR			szProp,
						LPCTSTR			szString)
{
	VARIANTARG			varg;
	InitVariantFromString(szString, &varg);
	this->setProperty(szProp, &varg);
	VariantClear(&varg);
}

HRESULT	CDispObject::InvokeMethod(
						LPCTSTR			szMethod,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult)
{
	DISPID		dispid;
	Utils_GetMemid(this->m_pdisp, szMethod, &dispid);
	return Utils_InvokeMethod(this->m_pdisp, dispid, pvarg, cArgs, pVarResult);
}

DISPID CDispObject::GetDispid(
						LPCTSTR			szMethod)
{
	DISPID		dispid;
	Utils_GetMemid(this->m_pdisp, szMethod, &dispid);
	return dispid;
}

HRESULT CDispObject::InvokeMethod(
						DISPID			dispid,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult)
{
	return Utils_InvokeMethod(this->m_pdisp, dispid, pvarg, cArgs, pVarResult);
}

HRESULT CDispObject::DoInvoke(
	DISPID			dispid,
	WORD			wFlags,
	VARIANTARG	*	pvarg,
	int				cArgs,
	VARIANT		*	pVarResult)
{
	return Utils_DoInvoke(this->m_pdisp, dispid, wFlags, pvarg, cArgs, pVarResult);
}

HRESULT CDispObject::GetClassInfo(
	ITypeInfo	**	ppClassInfo)
{
	HRESULT			hr;
	IProvideClassInfo*	pProvideClassInfo	= NULL;
	*ppClassInfo = NULL;
	if (NULL == this->m_pdisp) return E_FAIL;
	hr = this->m_pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*)&pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(ppClassInfo);
		pProvideClassInfo->Release();
	}
	return hr;
}

HRESULT CDispObject::GetRefTypeInfo(
	LPCTSTR			szInterface,
	ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	ITypeInfo	*	pClassInfo	= NULL;
	BOOL			fSuccess = FALSE;
	*ppTypeInfo = NULL;
	hr = this->GetClassInfo(&pClassInfo);
	if (SUCCEEDED(hr))
	{
		fSuccess = Utils_FindImplClassName(pClassInfo, szInterface, ppTypeInfo);
		pClassInfo->Release();
	}
	return fSuccess ? S_OK : E_FAIL;
}

HRESULT CDispObject::GetNamedIterface(
	LPCTSTR			szInterface,
	IDispatch	**	ppdisp)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo	= NULL;
	TYPEATTR		*	pTypeAttr;
	*ppdisp = NULL;
	hr = this->GetRefTypeInfo(szInterface, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = this->m_pdisp->QueryInterface(pTypeAttr->guid, (LPVOID*)ppdisp);
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
	return hr;
}

// custom sink
void CDispObject::ConnectCustomSink()
{
	HRESULT			hr;
	CImpISink	*	pSink = NULL;
	IUnknown	*	pUnkSink = NULL;

	pSink = new CImpISink(this);
	hr = pSink->QueryInterface(this->m_iidSink, (LPVOID*)&pUnkSink);
	if (SUCCEEDED(hr))
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, this->m_iidSink, TRUE, &this->m_dwCookie);
		pUnkSink->Release();
	}
	else delete pSink;
}

void CDispObject::GetSinkDispids(
	ITypeInfo	*	pTypeInfo)
{

}

void CDispObject::OnSinkEvent(
	DISPID			dispid,
	DISPPARAMS	*	pDispParams)
{

}

// property notification
void CDispObject::ConnectPropertyNotifySink()
{
	HRESULT				hr;
	CImpIPropNotify*	pPropNotify = NULL;
	IUnknown		*	pUnkSink = NULL;
	pPropNotify = new CImpIPropNotify(this);
	hr = pPropNotify->QueryInterface(IID_IPropertyNotifySink, (LPVOID*)&pUnkSink);
	if (SUCCEEDED(hr))
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, IID_IPropertyNotifySink, TRUE, &this->m_dwPropNotify);
		pUnkSink->Release();
	}
	else delete pPropNotify;
}

void CDispObject::OnPropChanged(
	DISPID			dispid)
{

}

// custom sink implementation
CDispObject::CImpISink::CImpISink(CDispObject * pDispObject) :
	m_pDispObject(pDispObject),
	m_cRefs(0)
{
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo = NULL;
	TYPEATTR	*	pTypeAttr;

	if (NULL == this->m_pDispObject->m_pdisp) return;
	hr = Utils_GetSinkInterfaceID(this->m_pDispObject->m_pdisp, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		// get interface id
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pDispObject->m_iidSink = pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		// get the dispatch ids
		this->m_pDispObject->GetSinkDispids(pTypeInfo);
		pTypeInfo->Release();
	}
}

CDispObject::CImpISink::~CImpISink()
{

}

// IUnknown methods
STDMETHODIMP CDispObject::CImpISink::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || riid == this->m_pDispObject->m_iidSink)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CDispObject::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDispObject::CImpISink::Release()
{
	ULONG			cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CDispObject::CImpISink::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CDispObject::CImpISink::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CDispObject::CImpISink::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CDispObject::CImpISink::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	this->m_pDispObject->OnSinkEvent(dispIdMember, pDispParams);
	return S_OK;
}

CDispObject::CImpIPropNotify::CImpIPropNotify(CDispObject * pDispObject) :
	m_pDispObject(pDispObject),
	m_cRefs(0)
{

}

CDispObject::CImpIPropNotify::~CImpIPropNotify()
{

}

// IUnknown methods
STDMETHODIMP CDispObject::CImpIPropNotify::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyNotifySink == riid)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CDispObject::CImpIPropNotify::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDispObject::CImpIPropNotify::Release()
{
	ULONG			cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;

}

// IPropertyNotifySink methods
STDMETHODIMP CDispObject::CImpIPropNotify::OnChanged(
	DISPID			dispID)
{
	this->m_pDispObject->OnPropChanged(dispID);
	return S_OK;
}

STDMETHODIMP CDispObject::CImpIPropNotify::OnRequestEdit(
	DISPID			dispID)
{
	return S_OK;
}