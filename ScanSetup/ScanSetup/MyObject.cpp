#include "StdAfx.h"
#include "MyObject.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"
#include "DlgAcquisitionSetup.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// object reference count
	m_cRefs(0),
/*
	// acquisition objects
	m_SciScan(this),
	m_SciInput(),
	m_FilterWheel(),
	m_SlitServer(),
	m_MonoHandler(),
	m_OptFile(this),
*/
	// dialog window
	m_hwndDialog(NULL),
	// scan parameters
	m_StartWavelength(0.0),
	m_EndWavelength(0.0),
	m_StepSize(1.0),
	m_DwellTime(55),
	m_TransitionDelayMultiplier(3),
	m_SignalAveraging(1),
	m_ScanAveraging(1),
	m_MultipleScans(1),
	m_AutoTimeConstantAvailable(FALSE),
	m_AutoTimeConstantOn(FALSE),
	// pulsed mode
	m_PulsedModeAvailable(FALSE),
	m_PulsedMode(FALSE)
{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
	// scan type default is sample
	StringCchCopy(m_szScanType, MAX_PATH, L"Sample");
	// working directory
	m_szWorkingDirectory[0] = '\0';
	this->m_szComment[0] = '\0';
}

CMyObject::~CMyObject(void)
{
	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
}

// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || MY_IID == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT						hr;

	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2)
	{
		hr = Utils_CreateConnectionPoint(this, MY_IIDSINK,
				&(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
		if (SUCCEEDED(hr))
		{
			hr = Utils_CreateConnectionPoint(this, IID_IPropertyNotifySink,
				&(this->m_paConnPts[CONN_PT_PROPNOTIFY]));
		}
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}

HRESULT CMyObject::GetClassInfo(
									ITypeInfo	**	ppTI)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

HRESULT CMyObject::GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	ITypeInfo	*	pClassInfo;
	BOOL			fSuccess	= FALSE;
	*ppTypeInfo	= NULL;
	hr = this->GetClassInfo(&pClassInfo);
	if (SUCCEEDED(hr))
	{
		fSuccess = Utils_FindImplClassName(pClassInfo, szInterface, ppTypeInfo);
		pClassInfo->Release();
	}
	return fSuccess ? S_OK : E_FAIL;
}

// recall the last working directory
void CMyObject::RecallWorkingDirectory()
{
	HRESULT			hr;
	CLSID			clsid;
	IDispatch	*	pdisp		= NULL;
	DISPID			dispid;
	VARIANT			varResult;
	TCHAR			szLastDirectory[MAX_PATH];
	
	hr = CLSIDFromProgID(L"Sciencetech.SciFileOpen.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdisp, L"LastWorkingDirectory", &dispid);
		hr = Utils_InvokePropertyGet(pdisp, dispid, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToString(varResult, szLastDirectory, MAX_PATH);
			if (SUCCEEDED(hr))
			{
				StringCchCopy(this->m_szWorkingDirectory, MAX_PATH, szLastDirectory);
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
}

// sink events
void CMyObject::FireSciScanPropChanged(
	LPCTSTR			propName)
{
	VARIANTARG			varg;
	InitVariantFromString(propName, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_SciScanPropChanged, &varg, 1);
	VariantClear(&varg);
}

BOOL CMyObject::FireRequestFriendlyName(
	LPCTSTR			ProgID,
	LPTSTR			FriendlyName,
	UINT			nBufferSize)
{
	VARIANTARG		avarg[2];
	BSTR			bstr = NULL;
	BOOL			fSuccess = FALSE;
	FriendlyName[0] = '\0';
	InitVariantFromString(ProgID, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BSTR;
	avarg[0].pbstrVal = &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestFriendlyName, avarg, 2);
	VariantClear(&avarg[0]);
	if (NULL != bstr)
	{
		StringCchCopy(FriendlyName, nBufferSize, bstr);
		fSuccess = TRUE;
		SysFreeString(bstr);
	}
	return fSuccess;
}

void CMyObject::FireObjectPropChanged(
	LPCTSTR			ObjectName,
	LPCTSTR			Property)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(ObjectName, &avarg[1]);
	InitVariantFromString(Property, &avarg[0]);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_ObjectPropChanged, avarg, 2);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
}

void CMyObject::FireRequestSetupDisplay(
	LPCTSTR			component)
{
	VARIANTARG			varg;
	InitVariantFromString(component, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetupDisplay, &varg, 1);
	VariantClear(&varg);
}

BOOL CMyObject::FireRequestReadSignal(
	double		*	percent)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	*percent = 0.0;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_R8;
	avarg[1].pdblVal = percent;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestReadSignal, avarg, 2);
	return VARIANT_TRUE == handled;
}

double CMyObject::FireRequestGetWavelength()
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	double			wavelength = 0.0;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_R8;
	avarg[1].pdblVal = &wavelength;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestGetWavelength, avarg, 2);
	return wavelength;
}

BOOL CMyObject::FireRequestSetWavelength(
	double			wavelength)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	InitVariantFromDouble(wavelength, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetWavelength, avarg, 2);
	return VARIANT_TRUE == handled;
}

BOOL CMyObject::FireRequestGetAutoTimeConstant(
	BOOL		*	fAvailable)
{
	VARIANTARG			avarg[3];
	VARIANT_BOOL		fvarAvailable = VARIANT_FALSE;
	VARIANT_BOOL		fSet = VARIANT_FALSE;
	VARIANT_BOOL		handled = VARIANT_FALSE;
	*fAvailable = FALSE;
	VariantInit(&avarg[2]);
	avarg[2].vt = VT_BYREF | VT_BOOL;
	avarg[2].pboolVal = &fvarAvailable;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_BOOL;
	avarg[1].pboolVal = &fSet;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestGetAutoTimeConstant, avarg, 3);
	*fAvailable = VARIANT_TRUE == fvarAvailable;
	return VARIANT_TRUE == fSet;
}

BOOL CMyObject::FireRequestSetAutoTimeConstant(
	BOOL			fSet)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	InitVariantFromBoolean(fSet, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetAutoTimeConstant, avarg, 2);
	return VARIANT_TRUE == handled;
}

BOOL CMyObject::FireRequestGetPulsedMode(
	BOOL		*	fAvailable)
{
	VARIANTARG		avarg[3];
	VARIANT_BOOL	fvarAvailable = VARIANT_FALSE;
	VARIANT_BOOL	fSet = VARIANT_FALSE;
	VARIANT_BOOL	handled = VARIANT_FALSE;
	*fAvailable = FALSE;
	VariantInit(&avarg[2]);
	avarg[2].vt = VT_BYREF | VT_BOOL;
	avarg[2].pboolVal = &fvarAvailable;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_BOOL;
	avarg[1].pboolVal = &fSet;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestGetPulsedMode, avarg, 3);
	*fAvailable = VARIANT_TRUE == fvarAvailable;
	return VARIANT_TRUE == fSet;
}

BOOL CMyObject::FireRequestSetPulsedMode(
	BOOL			fSet)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	InitVariantFromBoolean(fSet, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetPulsedMode, avarg, 2);
	return VARIANT_TRUE == handled;
}



// access to objects
//CSciScan* CMyObject::GetSciScan()
//{
//	return &this->m_SciScan;
//}
//CSciInput* CMyObject::GetSciInput()
//{
//	return &this->m_SciInput;
//}
//CFilterWheel* CMyObject::GetSciFilterWheel()
//{
//	return &this->m_FilterWheel;
//}
//CSlitServer* CMyObject::GetSlitServer()
//{
//	return &this->m_SlitServer;
//}
//CMonoHandler* CMyObject::GetMonoHandler()
//{
//	return &this->m_MonoHandler;
//}
//COptFile* CMyObject::GetOptFile()
//{
//	return &this->m_OptFile;
//}

// dialog window
HWND CMyObject::GetDialogWindow()
{
	return this->m_hwndDialog;
}

void CMyObject::SetDialogWindow(
	HWND			hwndDlg)
{
	this->m_hwndDialog = hwndDlg;
}

CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL)
{
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	switch (dispIdMember)
	{
/*
	case DISPID_SciScan:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSciScan(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT) || 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetSciScan(pDispParams);
		break;
	case DISPID_SciInput:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSciInput(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT) || 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetSciInput(pDispParams);
		break;
	case DISPID_FilterWheel:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetFilterWheel(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT) || 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetFilterWheel(pDispParams);
		break;
	case DISPID_SlitServer:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSlitServer(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT) || 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetSlitServer(pDispParams);
		break;
	case DISPID_MonoHandler:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetMonoHandler(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT) || 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetMonoHandler(pDispParams);
		break;
	case DISPID_OptFile:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetOptFile(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT) || 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetOptFile(pDispParams);
		break;
*/
	case DISPID_DoScanSetup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DoScanSetup(pDispParams, pVarResult);
		break;
	case DISPID_ScanType:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetScanType(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetScanType(pDispParams);
		break;
	case DISPID_WorkingDirectory:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetWorkingDirectory(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetWorkingDirectory(pDispParams);
		break;
	case DISPID_Comment:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetComment(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetComment(pDispParams);
		break;
	case DISPID_StartWavelength:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetStartWavelength(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetStartWavelength(pDispParams);
		break;
	case DISPID_EndWavelength:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetEndWavelength(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetEndWavelength(pDispParams);
		break;
	case DISPID_StepSize:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetStepSize(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetStepSize(pDispParams);
		break;
	case DISPID_NumberOfSteps:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetNumberOfSteps(pVarResult);
		break;
	case DISPID_DwellTime:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDwellTime(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDwellTime(pDispParams);
		break;
	case DISPID_TransitionDelayMultiplier:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetTransitionDelayMultiplier(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetTransitionDelayMultiplier(pDispParams);
		break;
	case DISPID_SignalAveraging:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSignalAveraging(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetSignalAveraging(pDispParams);
		break;
	case DISPID_ScanAveraging:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetScanAveraging(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetScanAveraging(pDispParams);
		break;
	case DISPID_MultipleScans:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetMultipleScans(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetMultipleScans(pDispParams);
		break;
//	case DISPID_AutoTimeContant:
//		if (0 != (wFlags & DISPATCH_PROPERTYGET))
//			return this->GetAutoTimeConstant(pVarResult);
//		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
//			return this->SetAutoTimeConstant(pDispParams);
//		break;
//	case DISPID_PulsedMode:
//		if (0 != (wFlags & DISPATCH_PROPERTYGET))
//			return this->GetPulsedMode(pVarResult);
//		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
//			return this->SetPulsedMode(pDispParams);
//		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

/*
HRESULT CMyObject::CImpIDispatch::GetSciScan(
	VARIANT		*	pVarResult)
{
	IDispatch	*	pdisp;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_SciScan.getMyObject(&pdisp))
	{
		InitVariantFromDispatch(pdisp, pVarResult);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetSciScan(
	DISPPARAMS	*	pDispParams)
{
	if (1 == pDispParams->cArgs && VT_DISPATCH == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_SciScan.doInit(pDispParams->rgvarg[0].pdispVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSciInput(
	VARIANT		*	pVarResult)
{
	IDispatch	*	pdisp;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_SciInput.getMyObject(&pdisp))
	{
		InitVariantFromDispatch(pdisp, pVarResult);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetSciInput(
	DISPPARAMS	*	pDispParams)
{
	if (1 == pDispParams->cArgs && VT_DISPATCH == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_SciInput.doInit(pDispParams->rgvarg[0].pdispVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetFilterWheel(
	VARIANT		*	pVarResult)
{
	IDispatch	*	pdisp;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_FilterWheel.getMyObject(&pdisp))
	{
		InitVariantFromDispatch(pdisp, pVarResult);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetFilterWheel(
	DISPPARAMS	*	pDispParams)
{
	if (1 == pDispParams->cArgs && VT_DISPATCH == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_FilterWheel.doInit(pDispParams->rgvarg[0].pdispVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSlitServer(
	VARIANT		*	pVarResult)
{
	IDispatch	*	pdisp;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_SlitServer.getMyObject(&pdisp))
	{
		InitVariantFromDispatch(pdisp, pVarResult);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetSlitServer(
	DISPPARAMS	*	pDispParams)
{
	if (1 == pDispParams->cArgs && VT_DISPATCH == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_SlitServer.doInit(pDispParams->rgvarg[0].pdispVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMonoHandler(
	VARIANT		*	pVarResult)
{
	IDispatch		*	pdisp;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_MonoHandler.getMyObject(&pdisp))
	{
		InitVariantFromDispatch(pdisp, pVarResult);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMonoHandler(
	DISPPARAMS	*	pDispParams)
{
	if (1 == pDispParams->cArgs && VT_DISPATCH == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_MonoHandler.doInit(pDispParams->rgvarg[0].pdispVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetOptFile(
	VARIANT		*	pVarResult)
{
	IDispatch		*	pdisp;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_OptFile.getMyObject(&pdisp))
	{
		InitVariantFromDispatch(pdisp, pVarResult);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT  CMyObject::CImpIDispatch::SetOptFile(
	DISPPARAMS	*	pDispParams)
{
	if (1 == pDispParams->cArgs && VT_DISPATCH == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_OptFile.doInit(pDispParams->rgvarg[0].pdispVal);
	}
	return S_OK;
}
*/

HRESULT CMyObject::CImpIDispatch::DoScanSetup(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT					hr;
	VARIANTARG				varg;
	UINT					uArgErr;
	CDlgAcquisitionSetup	dlg(this->m_pBackObj);
	BOOL					fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	// recall the last selected working directory
	this->m_pBackObj->RecallWorkingDirectory();
	fSuccess = dlg.DoOpenDialog((HWND)varg.lVal, GetServer()->GetInstance());
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetScanType(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_pBackObj->m_szScanType, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetScanType(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_pBackObj->m_szScanType, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_ScanType);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetWorkingDirectory(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_pBackObj->m_szWorkingDirectory, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetWorkingDirectory(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_pBackObj->m_szWorkingDirectory, MAX_PATH, varg.bstrVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_WorkingDirectory);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetComment(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_pBackObj->m_szComment, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetComment(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_pBackObj->m_szComment, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	this->m_pBackObj->FireObjectPropChanged(L"Root", L"Comment");
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetStartWavelength(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_pBackObj->m_StartWavelength, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetStartWavelength(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_StartWavelength = varg.dblVal;
	Utils_OnPropChanged(this->m_pBackObj, DISPID_StartWavelength);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetEndWavelength(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_pBackObj->m_EndWavelength, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetEndWavelength(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_EndWavelength = varg.dblVal;
	Utils_OnPropChanged(this->m_pBackObj, DISPID_EndWavelength);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetStepSize(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_pBackObj->m_StepSize, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetStepSize(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_StepSize = varg.dblVal;
	Utils_OnPropChanged(this->m_pBackObj, DISPID_StepSize);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetNumberOfSteps(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_StepSize != 0.0)
	{
		long		nSteps = (long)floor((this->m_pBackObj->m_EndWavelength - this->m_pBackObj->m_StartWavelength) / this->m_pBackObj->m_StepSize);
		InitVariantFromInt32(nSteps, pVarResult);
	}
	else
	{
		InitVariantFromInt32(-1, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetDwellTime(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_DwellTime, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetDwellTime(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_DwellTime = varg.lVal;
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetTransitionDelayMultiplier(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_TransitionDelayMultiplier, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetTransitionDelayMultiplier(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_TransitionDelayMultiplier = varg.lVal;
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSignalAveraging(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_SignalAveraging, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetSignalAveraging(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_SignalAveraging = varg.lVal;
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetScanAveraging(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_ScanAveraging, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetScanAveraging(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_ScanAveraging = varg.lVal;
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMultipleScans(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_MultipleScans, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMultipleScans(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_MultipleScans = varg.lVal;
	return S_OK;
}

/*
HRESULT CMyObject::CImpIDispatch::GetAutoTimeConstant(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_AutoTimeConstantAvailable)
	{
		InitVariantFromInt32(this->m_pBackObj->m_AutoTimeConstantOn ? 1 : 0, pVarResult);
	}
	else
	{
		InitVariantFromString(L"Not Available", pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAutoTimeConstant(
	DISPPARAMS	*	pDispParams)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_BSTR == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_AutoTimeConstantAvailable = FALSE;
		this->m_pBackObj->m_AutoTimeConstantOn = FALSE;
	}
	else
	{
		HRESULT			hr;
		VARIANTARG		varg;
		UINT			uArgErr;
		VariantInit(&varg);
		hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
		if (FAILED(hr)) return hr;
		this->m_pBackObj->m_AutoTimeConstantAvailable = TRUE;
		this->m_pBackObj->m_AutoTimeConstantOn = 0 != varg.lVal;
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetPulsedMode(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_PulsedModeAvailable)
	{
		InitVariantFromInt32(this->m_pBackObj->m_PulsedMode ? 1 : 0, pVarResult);
	}
	else
	{
		InitVariantFromString(L"Not Available", pVarResult);
	}
	return S_OK;

}

HRESULT CMyObject::CImpIDispatch::SetPulsedMode(
	DISPPARAMS	*	pDispParams)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_BSTR == pDispParams->rgvarg[0].vt)
	{
		this->m_pBackObj->m_PulsedModeAvailable = FALSE;
		this->m_pBackObj->m_PulsedMode = FALSE;
	}
	else
	{
		HRESULT			hr;
		VARIANTARG		varg;
		UINT			uArgErr;
		VariantInit(&varg);
		hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
		if (FAILED(hr)) return hr;
		this->m_pBackObj->m_PulsedModeAvailable = TRUE;
		this->m_pBackObj->m_PulsedMode = 0 != varg.lVal;
	}
	return S_OK;

}
*/

CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI) 
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID)
{
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID	= MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID	= GUID_NULL;
		return E_INVALIDARG;	
	}
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP)
{
	HRESULT					hr		= CONNECT_E_NOCONNECTION;
	IConnectionPoint	*	pConnPt	= NULL;
	*ppCP	= NULL;
	if (MY_IIDSINK == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	else if (IID_IPropertyNotifySink == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_PROPNOTIFY];
	if (NULL != pConnPt)
	{
		*ppCP = pConnPt;
		pConnPt->AddRef();
		hr = S_OK;
	}
	return hr;
}

