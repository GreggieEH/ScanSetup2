#pragma once
//#include "SciScan.h"
//#include "SciInput.h"
//#include "FilterWheel.h"
//#include "SlitServer.h"
//#include "MonoHandler.h"
//#include "OptFile.h"

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
	// sink events
	void						FireSciScanPropChanged(
									LPCTSTR			propName);
	BOOL						FireRequestFriendlyName(
									LPCTSTR			ProgID,
									LPTSTR			FriendlyName,
									UINT			nBufferSize);
	void						FireObjectPropChanged(
									LPCTSTR			ObjectName,
									LPCTSTR			Property);
	void						FireRequestSetupDisplay(
									LPCTSTR			component);
	BOOL						FireRequestReadSignal(
									double		*	percent);
	double						FireRequestGetWavelength();
	BOOL						FireRequestSetWavelength(
									double			wavelength);
	BOOL						FireRequestGetAutoTimeConstant(
									BOOL		*	fAvailable);
	BOOL						FireRequestSetAutoTimeConstant(
									BOOL			fSet);
	BOOL						FireRequestGetPulsedMode(
									BOOL		*	fAvailable);
	BOOL						FireRequestSetPulsedMode(
									BOOL			fSet);


	// access to objects
//	CSciScan				*	GetSciScan();
//	CSciInput				*	GetSciInput();
//	CFilterWheel			*	GetSciFilterWheel();
//	CSlitServer				*	GetSlitServer();
//	CMonoHandler			*	GetMonoHandler();
//	COptFile				*	GetOptFile();
	// dialog window
	HWND						GetDialogWindow();
	void						SetDialogWindow(
									HWND			hwndDlg);
protected:
	HRESULT						GetClassInfo(
									ITypeInfo	**	ppTI);
	HRESULT						GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo);
	// recall the last working directory
	void						RecallWorkingDirectory();
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIDispatch();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
/*
		HRESULT					GetSciScan(
									VARIANT		*	pVarResult);
		HRESULT					SetSciScan(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetSciInput(
									VARIANT		*	pVarResult);
		HRESULT					SetSciInput(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetFilterWheel(
									VARIANT		*	pVarResult);
		HRESULT					SetFilterWheel(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetSlitServer(
									VARIANT		*	pVarResult);
		HRESULT					SetSlitServer(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMonoHandler(
									VARIANT		*	pVarResult);
		HRESULT					SetMonoHandler(
									DISPPARAMS	*	pDispParams);

		HRESULT					GetOptFile(
									VARIANT		*	pVarResult);
		HRESULT					SetOptFile(
									DISPPARAMS	*	pDispParams);
*/
		HRESULT					DoScanSetup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetScanType(
									VARIANT		*	pVarResult);
		HRESULT					SetScanType(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetWorkingDirectory(
									VARIANT		*	pVarResult);
		HRESULT					SetWorkingDirectory(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetComment(
									VARIANT		*	pVarResult);
		HRESULT					SetComment(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetStartWavelength(
									VARIANT		*	pVarResult);
		HRESULT					SetStartWavelength(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetEndWavelength(
									VARIANT		*	pVarResult);
		HRESULT					SetEndWavelength(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetStepSize(
									VARIANT		*	pVarResult);
		HRESULT					SetStepSize(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetNumberOfSteps(
									VARIANT		*	pVarResult);
		HRESULT					GetDwellTime(
									VARIANT		*	pVarResult);
		HRESULT					SetDwellTime(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetTransitionDelayMultiplier(
									VARIANT		*	pVarResult);
		HRESULT					SetTransitionDelayMultiplier(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetSignalAveraging(
									VARIANT		*	pVarResult);
		HRESULT					SetSignalAveraging(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetScanAveraging(
									VARIANT		*	pVarResult);
		HRESULT					SetScanAveraging(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMultipleScans(
									VARIANT		*	pVarResult);
		HRESULT					SetMultipleScans(
									DISPPARAMS	*	pDispParams);
//		HRESULT					GetAutoTimeConstant(
//									VARIANT		*	pVarResult);
//		HRESULT					SetAutoTimeConstant(
//									DISPPARAMS	*	pDispParams);
//		HRESULT					GetPulsedMode(
//									VARIANT		*	pVarResult);
//		HRESULT					SetPulsedMode(
//									DISPPARAMS	*	pDispParams);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	// data members
	CImpIDispatch			*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2	*	m_pImpIProvideClassInfo2;
	// outer unknown for aggregation
	IUnknown				*	m_pUnkOuter;
	// object reference count
	ULONG						m_cRefs;
	// connection points array
	IConnectionPoint		*	m_paConnPts[MAX_CONN_PTS];
/*
	// acquisition objects
	CSciScan					m_SciScan;
	CSciInput					m_SciInput;
	CFilterWheel				m_FilterWheel;
	CSlitServer					m_SlitServer;
	CMonoHandler				m_MonoHandler;
	COptFile					m_OptFile;
*/
	// scan type
	TCHAR						m_szScanType[MAX_PATH];
	// working directory
	TCHAR						m_szWorkingDirectory[MAX_PATH];
	// dialog window
	HWND						m_hwndDialog;
	// scan parameters
	TCHAR						m_szComment[MAX_PATH];
	double						m_StartWavelength;
	double						m_EndWavelength;
	double						m_StepSize;
	long						m_DwellTime;
	long						m_TransitionDelayMultiplier;
	long						m_SignalAveraging;
	long						m_ScanAveraging;
	long						m_MultipleScans;
//	[id(DISPID_AutoTimeContant), helpstring("Flag auto time constant, possible values = \"Not Available\", 0 (off), 1(on)")]
//	VARIANT			AutoTimeConstant;
	BOOL						m_AutoTimeConstantAvailable;
	BOOL						m_AutoTimeConstantOn;
	// pulsed mode
	BOOL						m_PulsedModeAvailable;
	BOOL						m_PulsedMode;

};
