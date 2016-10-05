#pragma once

class CMyObject;

class CDlgAcquisitionSetup
{
public:
	CDlgAcquisitionSetup(CMyObject * pMyObject);
	~CDlgAcquisitionSetup();
	BOOL			DoOpenDialog(
						HWND			hwndParent,
						HINSTANCE		hInst);
protected:
	BOOL			DlgProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	BOOL			OnNotify(
						LPNMHDR			pnmh);
	BOOL			OnGetDefID();
	void			OnToolTipText(
						LPNMTTDISPINFO	pTTDispInfo);


	// display values


	// form bitmap for button
	HBITMAP			FormBitmapImage(
						HWND			hwnd,
						COLORREF		colBackground,
						LPCTSTR			szText,
						COLORREF		colText);
	// display start wave
	void			DisplayStartWave();
	void			SetStartWave();
	// end wave
	void			DisplayEndWave();
	void			SetEndWave();
	// step size	
	void			DisplayStepSize();
	void			SetStepSize();
	// display number of steps
	void			DisplayNumberOfSteps();
	// dwell time
	void			DisplayDwellTime();
	void			SetDwellTime();
	// Signal Averaging
	void			DisplaySignalAveraging();
	void			SetSignalAveraging();
	// Scan Averaging
	void			DisplayScanAveraging();
	void			SetScanAveraging();
	// transition delay multiplier
	void			DisplayTransitionDelayMultiplier();
	void			SetTransitionDelayMultiplier();
	// multiple scans
	void			DisplayMultipleScans();
	void			SetMultipleScans();
	// mode
	void			DisplayCurrentMode();
	void			OnClickedMode(
						UINT		nID);
	// current wavelength
	void			DisplayCurrentWavelength();
	void			SetCurrentWavelength();
	void			OnDeltaPosCurrentWave(
						long		iDelta);
	void			OnMonochromatorSetup();
	// read the current signal
	void			SetReadingTimerOnOff(
						BOOL		fTimerOn);
	void			OnTimerReadSignal();
	// display the scan type
	void			DisplayScanType();
	// set the scan type
	void			SetScanType(
						UINT		controlID);
	// working directory
	void			DisplayWorkingDirectory();
	void			SetWorkingDirectory();
	void			OnBrowseWorkingDirectory();
	void			SetWorkingDirectory(
						LPCTSTR		szWorkingDirectory);
	BOOL			WorkingDirectorySet();		// has the working directory been defined?
	// comment
	void			DisplayComment();
	void			SetComment();
	// Auto Integration time
	void			DisplayAutoIntegrationTime();
	void			OnClickedCheckIntegrationTimeOn();		// toggle auto integration time on / off
	void			OnSetupAutoIntegrationTime();
	// sink
	void			ConnectSink();
	void			DisconnectSink();
	// sink events
	void			OnSciScanPropChanged(
						LPCTSTR		szPropName);
	void			OnObjectPropChanged(
						LPCTSTR		ObjectName,
						LPCTSTR		Property);
	// get our object
	BOOL			GetOurObject(
						IDispatch	**	ppdisp);
	double			GetDoubleProperty(
						DISPID			dispid);
	void			SetDoubleProperty(
						DISPID			dispid,
						double			dval);
	long			GetLongProperty(
						DISPID			dispid);
	void			SetLongProperty(
						DISPID			dispid,
						long			lval);
	// property change notifications
	void			ConnectPropNotify();
	void			DisconnectPropNotify();
	void			OnPropChanged(
						DISPID			dispid);
	// pulsed mode
	BOOL			GetPulsedMode();
	void			SetPulsedMode(
						BOOL			fPulsedMode);
	// Auto integration time flag
	BOOL			GetAutoIntegrationAvailable();
	BOOL			GetAutoIntegrationTime();
	void			SetAutoIntegrationTime(
						BOOL			fAutoIntegrationTime);

private:
	CMyObject	*	m_pMyObject;
	HWND			m_hwndDlg;
	BOOL			m_fAllowClose;
	// bitmap image for the start button
	HBITMAP			m_hBitmapStartButton;
	// scanSetup sink
	DWORD			m_dwCookie;
	// property change notifications
	DWORD			m_dwPropChange;

	friend LRESULT CALLBACK DlgProcAcquisitionSetup(HWND, UINT, WPARAM, LPARAM);
	// custom sink implmentation
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CDlgAcquisitionSetup * pDlgAcquisitionSetup);
		~CImpISink();
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
		void					OnSciScanPropChanged(
									DISPPARAMS	*	pDispParams);
		void					OnObjectPropChanged(
									DISPPARAMS	*	pDispParams);
	private:
		CDlgAcquisitionSetup*	m_pDlgAcquisitionSetup;
		ULONG					m_cRefs;
	};
	class CImpIPropNotify : public IPropertyNotifySink
	{
	public:
		CImpIPropNotify(CDlgAcquisitionSetup * pDlgAcquisitionSetup);
		~CImpIPropNotify();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPropertyNotifySink methods
		STDMETHODIMP			OnChanged(
									DISPID			dispID);
		STDMETHODIMP			OnRequestEdit(
									DISPID			dispID);
	private:
		CDlgAcquisitionSetup*	m_pDlgAcquisitionSetup;
		ULONG					m_cRefs;
	};
	friend CImpISink;
	friend CImpIPropNotify;
};

