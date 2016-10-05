#pragma once

class CDispObject
{
public:
	CDispObject(void);
	virtual ~CDispObject(void);
	BOOL			doInit(
						LPCTSTR			szProgID);
	virtual BOOL	doInit(
						IDispatch	*	pdisp);
	BOOL			getMyObject(
						IDispatch	**	ppdisp);
	// properties
	HRESULT			getProperty(
						LPCTSTR			szProp,
						VARIANT		*	Value);
	HRESULT			setProperty(
						LPCTSTR			szProp,
						VARIANTARG	*	Value);
	BOOL			getBoolProperty(
						LPCTSTR			szProp);
	void			setBoolProperty(
						LPCTSTR			szProp,
						BOOL			bValue);
	long			getLongProperty(
						LPCTSTR			szProp);
	void			setLongProperty(
						LPCTSTR			szProp,
						long			lval);
	short int		getShortProperty(
						LPCTSTR			szProp);
	void			setShortProperty(
						LPCTSTR			szProp,
						short int		sval);
	double			getDoubleProperty(
						LPCTSTR			szProp);
	void			setDoubleProperty(
						LPCTSTR			szProp,
						double			dval);
	BOOL			getStringProperty(
						LPCTSTR			szProp,
						LPTSTR			szString,
						UINT			nBufferSize);
	BOOL			getStringProperty(
						LPCTSTR			szProp,
						LPTSTR		*	szString);
	void			setStringProperty(
						LPCTSTR			szProp,
						LPCTSTR			szString);
	HRESULT			InvokeMethod(
						LPCTSTR			szMethod,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult);
	HRESULT			InvokeMethod(
						DISPID			dispid,
						VARIANTARG	*	pvarg, 
						int				cArgs,
						VARIANT		*	pVarResult);
	HRESULT			DoInvoke(
						DISPID			dispid,
						WORD			wFlags,
						VARIANTARG	*	pvarg,
						int				cArgs,
						VARIANT		*	pVarResult);
	DISPID			GetDispid(
						LPCTSTR			szMethod);
protected:
	HRESULT			GetClassInfo(
						ITypeInfo	**	ppClassInfo);
	HRESULT			GetRefTypeInfo(
						LPCTSTR			szInterface,
						ITypeInfo	**	ppTypeInfo);
	HRESULT			GetNamedIterface(
						LPCTSTR			szInterface,
						IDispatch	**	ppdisp);
	// close up
	void			CloseUp();
	// custom sink
	void			ConnectCustomSink();
	virtual void	GetSinkDispids(
						ITypeInfo	*	pTypeInfo);
	virtual void	OnSinkEvent(
						DISPID			dispid,
						DISPPARAMS	*	pDispParams);
	// property notification
	void			ConnectPropertyNotifySink();
	virtual void	OnPropChanged(
						DISPID			dispid);


private:
	IDispatch	*	m_pdisp;
	// custom sink
	IID				m_iidSink;
	DWORD			m_dwCookie;
	// property notify sink
	DWORD			m_dwPropNotify;

	// custom sink implmentation
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CDispObject * pDispObject);
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
		CDispObject			*	m_pDispObject;
		ULONG					m_cRefs;
	};
	class CImpIPropNotify : public IPropertyNotifySink
	{
	public:
		CImpIPropNotify(CDispObject * pDispObject);
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
	protected:
		CDispObject			*	m_pDispObject;
		ULONG					m_cRefs;
	};
	friend CImpISink;
	friend CImpIPropNotify;
};
