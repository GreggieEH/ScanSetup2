// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "Server.h"
#include <initguid.h>
#include "MyGuids.h"

CServer * g_pServer = NULL;

BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD fdwReason,
	LPVOID lpvReserved
)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		g_pServer = new CServer(hinstDLL);
		break;
	case DLL_PROCESS_DETACH:
		Utils_DELETE_POINTER(g_pServer);
		break;
	default:
		break;
	}
	return TRUE;
}

CServer * GetServer()
{
	return g_pServer;
}

// get named interface
BOOL GetNamedInterface(
	IDispatch		*	pdisp,
	LPCTSTR				szName,
	IDispatch		**	ppdispNamed)
{
	HRESULT				hr;
	IProvideClassInfo*	pProvideClassInfo;
	ITypeInfo		*	pClassInfo = NULL;
	ITypeInfo		*	pTypeInfo = NULL;
	TYPEATTR		*	pTypeAttr;

	*ppdispNamed = NULL;
	hr = pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*)&pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = Utils_FindImplClassName(pClassInfo, szName, &pTypeInfo) ? S_OK : E_FAIL;
		pClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = pdisp->QueryInterface(pTypeAttr->guid, (LPVOID*)ppdispNamed);
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
	return SUCCEEDED(hr);
}

// create a tool tips window
HWND MyCreateToolTips(HWND hwndMain, HINSTANCE hinst)
{
	HWND				hwndTT;
	// create the tool tips window
	hwndTT = CreateWindowEx(NULL, TOOLTIPS_CLASS, NULL,
		WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX,
		CW_USEDEFAULT, CW_USEDEFAULT,
		CW_USEDEFAULT, CW_USEDEFAULT,
		hwndMain, NULL, hinst,
		NULL);
	SetWindowPos(hwndTT, HWND_TOPMOST, 0, 0, 0, 0,
		SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
	return hwndTT;
}

// add a tool to the tool tips
void AddTool(
	HWND				hwndParent,
	HWND				hwndItem,
	HWND				hwndTT)
{
	TOOLINFO			ti;					// tool info
	BOOL				fSuccess;
	RECT				rc;

	GetWindowRect(hwndItem, &rc);
	MapWindowPoints(NULL, hwndParent, (LPPOINT)&rc, 2);

	ZeroMemory((PVOID)&ti, sizeof(TOOLINFO));
	ti.cbSize = sizeof(TOOLINFO) - 4;
	ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;			// TTF_IDISHWND | TTF_SUBCLASS;
	ti.hwnd = hwndParent;				// control window receives message notifications
	ti.lpszText = LPSTR_TEXTCALLBACK;
	ti.hinst = NULL;
	ti.lParam = 0;
	ti.uId = (UINT_PTR)hwndItem;
	CopyRect(&(ti.rect), &rc);

	fSuccess = SendMessage(hwndTT, TTM_ADDTOOL, 0, (LPARAM)&ti);
	BOOL		stat = TRUE;
}
