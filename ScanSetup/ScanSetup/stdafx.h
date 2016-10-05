// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h>
#include <ole2.h>
#include <olectl.h>

#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <math.h>

#include <propvarutil.h>
#include <Shobjidl.h>
#include <ShlObj.h>
#include <Shellapi.h>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

// visual styles
//#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
//#pragma comment( linker, "/manifestdependency:\"type='win32' \
//    name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
 //   processorArchitecture='*' publicKeyToken='6595b64144ccf1df' \
   // language='*'\"")

#pragma comment( lib, "comctl32.lib")
class CServer;
CServer * GetServer();
// get named interface
BOOL GetNamedInterface(
	IDispatch		*	pdisp,
	LPCTSTR				szName,
	IDispatch		**	ppdispNamed);

// create a tool tips window
HWND MyCreateToolTips(HWND hwndMain, HINSTANCE hinst);
// add a tool to the tool tips
void AddTool(
	HWND				hwndParent,
	HWND				hwndItem,
	HWND				hwndTT);
// definitions
#define				MY_CLSID					CLSID_ScanSetup
#define				MY_LIBID					LIBID_ScanSetup
#define				MY_IID						IID_IScanSetup
#define				MY_IIDSINK					IID__ScanSetup

#define				MAX_CONN_PTS				2
#define				CONN_PT_CUSTOMSINK			0
#define				CONN_PT_PROPNOTIFY			1

#define				FRIENDLY_NAME				TEXT("ScanSetup")
#define				PROGID						TEXT("Sciencetech.ScanSetup.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.ScanSetup")