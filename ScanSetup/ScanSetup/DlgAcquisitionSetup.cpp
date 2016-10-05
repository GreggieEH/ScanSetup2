#include "stdafx.h"
#include "DlgAcquisitionSetup.h"
#include "MyObject.h"
//#include "FilterWheel.h"
//#include "MonoHandler.h"
//#include "OptFile.h"
//#include "SciInput.h"
//#include "SciScan.h"
//#include "SlitServer.h"
#include "MyGuids.h"
#include "dispids.h"
#include "Server.h"

// reading timer
#define				TMR_READING			0x0100

CDlgAcquisitionSetup::CDlgAcquisitionSetup(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_hwndDlg(NULL),
	m_fAllowClose(TRUE),
	// bitmap image for the start button
	m_hBitmapStartButton(NULL),
	// scanSetup sink
	m_dwCookie(0),
	// property change notifications
	m_dwPropChange(0)
{
}

CDlgAcquisitionSetup::~CDlgAcquisitionSetup()
{
	if (NULL != this->m_hBitmapStartButton)
	{
		DeleteObject((HGDIOBJ) this->m_hBitmapStartButton);
		this->m_hBitmapStartButton = NULL;
	}
}

BOOL CDlgAcquisitionSetup::DoOpenDialog(
	HWND			hwndParent,
	HINSTANCE		hInst)
{
	// open the dialog
	BOOL		fSuccess = IDOK == DialogBoxParam(
					hInst, MAKEINTRESOURCE(IDD_DIALOGAcquisitionSetup), hwndParent, 
					(DLGPROC)DlgProcAcquisitionSetup, (LPARAM) this);
	// disconnect the sinks
	this->DisconnectSink();		
	this->DisconnectPropNotify();
	return fSuccess;
}

LRESULT CALLBACK DlgProcAcquisitionSetup(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgAcquisitionSetup*	pDlg = NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgAcquisitionSetup*)lParam;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		pDlg->m_hwndDlg = hwndDlg;
	}
	else
	{
		pDlg = (CDlgAcquisitionSetup*)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
		return pDlg->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

BOOL CDlgAcquisitionSetup::DlgProc(
	UINT			uMsg,
	WPARAM			wParam,
	LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
	case WM_NOTIFY:
		return this->OnNotify((LPNMHDR)lParam);
	case DM_GETDEFID:
		return this->OnGetDefID();
	case WM_TIMER:
		if (TMR_READING == wParam)
		{
			this->OnTimerReadSignal();
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgAcquisitionSetup::OnInitDialog()
{
	// store the dialog window
	this->m_pMyObject->SetDialogWindow(this->m_hwndDlg);
	// set the start button bitmap
	this->m_hBitmapStartButton = this->FormBitmapImage(
		GetDlgItem(this->m_hwndDlg, IDOK), RGB(0, 255, 0),
		L"Start Scan", RGB(0, 0, 0));
	SendMessage(GetDlgItem(this->m_hwndDlg, IDOK),
		BM_SETIMAGE, IMAGE_BITMAP, (LPARAM) this->m_hBitmapStartButton);

	// display the current settings
	// SciScan
	this->DisplayStartWave();
	this->DisplayEndWave();
	this->DisplayStepSize();
	this->DisplayNumberOfSteps();
	this->DisplayDwellTime();
	this->DisplayTransitionDelayMultiplier();
	this->DisplaySignalAveraging();
	this->DisplayScanAveraging();
	this->DisplayMultipleScans();
	this->DisplayCurrentWavelength();
	this->DisplayAutoIntegrationTime();
	// display the current acquisition mode
	this->DisplayCurrentMode();
	// scan type
	this->DisplayScanType();
	// working directory
	this->DisplayWorkingDirectory();
	// comment
	this->DisplayComment();
	// connect sinks
	this->ConnectSink();
	this->ConnectPropNotify();
	// tool tips
	HWND		hwndTT = MyCreateToolTips(this->m_hwndDlg, GetServer()->GetInstance());
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITWORKINGDIRECTORY), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_BROWSEWORKINGDIRECTORY), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITCOMMENT), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITSTARTWAVELENGTH), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITENDWAVELENGTH), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITSTEPSIZE), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_NUMBEROFSTEPS), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITDWELLTIME), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITTIMEDELAYMULTIPLIER), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITSIGNALAVERAGING), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITSCANAVERAGING), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITMULTIPLESCANS), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_CHKSAMPLE), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_CHKINTENSITYCALIBRATION), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_CHKBACKGROUND), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDOK), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_CHKPULSEDMODE), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_CHKCONTINUOUS), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_SETUPACQMODE), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_DACSETUP), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_SLITSSETUP), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_FILTERWHEELSETUP), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_DATASETUP), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_EDITWAVELENGTH), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_UPDWAVELENGTH), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_SETWAVELENGTH), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_MONOCHROMATORSETUP), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), hwndTT);
	AddTool(this->m_hwndDlg, GetDlgItem(this->m_hwndDlg, IDC_SETUPAUTOINTEGRATIONTIME), hwndTT);
	SendMessage(hwndTT, TTM_ACTIVATE, TRUE, 0);
	// center this window
	Utils_CenterWindow(this->m_hwndDlg);
	// start the read signal timer
//	this->SetReadingTimerOnOff(TRUE);


	return TRUE;
}

BOOL CDlgAcquisitionSetup::OnCommand(
	WORD			wmID,
	WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
		if (!this->m_fAllowClose)
		{
			this->m_fAllowClose = TRUE;
			return TRUE;
		}
		// check if working directory set
		if (!this->WorkingDirectorySet())
		{
			MessageBox(this->m_hwndDlg, L"Working Directory Not Set", L"ScanSetup", MB_OK);
			return TRUE;
		}
		// else closeup
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		// clear the dialog window
		this->m_pMyObject->SetDialogWindow(NULL);
		return TRUE;
	case IDC_CHKPULSEDMODE:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedMode(IDC_CHKPULSEDMODE);
			return TRUE;
		}
		break;
	case IDC_CHKCONTINUOUS:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedMode(IDC_CHKCONTINUOUS);
			return TRUE;
		}
		break;
	case IDC_SETUPACQMODE:
		EnableWindow(this->m_hwndDlg, FALSE);
//		this->m_pMyObject->GetSciInput()->DisplaySetup();
		this->m_pMyObject->FireRequestSetupDisplay(L"Input Display");
		EnableWindow(this->m_hwndDlg, TRUE);
		this->DisplayAutoIntegrationTime();
		return TRUE;
	case IDC_DACSETUP:
		EnableWindow(this->m_hwndDlg, FALSE);
//		this->m_pMyObject->GetSciInput()->DisplayDAC();
		this->m_pMyObject->FireRequestSetupDisplay(L"DAC Display");
		EnableWindow(this->m_hwndDlg, TRUE);
		this->DisplayAutoIntegrationTime();
		return TRUE;
	case IDC_SLITSSETUP:
		EnableWindow(this->m_hwndDlg, FALSE);
//		this->m_pMyObject->GetSlitServer()->ViewSetup();
		this->m_pMyObject->FireRequestSetupDisplay(L"Slit");
		EnableWindow(this->m_hwndDlg, TRUE);
		return TRUE;
	case IDC_FILTERWHEELSETUP:
		EnableWindow(this->m_hwndDlg, FALSE);
//		this->m_pMyObject->GetSciFilterWheel()->SetupFilter();
		this->m_pMyObject->FireRequestSetupDisplay(L"FilterWheel");
		EnableWindow(this->m_hwndDlg, TRUE);
		return TRUE;
	case IDC_EDITSTARTWAVELENGTH:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetStartWave();
			return TRUE;
		}
		break;
	case IDC_EDITENDWAVELENGTH:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetEndWave();
			return TRUE;
		}
		break;
	case IDC_EDITSTEPSIZE:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetStepSize();
			return TRUE;
		}
		break;
	case IDC_EDITDWELLTIME:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetDwellTime();
			return TRUE;
		}
		break;
	case IDC_EDITTIMEDELAYMULTIPLIER:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetTransitionDelayMultiplier();
			return TRUE;
		}
		break;
	case IDC_EDITSIGNALAVERAGING:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetSignalAveraging();
			return TRUE;
		}
		break;
	case IDC_EDITSCANAVERAGING:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetScanAveraging();
			return TRUE;
		}
		break;
	case IDC_SETWAVELENGTH:
		this->SetCurrentWavelength();
		return TRUE;
	case IDC_DATASETUP:
		EnableWindow(this->m_hwndDlg, FALSE);
		//		this->m_pMyObject->GetSciInput()->DisplayDAC();
		this->m_pMyObject->FireRequestSetupDisplay(L"Data Setup");
		EnableWindow(this->m_hwndDlg, TRUE);
		/*
		{
			// set AD info string
			LPTSTR			szInfoString = NULL;
			if (this->m_pMyObject->GetSciInput()->GetInfoString(&szInfoString))
			{
				this->m_pMyObject->GetOptFile()->SetADInfoString(szInfoString);
				CoTaskMemFree((LPVOID)szInfoString);
			}
			this->m_pMyObject->GetOptFile()->Setup(this->m_hwndDlg);
		}
		*/
		return TRUE;
	// scan type
	case IDC_CHKSAMPLE:
	case IDC_CHKINTENSITYCALIBRATION:
	case IDC_CHKBACKGROUND:
		if (BN_CLICKED == wmEvent)
		{
			this->SetScanType(wmID);
			return TRUE;
		}
		break;
	case IDC_MONOCHROMATORSETUP:
		EnableWindow(this->m_hwndDlg, FALSE);
//		this->m_pMyObject->GetMonoHandler()->DoSetup();
		this->m_pMyObject->FireRequestSetupDisplay(L"Monochromator");
		EnableWindow(this->m_hwndDlg, TRUE);
		return TRUE;
	case IDC_EDITCOMMENT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetComment();
			return TRUE;
		}
		break;
	case IDC_EDITWORKINGDIRECTORY:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetWorkingDirectory();
			return TRUE;
		}
		break;
	case IDC_BROWSEWORKINGDIRECTORY:
		EnableWindow(this->m_hwndDlg, FALSE);
		this->OnBrowseWorkingDirectory();
		EnableWindow(this->m_hwndDlg, TRUE);
		return TRUE;
	case IDC_CHECKINTEGRATIONTIMEON:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedCheckIntegrationTimeOn();
			return TRUE;
		}
		break;
	case IDC_SETUPAUTOINTEGRATIONTIME:
		EnableWindow(this->m_hwndDlg, FALSE);
		this->OnSetupAutoIntegrationTime();
//		this->m_pMyObject->GetSciInput()->DisplayDAC();
		EnableWindow(this->m_hwndDlg, TRUE);
		this->DisplayAutoIntegrationTime();
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgAcquisitionSetup::OnNotify(
	LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code && IDC_UPDWAVELENGTH == pnmh->idFrom)
	{
		LPNMUPDOWN pnmud = (LPNMUPDOWN)pnmh;
		this->OnDeltaPosCurrentWave(pnmud->iDelta);
		return TRUE;
	}
	else if (TTN_GETDISPINFO == pnmh->code)
	{
		this->OnToolTipText((LPNMTTDISPINFO)pnmh);
		return TRUE;
	}

	return FALSE;
}

void CDlgAcquisitionSetup::OnToolTipText(
	LPNMTTDISPINFO	pTTDispInfo)
{
	UINT		nID;
	HINSTANCE	hInst = GetServer()->GetInstance();
	if (0 != (TTF_IDISHWND & pTTDispInfo->uFlags))
		nID = GetDlgCtrlID((HWND)pTTDispInfo->hdr.idFrom);
	else
		nID = pTTDispInfo->hdr.idFrom;
	switch (nID)
	{
	case IDC_EDITWORKINGDIRECTORY:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITWORKINGDIRECTORY;
		break;
	case IDC_BROWSEWORKINGDIRECTORY:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_BROWSEWORKINGDIRECTORY;
		break;
	case IDC_EDITCOMMENT:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITCOMMENT;
		break;
	case IDC_EDITSTARTWAVELENGTH:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITSTARTWAVELENGTH;
		break;
	case IDC_EDITENDWAVELENGTH:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITENDWAVELENGTH;
		break;
	case IDC_EDITSTEPSIZE:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITSTEPSIZE;
		break;
	case IDC_NUMBEROFSTEPS:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_NUMBEROFSTEPS;
		break;
	case IDC_EDITDWELLTIME:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITDWELLTIME;
		break;
	case IDC_EDITTIMEDELAYMULTIPLIER:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITTIMEDELAYMULTIPLIER;
		break;
	case IDC_EDITSIGNALAVERAGING:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITSIGNALAVERAGING;
		break;
	case IDC_EDITSCANAVERAGING:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITSCANAVERAGING;
		break;
	case IDC_EDITMULTIPLESCANS:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITMULTIPLESCANS;
		break;
	case IDC_CHKSAMPLE:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_CHKSAMPLE;
		break;
	case IDC_CHKINTENSITYCALIBRATION:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_CHKINTENSITYCALIBRATION;
		break;
	case IDC_CHKBACKGROUND:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_CHKBACKGROUND;
		break;
	case IDOK:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_IDOK;
		break;
	case IDC_CHKPULSEDMODE:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_CHKPULSEDMODE;
		break;
	case IDC_CHKCONTINUOUS:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_CHKCONTINUOUS;
		break;
	case IDC_SETUPACQMODE:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_SETUPACQMODE;
		break;
	case IDC_DACSETUP:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_DACSETUP;
		break;
	case IDC_SLITSSETUP:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_SLITSSETUP;
		break;
	case IDC_FILTERWHEELSETUP:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_FILTERWHEELSETUP;
		break;
	case IDC_DATASETUP:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_DATASETUP;
		break;
	case IDC_EDITWAVELENGTH:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_EDITWAVELENGTH;
		break;
	case IDC_UPDWAVELENGTH:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_UPDWAVELENGTH;
		break;
	case IDC_SETWAVELENGTH:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_SETWAVELENGTH;
		break;
	case IDC_MONOCHROMATORSETUP:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_MONOCHROMATORSETUP;
		break;
	case IDC_CHECKINTEGRATIONTIMEON:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_CHECKINTEGRATIONTIMEON;
		break;
	case IDC_SETUPAUTOINTEGRATIONTIME:
		pTTDispInfo->hinst = hInst;
		pTTDispInfo->lpszText = (LPTSTR)IDS_SETUPAUTOINTEGRATIONTIME;
		break;
	default:
		break;
	}
}

BOOL CDlgAcquisitionSetup::OnGetDefID()
{
	SHORT			shKeyState;			// state of the return key
	HWND			hwndFocus;			// control with the focus
	UINT			nID;				// control id with the focus

											// check the return key state
	shKeyState = GetKeyState(VK_RETURN);
	if (0 != (0x8000 & shKeyState))
	{
		// who has the focus?
		hwndFocus = GetFocus();
		nID = GetDlgCtrlID(hwndFocus);
		switch (nID)
		{
		case IDC_EDITSTARTWAVELENGTH:
			this->m_fAllowClose = FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITENDWAVELENGTH));
			return TRUE;
		case IDC_EDITENDWAVELENGTH:
			this->m_fAllowClose = FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITSTEPSIZE));
			return TRUE;
		case IDC_EDITSTEPSIZE:
			this->m_fAllowClose = FALSE;
			return TRUE;
		case IDC_EDITDWELLTIME:
			this->m_fAllowClose = FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITTIMEDELAYMULTIPLIER));
			return TRUE;
		case IDC_EDITTIMEDELAYMULTIPLIER:
			this->m_fAllowClose = FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITSIGNALAVERAGING));
			return TRUE;
		case IDC_EDITSIGNALAVERAGING:
			this->m_fAllowClose = FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITSCANAVERAGING));
			return TRUE;
		case IDC_EDITSCANAVERAGING:
			this->m_fAllowClose = FALSE;
			return TRUE;
		case IDC_EDITMULTIPLESCANS:
			this->m_fAllowClose = FALSE;
			return TRUE;
		case IDC_EDITWAVELENGTH:
			this->m_fAllowClose = FALSE;
			return TRUE;
		case IDC_EDITWORKINGDIRECTORY:
			this->m_fAllowClose = FALSE;
			return TRUE;
		case IDC_EDITCOMMENT:
			this->m_fAllowClose = FALSE;
			this->SetComment();
			return TRUE;
		default:
			break;
		}
	}
	return FALSE;
}

// form bitmap for button
HBITMAP CDlgAcquisitionSetup::FormBitmapImage(
	HWND			hwnd,
	COLORREF		colBackground,
	LPCTSTR			szText,
	COLORREF		colText)
{
	HDC					hdc;
	HDC					hdcMem;
	HBITMAP				hbmOld;					// memory bitmap
	RECT				rc;			// window rectangle
	HBRUSH				hbrBackground;			// background brush
	HBITMAP				hbm;									// returned bitmap
	TCHAR				szString[MAX_PATH];
	size_t				slen;
	SIZE				sz;
	UINT				ta;
	int					x, y;
	int					bkMode;									// background mode
	COLORREF			oldTextColor;							// previous text colour

	hdc = GetDC(hwnd);
	GetClientRect(hwnd, &rc);
	hdcMem = CreateCompatibleDC(hdc);
	hbm = CreateCompatibleBitmap(hdc, rc.right - rc.left, rc.bottom - rc.top);
	ReleaseDC(hwnd, hdc);
	hbmOld = (HBITMAP)SelectObject(hdcMem, (HGDIOBJ)hbm);
	// fill in the background in green
	hbrBackground = CreateSolidBrush(colBackground);
	FillRect(hdcMem, &rc, hbrBackground);
	DeleteObject(hbrBackground);
	// center the text in the window
	StringCchCopy(szString, MAX_PATH, szText);
	StringCchLength(szString, MAX_PATH, &slen);
	GetTextExtentPoint32(hdcMem, szString, slen, &sz);
	// align the text
	ta = SetTextAlign(hdcMem, TA_BOTTOM | TA_LEFT);
	x = ((rc.right - rc.left) - sz.cx) / 2;
	y = rc.bottom - (((rc.bottom - rc.top) - sz.cy) / 2);
	// set the background mode
	bkMode = SetBkMode(hdcMem, TRANSPARENT);
	// output the text
	oldTextColor = SetTextColor(hdcMem, colText);
	TextOut(hdcMem, x, y, szString, slen);
	SetTextColor(hdcMem, oldTextColor);
	// reset the bkmode and text alignment
	SetBkMode(hdcMem, bkMode);
	SetTextAlign(hdcMem, ta);
	// cleanup
	SelectObject(hdcMem, (HGDIOBJ)hbmOld);
	DeleteDC(hdcMem);
	return hbm;
}

// display start wave
void CDlgAcquisitionSetup::DisplayStartWave()
{
	TCHAR			szString[MAX_PATH];
//	_stprintf_s(szString, L"%5.1f nm", this->m_pMyObject->GetSciScan()->GetstartWave());
	_stprintf_s(szString, L"%5.1f nm", this->GetDoubleProperty(DISPID_StartWavelength));

	SetDlgItemText(this->m_hwndDlg, IDC_EDITSTARTWAVELENGTH, szString);
}

void CDlgAcquisitionSetup::SetStartWave()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	BOOL			fSuccess = FALSE;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITSTARTWAVELENGTH, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			fSuccess = TRUE;
//			this->m_pMyObject->GetSciScan()->SetstartWave(fval);
			this->SetDoubleProperty(DISPID_StartWavelength, fval);
		}
	}
	if (!fSuccess) this->DisplayStartWave();
}

// end wave
void CDlgAcquisitionSetup::DisplayEndWave()
{
	TCHAR			szString[MAX_PATH];
//	_stprintf_s(szString, L"%5.1f nm", this->m_pMyObject->GetSciScan()->GetEndWave());
	_stprintf_s(szString, L"%5.1f nm", this->GetDoubleProperty(DISPID_EndWavelength));
	SetDlgItemText(this->m_hwndDlg, IDC_EDITENDWAVELENGTH, szString);
}

void CDlgAcquisitionSetup::SetEndWave()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	BOOL			fSuccess = FALSE;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITENDWAVELENGTH, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			fSuccess = TRUE;
//			this->m_pMyObject->GetSciScan()->SetEndWave(fval);
			this->SetDoubleProperty(DISPID_EndWavelength, fval);

		}
	}
	if (!fSuccess) this->DisplayEndWave();
}

// step size	
void CDlgAcquisitionSetup::DisplayStepSize()
{
	TCHAR			szString[MAX_PATH];
//	_stprintf_s(szString, L"%4.2f nm", this->m_pMyObject->GetSciScan()->GetStepSize());
	_stprintf_s(szString, L"%4.2f nm", this->GetDoubleProperty(DISPID_StepSize));
	SetDlgItemText(this->m_hwndDlg, IDC_EDITSTEPSIZE, szString);
}

void CDlgAcquisitionSetup::SetStepSize()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	BOOL			fSuccess = FALSE;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITSTEPSIZE, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			fSuccess = TRUE;
//			this->m_pMyObject->GetSciScan()->SetStepSize(fval);
			this->SetDoubleProperty(DISPID_StepSize, fval);
		}
	}
	if (!fSuccess) this->DisplayStepSize();
}

// display number of steps
void CDlgAcquisitionSetup::DisplayNumberOfSteps()
{
//	SetDlgItemInt(this->m_hwndDlg, IDC_NUMBEROFSTEPS, this->m_pMyObject->GetSciScan()->GetNumberOfSteps(), TRUE);
	SetDlgItemInt(this->m_hwndDlg, IDC_NUMBEROFSTEPS, this->GetLongProperty(DISPID_NumberOfSteps), TRUE);
}

// dwell time
void CDlgAcquisitionSetup::DisplayDwellTime()
{
//	long	dwellTime = (long)floor((this->m_pMyObject->GetSciScan()->GetDwellTime() * 1000.0) + 0.5);
	long	dwellTime = this->GetLongProperty(DISPID_DwellTime);
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITDWELLTIME, dwellTime, FALSE);
}

void CDlgAcquisitionSetup::SetDwellTime()
{
	long	dwellTime = GetDlgItemInt(this->m_hwndDlg, IDC_EDITDWELLTIME, NULL, FALSE);
//	this->m_pMyObject->GetSciScan()->SetDwellTime(dwellTime / 1000.0);
	this->SetLongProperty(DISPID_DwellTime, dwellTime);
}

// Signal Averaging
void CDlgAcquisitionSetup::DisplaySignalAveraging()
{
//	SetDlgItemInt(this->m_hwndDlg, IDC_EDITSIGNALAVERAGING,
//		this->m_pMyObject->GetSciScan()->GetSignalAveraging(), FALSE);
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITSIGNALAVERAGING,
		this->GetLongProperty(DISPID_SignalAveraging), FALSE);
}

void CDlgAcquisitionSetup::SetSignalAveraging()
{
	short int sval = GetDlgItemInt(this->m_hwndDlg, IDC_EDITSIGNALAVERAGING, NULL, FALSE);
//	this->m_pMyObject->GetSciScan()->SetSignalAveraging(sval);
	this->SetLongProperty(DISPID_SignalAveraging, sval);
}

// Scan Averaging
void CDlgAcquisitionSetup::DisplayScanAveraging()
{
//	SetDlgItemInt(this->m_hwndDlg, IDC_EDITSCANAVERAGING,
//		this->m_pMyObject->GetSciScan()->GetScanAveraging(), FALSE);
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITSCANAVERAGING,
		this->GetLongProperty(DISPID_ScanAveraging), FALSE);
}

void CDlgAcquisitionSetup::SetScanAveraging()
{
	short int sval = GetDlgItemInt(this->m_hwndDlg, IDC_EDITSCANAVERAGING, NULL, FALSE);
//	this->m_pMyObject->GetSciScan()->SetScanAveraging(sval);
	this->SetLongProperty(DISPID_ScanAveraging, sval);
}

// transition delay multiplier
void CDlgAcquisitionSetup::DisplayTransitionDelayMultiplier()
{
//	SetDlgItemInt(this->m_hwndDlg, IDC_EDITTIMEDELAYMULTIPLIER,
//		this->m_pMyObject->GetSciScan()->GetTransitionDelayMultiplier(), FALSE);
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITTIMEDELAYMULTIPLIER,
		this->GetLongProperty(DISPID_TransitionDelayMultiplier), FALSE);

}

void CDlgAcquisitionSetup::SetTransitionDelayMultiplier()
{
	long	TDM = GetDlgItemInt(this->m_hwndDlg, IDC_EDITTIMEDELAYMULTIPLIER, NULL, FALSE);
//	this->m_pMyObject->GetSciScan()->SetTransitionDelayMultiplier(TDM);
	this->SetLongProperty(DISPID_TransitionDelayMultiplier, TDM);
}

// multiple scans
void CDlgAcquisitionSetup::DisplayMultipleScans()
{
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITMULTIPLESCANS, 1, FALSE);
}

void CDlgAcquisitionSetup::SetMultipleScans()
{
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITMULTIPLESCANS, 1, FALSE);
}

void CDlgAcquisitionSetup::DisplayCurrentMode()
{
//	if (this->m_pMyObject->GetSciInput()->GetPulsedMode())
	if (this->GetPulsedMode())
	{
		// pulse mode is selected
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKPULSEDMODE), BST_CHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKCONTINUOUS), BST_UNCHECKED);
		// disable DAC setup button
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_DACSETUP), FALSE);
		// only allow sample acquisition
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSAMPLE), BST_CHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINTENSITYCALIBRATION), BST_UNCHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKBACKGROUND), BST_UNCHECKED);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHKINTENSITYCALIBRATION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHKBACKGROUND), FALSE);
		// integration time disabled
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), BST_UNCHECKED);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETUPAUTOINTEGRATIONTIME), FALSE);
	}
	else
	{
		// continuous mode
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKPULSEDMODE), BST_UNCHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKCONTINUOUS), BST_CHECKED);
		// enable DAC setup button
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_DACSETUP), TRUE);
		// enable background and calibration scans
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHKINTENSITYCALIBRATION), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHKBACKGROUND), TRUE);
		// integration time enabled
	//	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), BST_CHECKED);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETUPAUTOINTEGRATIONTIME), TRUE);
	}
}

void CDlgAcquisitionSetup::OnClickedMode(
	UINT		nID)
{
//	this->m_pMyObject->GetSciInput()->SetPulsedMode(IDC_CHKPULSEDMODE == nID);
	this->SetPulsedMode(IDC_CHKPULSEDMODE == nID);
	this->DisplayCurrentMode();
}

// pulsed mode
BOOL CDlgAcquisitionSetup::GetPulsedMode()
{
	BOOL		fAvail = FALSE;
	return this->m_pMyObject->FireRequestGetPulsedMode(&fAvail);
//	return 0 != (this->GetLongProperty(DISPID_PulsedMode));
}
void CDlgAcquisitionSetup::SetPulsedMode(
	BOOL			fPulsedMode)
{
//	this->SetLongProperty(DISPID_PulsedMode, fPulsedMode ? 1 : 0);
	this->m_pMyObject->FireRequestSetPulsedMode(fPulsedMode);
}

// current wavelength
void CDlgAcquisitionSetup::DisplayCurrentWavelength()
{
	TCHAR			szString[MAX_PATH];
//	_stprintf_s(szString, L"%5.1f nm", this->m_pMyObject->GetMonoHandler()->GetCurrentWavelength());
	_stprintf_s(szString, L"%5.1f nm", this->m_pMyObject->FireRequestGetWavelength());
	SetDlgItemText(this->m_hwndDlg, IDC_EDITWAVELENGTH, szString);
}

void CDlgAcquisitionSetup::SetCurrentWavelength()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITWAVELENGTH, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
//			this->m_pMyObject->GetMonoHandler()->SetCurrentWavelength(fval);
			this->m_pMyObject->FireRequestSetWavelength(fval);
		}
	}
	this->DisplayCurrentWavelength();
}

void CDlgAcquisitionSetup::OnDeltaPosCurrentWave(
	long		iDelta)
{
//	double			currentWavelength = this->m_pMyObject->GetMonoHandler()->GetCurrentWavelength() - (iDelta * 1.0);
//	this->m_pMyObject->GetMonoHandler()->SetCurrentWavelength(currentWavelength);
	double		currentWavelength = this->m_pMyObject->FireRequestGetWavelength() - (iDelta * 1.0);
	this->m_pMyObject->FireRequestSetWavelength(currentWavelength);
	this->DisplayCurrentWavelength();
}

void CDlgAcquisitionSetup::OnMonochromatorSetup()
{
//	this->m_pMyObject->GetMonoHandler()->DoSetup();
	this->m_pMyObject->FireRequestSetupDisplay(L"Monochromator");
}


// read the current signal
void CDlgAcquisitionSetup::SetReadingTimerOnOff(
			BOOL		fTimerOn)
{
	if (fTimerOn)
	{
		SetTimer(this->m_hwndDlg, TMR_READING, 100, NULL);
	}
	else
	{
		KillTimer(this->m_hwndDlg, TMR_READING);
	}
}

void CDlgAcquisitionSetup::OnTimerReadSignal()
{
/*
	double			signal = this->m_pMyObject->GetSciInput()->ReadSignal();
	double			voltageRange[2];
	this->m_pMyObject->GetSciInput()->GetVoltageRange(&voltageRange[0], &voltageRange[1]);
	if (voltageRange[1] > voltageRange[0])
	{
		long			percent = (long)floor(((signal - voltageRange[0]) / (voltageRange[1] - voltageRange[0]) * 100.0) + 0.5);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_PRGSIGNALLEVEL), PBM_SETPOS, (WPARAM)percent, 0);
	}
	*/
	double			percent;
	this->m_pMyObject->FireRequestReadSignal(&percent);
	SendMessage(GetDlgItem(this->m_hwndDlg, IDC_PRGSIGNALLEVEL), PBM_SETPOS, (WPARAM)
		(long) floor(percent + 0.5), 0);
}

// display the scan type
void CDlgAcquisitionSetup::DisplayScanType()
{
	IDispatch	*	pdisp;
	HRESULT			hr;
	VARIANT			varResult;
	TCHAR			szScanType[MAX_PATH];

	// set the default
	StringCchCopy(szScanType, MAX_PATH, L"Sample");
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_ScanType, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, szScanType, MAX_PATH);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSAMPLE), 0 == lstrcmpi(szScanType, L"Sample") ? BST_CHECKED : BST_UNCHECKED);
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINTENSITYCALIBRATION), 0 == lstrcmpi(szScanType, L"Intensity Calibration") ? BST_CHECKED : BST_UNCHECKED);
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKBACKGROUND), 0 == lstrcmpi(szScanType, L"Background") ? BST_CHECKED : BST_UNCHECKED);
}
// set the scan type
void CDlgAcquisitionSetup::SetScanType(
	UINT		controlID)
{
	TCHAR			szScanType[MAX_PATH];
	IDispatch	*	pdisp;

	StringCchCopy(szScanType, MAX_PATH, IDC_CHKSAMPLE == controlID ? L"Sample" :
		(IDC_CHKINTENSITYCALIBRATION == controlID ? L"Intensity Calibration" : L"Background"));
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetStringProperty(pdisp, DISPID_ScanType, szScanType);
		pdisp->Release();
	}
}

// working directory
void CDlgAcquisitionSetup::DisplayWorkingDirectory()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	TCHAR			workingDirectory[MAX_PATH];
	size_t			slen;
	HWND			hwnd;
	HDC				hdc;
	RECT			rc;

	workingDirectory[0] = '\0';
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_WorkingDirectory, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, workingDirectory, MAX_PATH);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	StringCchLength(workingDirectory, MAX_PATH, &slen);
	if (slen > 0)
	{
		// make sure that the text fits in the window
		hwnd = GetDlgItem(this->m_hwndDlg, IDC_EDITWORKINGDIRECTORY);
		hdc = GetDC(hwnd);
		GetClientRect(hwnd, &rc);
		PathCompactPath(hdc, workingDirectory, rc.right - rc.left);
		ReleaseDC(hwnd, hdc);
		SetDlgItemText(this->m_hwndDlg, IDC_EDITWORKINGDIRECTORY, workingDirectory);
	}
	else
	{
		// clear the text
		SetDlgItemText(this->m_hwndDlg, IDC_EDITWORKINGDIRECTORY, L"");
	}
}

void CDlgAcquisitionSetup::SetWorkingDirectory()
{
}

void CDlgAcquisitionSetup::OnBrowseWorkingDirectory()
{
	// use fileOpen object to select the working directory
	HRESULT				hr;
	CLSID				clsid;
	IDispatch		*	pdisp	= NULL;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	TCHAR				szWorkingDirectory[MAX_PATH];

	hr = CLSIDFromProgID(L"Sciencetech.SciFileOpen.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdisp, L"SelectWorkingDirectory", &dispid);
		InitVariantFromInt32((long) this->m_hwndDlg, &varg);
		hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		if (fSuccess)
		{
			Utils_GetMemid(pdisp, L"WorkingDirectory", &dispid);
			hr = Utils_InvokePropertyGet(pdisp, dispid, NULL, 0, &varResult);
			if (SUCCEEDED(hr))
			{
				VariantToString(varResult, szWorkingDirectory, MAX_PATH);
				this->SetWorkingDirectory(szWorkingDirectory);
				VariantClear(&varResult);
			}
		}
		pdisp->Release();
	}
}

void CDlgAcquisitionSetup::SetWorkingDirectory(
	LPCTSTR		szWorkingDirectory)
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetStringProperty(pdisp, DISPID_WorkingDirectory, szWorkingDirectory);
		pdisp->Release();
	}
}

BOOL CDlgAcquisitionSetup::WorkingDirectorySet()		// has the working directory been defined?
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	TCHAR			workingDirectory[MAX_PATH];
	size_t			slen;
	workingDirectory[0] = '\0';
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_WorkingDirectory, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, workingDirectory, MAX_PATH);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	StringCchLength(workingDirectory, MAX_PATH, &slen);
	return slen > 0;
}

// comment
void CDlgAcquisitionSetup::DisplayComment()
{
//	TCHAR			szComment[MAX_PATH];
//	this->m_pMyObject->GetOptFile()->GetComment(szComment, MAX_PATH);
	IDispatch	*	pdisp;
	HRESULT			hr;
	VARIANT			varResult;
	TCHAR			szComment[MAX_PATH];
	szComment[0] = '\0';
	if (this->GetOurObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_Comment, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, szComment, MAX_PATH);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	SetDlgItemText(this->m_hwndDlg, IDC_EDITCOMMENT, szComment);
}

void CDlgAcquisitionSetup::SetComment()
{
	TCHAR			szComment[MAX_PATH];
	IDispatch	*	pdisp;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITCOMMENT, szComment, MAX_PATH) > 0)
	{
//		this->m_pMyObject->GetOptFile()->SetComment(szComment);
		if (this->GetOurObject(&pdisp))
		{
			Utils_SetStringProperty(pdisp, DISPID_Comment, szComment);
			pdisp->Release();
		}
	}
}

// get our object
BOOL CDlgAcquisitionSetup::GetOurObject(
	IDispatch	**	ppdisp)
{
	HRESULT			hr;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)ppdisp);
	return SUCCEEDED(hr);
}

// Auto Integration time
void CDlgAcquisitionSetup::DisplayAutoIntegrationTime()
{
//	if (this->m_pMyObject->GetSciInput()->GetIntegrationAvailable())
	if (this->GetAutoIntegrationAvailable())
	{
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETUPAUTOINTEGRATIONTIME), TRUE);
//		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON),
//			this->m_pMyObject->GetSciInput()->GetAutoIntegrationTime() ? BST_CHECKED : BST_UNCHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON),
			this->GetAutoIntegrationTime() ? BST_CHECKED : BST_UNCHECKED);
	}
	else
	{
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETUPAUTOINTEGRATIONTIME), FALSE);
	}
}

void CDlgAcquisitionSetup::OnClickedCheckIntegrationTimeOn()		// toggle auto integration time on / off
{
	BOOL		fAuto = BST_CHECKED == Button_GetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTEGRATIONTIMEON));
//	this->m_pMyObject->GetSciInput()->SetAutoIntegrationTime(!fAuto);
	this->SetAutoIntegrationTime(!fAuto);
	this->DisplayAutoIntegrationTime();
}

void CDlgAcquisitionSetup::OnSetupAutoIntegrationTime()
{
//	this->m_pMyObject->GetSciInput()->SetupIntegrationTime(this->m_hwndDlg);
	if (this->GetAutoIntegrationAvailable())
		this->m_pMyObject->FireRequestSetupDisplay(L"AutoIntegrationTime");
}

// Auto integration time flag
BOOL CDlgAcquisitionSetup::GetAutoIntegrationAvailable()
{
	BOOL			fAvail = FALSE;
	this->m_pMyObject->FireRequestGetAutoTimeConstant(&fAvail);
	return fAvail;
}

BOOL CDlgAcquisitionSetup::GetAutoIntegrationTime()
{
	BOOL			fAvail = FALSE;
	return this->m_pMyObject->FireRequestGetAutoTimeConstant(&fAvail);
}

void CDlgAcquisitionSetup::SetAutoIntegrationTime(
	BOOL			fAutoIntegrationTime)
{
	this->m_pMyObject->FireRequestSetAutoTimeConstant(fAutoIntegrationTime);
}

// sink
void CDlgAcquisitionSetup::ConnectSink()
{
	HRESULT			hr;
	CImpISink	*	pSink;
	IUnknown	*	pUnkSink;

	pSink = new CImpISink(this);
	hr = pSink->QueryInterface(MY_IIDSINK, (LPVOID*)&pUnkSink);
	if (SUCCEEDED(hr))
	{
		Utils_ConnectToConnectionPoint(this->m_pMyObject, pUnkSink, MY_IIDSINK, TRUE, &this->m_dwCookie);
		pUnkSink->Release();
	}
	else delete pSink;
}

void CDlgAcquisitionSetup::DisconnectSink()
{
	Utils_ConnectToConnectionPoint(this->m_pMyObject, NULL, MY_IIDSINK, FALSE, &this->m_dwCookie);
}

// sink events
void CDlgAcquisitionSetup::OnSciScanPropChanged(
	LPCTSTR		szPropName)
{
	if (0 == lstrcmpi(szPropName, L"startWave"))
	{
		this->DisplayStartWave();
		this->DisplayNumberOfSteps();
	}
	else if (0 == lstrcmpi(szPropName, L"EndWave"))
	{
		this->DisplayEndWave();
		this->DisplayNumberOfSteps();
	}
	else if (0 == lstrcmpi(szPropName, L"StepSize"))
	{
		this->DisplayStepSize();
		this->DisplayNumberOfSteps();
	}
	else if (0 == lstrcmpi(szPropName, L"TransitionDelayMultiplier"))
	{
		this->DisplayTransitionDelayMultiplier();
	}
	else if (0 == lstrcmpi(szPropName, L"DwellTime"))
	{
		this->DisplayDwellTime();
	}
	else if (0 == lstrcmpi(szPropName, L"SignalAveraging"))
	{
		this->DisplaySignalAveraging();
	}
	else if (0 == lstrcmpi(szPropName, L"ScanAveraging"))
	{
		this->DisplayScanAveraging();
	}
}

void CDlgAcquisitionSetup::OnObjectPropChanged(
	LPCTSTR		ObjectName,
	LPCTSTR		Property)
{
	if (0 == lstrcmpi(ObjectName, L"OptFile") && 0 == lstrcmpi(Property, L"Comment"))
	{
		this->DisplayComment();
	}
}

// property change notifications
void CDlgAcquisitionSetup::ConnectPropNotify()
{
	HRESULT				hr;
	CImpIPropNotify*	pPropNotify;
	IUnknown		*	pUnkSink;
	pPropNotify = new CImpIPropNotify(this);
	hr = pPropNotify->QueryInterface(IID_IPropertyNotifySink, (LPVOID*)&pUnkSink);
	if (SUCCEEDED(hr))
	{
		Utils_ConnectToConnectionPoint(this->m_pMyObject, pUnkSink, IID_IPropertyNotifySink, TRUE, &this->m_dwPropChange);
		pUnkSink->Release();
	}
	else delete pPropNotify;
}

void CDlgAcquisitionSetup::DisconnectPropNotify()
{
	Utils_ConnectToConnectionPoint(this->m_pMyObject, NULL, IID_IPropertyNotifySink, FALSE, &this->m_dwPropChange);
}

void CDlgAcquisitionSetup::OnPropChanged(
	DISPID			dispid)
{
	if (DISPID_ScanType == dispid)
	{
		this->DisplayScanType();
	}
	else if (DISPID_WorkingDirectory == dispid)
	{
		this->DisplayWorkingDirectory();
	}
	else if (DISPID_StartWavelength == dispid)
	{
		this->DisplayStartWave();
		this->DisplayNumberOfSteps();
	}
	else if (DISPID_EndWavelength == dispid)
	{
		this->DisplayEndWave();
		this->DisplayNumberOfSteps();
	}
	else if (DISPID_StepSize == dispid)
	{
		this->DisplayStepSize();
		this->DisplayNumberOfSteps();
	}
}

double CDlgAcquisitionSetup::GetDoubleProperty(
	DISPID			dispid)
{
	double			dval = 0.0;
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		dval = Utils_GetDoubleProperty(pdisp, dispid);
		pdisp->Release();
	}
	return dval;
}

void CDlgAcquisitionSetup::SetDoubleProperty(
	DISPID			dispid,
	double			dval)
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetDoubleProperty(pdisp, dispid, dval);
		pdisp->Release();
	}
}

long CDlgAcquisitionSetup::GetLongProperty(
	DISPID			dispid)
{
	long			lval = 0;
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		lval = Utils_GetLongProperty(pdisp, dispid);
		pdisp->Release();
	}
	return lval;
}

void CDlgAcquisitionSetup::SetLongProperty(
	DISPID			dispid,
	long			lval)
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetLongProperty(pdisp, dispid, lval);
		pdisp->Release();
	}
}


CDlgAcquisitionSetup::CImpISink::CImpISink(CDlgAcquisitionSetup * pDlgAcquisitionSetup) :
	m_pDlgAcquisitionSetup(pDlgAcquisitionSetup),
	m_cRefs(0)
{

}

CDlgAcquisitionSetup::CImpISink::~CImpISink()
{

}

// IUnknown methods
STDMETHODIMP CDlgAcquisitionSetup::CImpISink::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || MY_IIDSINK == riid)
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

STDMETHODIMP_(ULONG) CDlgAcquisitionSetup::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDlgAcquisitionSetup::CImpISink::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CDlgAcquisitionSetup::CImpISink::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CDlgAcquisitionSetup::CImpISink::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CDlgAcquisitionSetup::CImpISink::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CDlgAcquisitionSetup::CImpISink::Invoke(
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
	case DISPID_SciScanPropChanged:
		this->OnSciScanPropChanged(pDispParams);
		break;
	case DISPID_ObjectPropChanged:
		this->OnObjectPropChanged(pDispParams);
		break;
	default:
		break;
	}
	return S_OK;
}

void CDlgAcquisitionSetup::CImpISink::OnSciScanPropChanged(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (SUCCEEDED(hr))
	{
		this->m_pDlgAcquisitionSetup->OnSciScanPropChanged(varg.bstrVal);
		VariantClear(&varg);
	}
}

void CDlgAcquisitionSetup::CImpISink::OnObjectPropChanged(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	TCHAR			ObjectName[MAX_PATH];
	TCHAR			Property[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (SUCCEEDED(hr))
	{
		StringCchCopy(ObjectName, MAX_PATH, varg.bstrVal);
		VariantClear(&varg);
		hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	}
	if (SUCCEEDED(hr))
	{
		StringCchCopy(Property, MAX_PATH, varg.bstrVal);
		VariantClear(&varg);
		this->m_pDlgAcquisitionSetup->OnObjectPropChanged(ObjectName, Property);
	}
}

CDlgAcquisitionSetup::CImpIPropNotify::CImpIPropNotify(CDlgAcquisitionSetup * pDlgAcquisitionSetup) :
	m_pDlgAcquisitionSetup(pDlgAcquisitionSetup),
	m_cRefs(0)
{

}

CDlgAcquisitionSetup::CImpIPropNotify::~CImpIPropNotify()
{

}

// IUnknown methods
STDMETHODIMP CDlgAcquisitionSetup::CImpIPropNotify::QueryInterface(
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

STDMETHODIMP_(ULONG) CDlgAcquisitionSetup::CImpIPropNotify::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDlgAcquisitionSetup::CImpIPropNotify::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IPropertyNotifySink methods
STDMETHODIMP CDlgAcquisitionSetup::CImpIPropNotify::OnChanged(
	DISPID			dispID)
{
	this->m_pDlgAcquisitionSetup->OnPropChanged(dispID);
	return S_OK;
}

STDMETHODIMP CDlgAcquisitionSetup::CImpIPropNotify::OnRequestEdit(
	DISPID			dispID)
{
	return S_OK;
}