// dispids.h
// dispatch ids

#pragma once

// properties
//#define				DISPID_SciScan			0x0100
//#define				DISPID_SciInput			0x0101
//#define				DISPID_FilterWheel		0x0102
//#define				DISPID_SlitServer		0x0103
//#define				DISPID_MonoHandler		0x0104
//#define				DISPID_OptFile			0x0105
#define				DISPID_ScanType			0x0106
#define				DISPID_WorkingDirectory	0x0107
#define				DISPID_Comment			0x0108
#define				DISPID_StartWavelength	0x0109
#define				DISPID_EndWavelength	0x010a
#define				DISPID_StepSize			0x010b
#define				DISPID_NumberOfSteps	0x010c
#define				DISPID_DwellTime		0x010d
#define				DISPID_TransitionDelayMultiplier	0x010e
#define				DISPID_SignalAveraging	0x010f
#define				DISPID_ScanAveraging	0x0110
#define				DISPID_MultipleScans	0x0111

// methods
#define				DISPID_DoScanSetup		0x0120

// events
#define				DISPID_SciScanPropChanged	0x0140
#define				DISPID_RequestFriendlyName	0x0141
#define				DISPID_ObjectPropChanged	0x0142
#define				DISPID_RequestSetupDisplay	0x0143
#define				DISPID_RequestReadSignal	0x0144
#define				DISPID_RequestGetWavelength	0x0145
#define				DISPID_RequestSetWavelength	0x0146
#define				DISPID_RequestGetAutoTimeConstant	0x0147
#define				DISPID_RequestSetAutoTimeConstant	0x0148
#define				DISPID_RequestGetPulsedMode	0x0149
#define				DISPID_RequestSetPulsedMode	0x014a