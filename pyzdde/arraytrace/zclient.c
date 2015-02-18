// Zclient: ZEMAX client template program
// Originally written by Kenneth Moore June 1997
// Copyright 1997-2006 Kenneth Moore
//
// Normally, none of this code needs to be modified. Simply include this file and
// compile and link with the code that contains "UserFunction".
// The zclient program is responsible for establishing communication
// with the ZEMAX server. All data from ZEMAX can be obtained by calling
// PostRequestMessage or PostArrayTraceMessage with the item name and a buffer to hold the data.
//
// Zclient will call UserFunction when the DDE communication link is established and ready.
// Zclient will automatically terminate the connection when UserFunction returns.
//
// Version 1.1 modified to support Array ray tracing September, 1997
// Version 1.2 modified for faster execution October, 1997
// Version 1.3 modified for faster execution November, 1997
// Version 1.4 modified to fix memory leak January, 1998
// Version 1.5 modified to add support for long path names and quotes November, 1998
// Version 1.6 modified to fix missing support for long path names and quotes in MakeEmptyWindow March, 1999
// Version 1.7 modified to fix memory leak in WM_DDE_ACK, March 1999
// Version 1.8 modified to add E-field data to DDERAYDATA for ZEMAX 10.0, December 2000
// Version 1.9 modified PostRequestMessage and PostArrayTraceMessage to return -1 if data failed (usually because of a timeout) or 0 otherwise, April 2001
// Version 2.0 modified WM_USER_INITIATE to distingush between 2 possibly simultaneous copies of ZEMAX running when responding to UDOP calls, September 1, 2006
// Version 2.1 modified to support Visual Studio 2005. Added the #pragma to disable the warning about deprecated functions
// Version 2.2 modified to move GotData=0 to more robust position. If ZEMAX returns data very quickly a deadlock can occur. November 30, 2007

//  Version 2.3 modified the typecast of uiLow and uiHi from UINT to UINT_PTR.  CODE ONLY WORKS IN 64-BIT (Debug or Release).  Recast (int) msg.wParam as a return argument for WINAPI

#include <windows.h>
#include <dde.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <math.h>

#define WM_USER_INITIATE (WM_USER + 1)
#define DDE_TIMEOUT 50000
#pragma warning ( disable : 4996 ) // functions like strcpy are now deprecated for security reasons; this disables the warning

typedef struct
	{
	double x, y, z, l, m, n, opd, intensity;
	double Exr, Exi, Eyr, Eyi, Ezr, Ezi;
	int wave, error, vigcode, want_opd;
	}DDERAYDATA;

LRESULT CALLBACK WndProc (HWND, UINT, WPARAM, LPARAM);
void WaitForData(HWND hwnd);
char *GetString(char *szBuffer, int n, char *szSubString);
void remove_quotes(char *s);
int  PostRequestMessage(char *szItem, char *szBuffer);
int  PostArrayTraceMessage(char *szBuffer, DDERAYDATA *RD);
void CenterWindow(HWND hwnd);
void UserFunction(char *szCommandLine);
void MakeEmptyWindow(int text, char *szAppName, char *szOptions);
void Get_2_5_10(double cmax, double *cscale);

/* global variables used by the client code */
char szAppName[] = "ZemaxClient";
int GotData, ngNumRays, ZEMAX_INSTANCE = 0;
char szGlobalBuffer[5000], szCommandLine[260];
HINSTANCE globalhInstance;
HWND hwndServer, hwndClient;
DDERAYDATA *rdpGRD = NULL;

int WINAPI WinMain (HINSTANCE hInstance, HINSTANCE hPrevInstance, PSTR szCmdLine, int iCmdShow)
{
HWND hwnd;
MSG msg;
WNDCLASSEX wndclass;

wndclass.cbSize        = sizeof (wndclass);
wndclass.style         = CS_HREDRAW | CS_VREDRAW;
wndclass.lpfnWndProc   = WndProc;
wndclass.cbClsExtra    = 0;
wndclass.cbWndExtra    = 0;
wndclass.hInstance     = hInstance;
wndclass.hIcon         = LoadIcon (NULL, IDI_APPLICATION);
wndclass.hCursor       = LoadCursor (NULL, IDC_ARROW);
wndclass.hbrBackground = (HBRUSH) GetStockObject (WHITE_BRUSH);
wndclass.lpszMenuName  = NULL;
wndclass.lpszClassName = szAppName;
wndclass.hIconSm       = LoadIcon (NULL, IDI_APPLICATION);
RegisterClassEx (&wndclass);

globalhInstance = hPrevInstance;

if (iCmdShow)
	{
	// do nothing. This argument is unused, and is only referenced here to avoid a compiler warning about unused function arguments.
   }

strcpy(szCommandLine, szCmdLine);

hwnd = CreateWindow (szAppName, "ZEMAX Client", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);
UpdateWindow (hwnd);
SendMessage (hwnd, WM_USER_INITIATE, 0, 0L);

while (GetMessage (&msg, NULL, 0, 0))
	{
	TranslateMessage (&msg);
	DispatchMessage (&msg);
	}
return (int) msg.wParam;
}

LRESULT CALLBACK WndProc (HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
ATOM          aApp, aTop, aItem;
DDEACK        DdeAck;
DDEDATA      *pDdeData;
GLOBALHANDLE  hDdeData;
WORD          wStatus;
UINT_PTR      uiLow, uiHi;

switch (iMsg)
	{
   case WM_CREATE :
		hwndServer = 0;
      return 0;

	case WM_USER_INITIATE :
   	/* find ZEMAX */
	
		// code added September 1, 2006 to identify which ZEMAX is calling us, in case more than 1 copy of ZEMAX is running
		// this currently is only supported by user defined operands.
      // aApp = GlobalAddAtom ("ZEMAX");
		if (1)
			{
			char szSub[500];
			GetString(szCommandLine, 5, szSub);
			if (strcmp(szSub,"ZEMAX1") == 0 || strcmp(szSub,"ZEMAX2") == 0)
				{
				if (strcmp(szSub,"ZEMAX1") == 0)
					{
					aApp = GlobalAddAtom ("ZEMAX1");
					ZEMAX_INSTANCE = 1;
					}
				if (strcmp(szSub,"ZEMAX2") == 0)
					{
					aApp = GlobalAddAtom ("ZEMAX2");
					ZEMAX_INSTANCE = 2;
					}
				}
			else
				{
		      aApp = GlobalAddAtom ("ZEMAX");
				ZEMAX_INSTANCE = 0;
				}
			}

      aTop = GlobalAddAtom ("RayData");

		SendMessage (HWND_BROADCAST, WM_DDE_INITIATE, (WPARAM) hwnd, MAKELONG (aApp, aTop));

		/* delete the atoms */
      GlobalDeleteAtom (aApp);
      GlobalDeleteAtom (aTop);

      /* If no response, terminate */
      if (hwndServer == NULL)
      	{
		   MessageBox (hwnd, "Cannot communicate with ZEMAX!", "Hello?", MB_ICONEXCLAMATION | MB_OK);
         DestroyWindow(hwnd);
         return 0;
         }

		hwndClient = hwnd;

      UserFunction(szCommandLine);

      /* terminate the DDE connection */
   	PostMessage(hwndServer, WM_DDE_TERMINATE, (WPARAM) hwnd, 0L);
      hwndServer = NULL;

      /* now TERMINATE! */
      DestroyWindow(hwnd);
      return 0;

   case WM_DDE_DATA :
   	/* here comes the data! */
      // wParam -- sending window handle
      // lParam -- DDEDATA memory handle & item atom
      UnpackDDElParam(WM_DDE_DATA, lParam, &uiLow, &uiHi);
      FreeDDElParam(WM_DDE_DATA, lParam);
		hDdeData  = (GLOBALHANDLE) uiLow;
      pDdeData = (DDEDATA *) GlobalLock (hDdeData);
      aItem     = (ATOM) uiHi;

      // Initialize DdeAck structure
      DdeAck.bAppReturnCode = 0;
      DdeAck.reserved       = 0;
      DdeAck.fBusy          = FALSE;
      DdeAck.fAck           = FALSE;


      // Check for matching format, put the data in the buffer
      if (pDdeData->cfFormat == CF_TEXT)
      	{       
         /* get the data back into RD */
			if (rdpGRD) memcpy(rdpGRD, (DDERAYDATA *) pDdeData->Value, (ngNumRays+1)*sizeof(DDERAYDATA));
			else strcpy(szGlobalBuffer, (char *) pDdeData->Value);
         }

      GotData = 1;
		GlobalDeleteAtom (aItem);

      // Acknowledge if necessary
      if (pDdeData->fAckReq == TRUE)
      	{
         wStatus = *((WORD *) &DdeAck);
         if (!PostMessage ((HWND) wParam, WM_DDE_ACK, (WPARAM) hwnd, PackDDElParam (WM_DDE_ACK, wStatus, aItem)))
         	{
				if (hDdeData)
					{
					GlobalUnlock (hDdeData);
					GlobalFree (hDdeData);
					}
            return 0;
            }
         }

      // Clean up
		GlobalUnlock (hDdeData);
      if (pDdeData->fRelease == TRUE || DdeAck.fAck == FALSE) GlobalFree (hDdeData);
      return 0;

   case WM_DDE_ACK:
   	/* we are receiving an acknowledgement */
      /* the only one we care about is in response to the WM_DDE_INITIATE; otherwise free just the memory */
      if (hwndServer == NULL)
      	{
			uiLow = (UINT) NULL;
			uiHi = (UINT) NULL;
         UnpackDDElParam(WM_DDE_ACK, lParam, &uiLow, &uiHi);
         FreeDDElParam(WM_DDE_ACK, lParam);
         hwndServer = (HWND) wParam;
         if (uiLow) GlobalDeleteAtom((ATOM) uiLow);
         if (uiHi) GlobalDeleteAtom((ATOM) uiHi);
         }
		else
			{
			HWND dummy;
			uiLow = (UINT) NULL;
			uiHi = (UINT) NULL;
         UnpackDDElParam(WM_DDE_ACK, lParam, &uiLow, &uiHi);
         FreeDDElParam(WM_DDE_ACK, lParam);
         dummy = (HWND) wParam;
         if (uiLow) GlobalDeleteAtom((ATOM) uiLow);
         if (uiHi) GlobalDeleteAtom((ATOM) uiHi);
			}
   	return 0;

   case WM_DDE_TERMINATE :
   	PostMessage(hwndServer, WM_DDE_TERMINATE, (WPARAM) hwnd, 0L);
      hwndServer = NULL;
      return 0;

   case WM_PAINT :
   	{
      PAINTSTRUCT ps;
   	BeginPaint(hwnd, &ps);
      EndPaint(hwnd, &ps);
      }
      return 0;

   case WM_CLOSE :
   	PostMessage(hwndServer, WM_DDE_TERMINATE, (WPARAM) hwnd, 0L);
   	break;             // for default processing

   case WM_DESTROY :
   	PostQuitMessage(0);
      return 0;
   }
   return DefWindowProc(hwnd, iMsg, wParam, lParam);
}

void WaitForData(HWND hwnd)
{
int sleep_count;
MSG msg;
DWORD dwTime;
dwTime = GetCurrentTime();

sleep_count = 0;

while (!GotData)
	{
	while (PeekMessage (&msg, hwnd, WM_DDE_FIRST, WM_DDE_LAST, PM_REMOVE))
		{
		DispatchMessage (&msg);
		}
	/* Give the server a chance to respond */	
	Sleep(0);
	sleep_count++;
	if (sleep_count > 10000)
		{
		if (GetCurrentTime() - dwTime > DDE_TIMEOUT) return;
		sleep_count = 0;
		}
	}
}

char * GetString(char *szBuffer, int n, char *szSubString)
{
int i, j, k;
char szTest[5000];

szSubString[0] = '\0';
i = 0;
j = 0;
k = 0;
while (szBuffer[i] && (k <= n) )
	{
   szTest[j] = szBuffer[i];

   if (szBuffer[i] == '"')
   	{

      i++;
      j++;
      szTest[j] = szBuffer[i];

      /* we have a double quote; keep reading until EOF or another double quote */
      while(szBuffer[i] != '"' && szBuffer[i])
      	{
	      i++;
   	   j++;
		   szTest[j] = szBuffer[i];
         }
      }

   if (szTest[j] == ' ' || szTest[j] == '\n' || szTest[j] == '\r' || szTest[j] == '\0' || szTest[j] == ',')
   	{
      szTest[j] = '\0';
      if (k == n)
      	{
         strcpy(szSubString, szTest);
			return szSubString;
         }
      k++;
      j = -1;
      }
   i++;
   j++;
   }

szTest[j] = '\0';
if (k == n) strcpy(szSubString, szTest);

return szSubString;
}

int PostRequestMessage(char *szItem, char *szBuffer)
{
ATOM aItem;

aItem = GlobalAddAtom(szItem);

/* clear the buffers */
szGlobalBuffer[0] = '\0';
szBuffer[0] = '\0';

GotData = 0;

if (!PostMessage(hwndServer, WM_DDE_REQUEST, (WPARAM) hwndClient, PackDDElParam(WM_DDE_REQUEST, CF_TEXT, aItem)))
	{
   MessageBox (hwndClient, "Cannot communicate with ZEMAX!", "Hello?", MB_ICONEXCLAMATION | MB_OK);
   GlobalDeleteAtom(aItem);
   return -1;
   }

WaitForData(hwndClient);
strcpy(szBuffer, szGlobalBuffer);

if (GotData) return 0;
else return -1;
}

int PostArrayTraceMessage(char *szBuffer, DDERAYDATA *RD)
{
ATOM aItem;
HGLOBAL hPokeData;
DDEPOKE * lpPokeData;
long numbytes;
int numrays;


if (RD[0].opd > 4)
	{
	/* NSC Rays */
	numrays = (int)RD[0].opd - 5;
	}
else
	{
	/* sequential rays */
	numrays = RD[0].error;
	}

/* point to where the data is */
rdpGRD = RD;
ngNumRays = numrays;

numbytes = (1+numrays)*sizeof(DDERAYDATA);
hPokeData = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, (LONG) sizeof(DDEPOKE) + numbytes);
lpPokeData = (DDEPOKE *) GlobalLock(hPokeData);
lpPokeData->fRelease = TRUE;
lpPokeData->cfFormat = CF_TEXT;
memcpy(lpPokeData->Value, RD, numbytes);

/* clear the buffers */
szGlobalBuffer[0] = '\0';
szBuffer[0] = '\0';

aItem = GlobalAddAtom("RayArrayData");
GlobalUnlock(hPokeData);

GotData = 0;

if (!PostMessage(hwndServer, WM_DDE_POKE, (WPARAM) hwndClient, PackDDElParam(WM_DDE_POKE, (UINT) hPokeData, aItem)))
	{
   MessageBox (hwndClient, "Cannot communicate with ZEMAX!", "Hello?", MB_ICONEXCLAMATION | MB_OK);
   GlobalDeleteAtom(aItem);
	GlobalFree(hPokeData);
   return -1;
   }
GlobalDeleteAtom(aItem);

WaitForData(hwndClient);

strcpy(szBuffer, szGlobalBuffer);

/* clear the pointer */
rdpGRD = NULL;

if (GotData) return 0;
else return -1;
}

void MakeEmptyWindow(int text, char *szAppName, char *szOptions)
{
char szOutputFile[260], szModuleName[260], szBuffer[5000];
FILE *output;

/* get the output file name */
GetString(szCommandLine, 2, szOutputFile);

/* get the module name */
GetModuleFileName(globalhInstance, szModuleName, 255);

if ((output = fopen(szOutputFile, "wt")) == NULL)
	{
   /* can't open the file!! */
   return;
   }

if (text)
	{
   fputs("System is invalid, cannot compute data.\n",output);
   fclose(output);
	/* create a text window. Note we pass back the filename, module name, and activesurf as a single setting parameter. */
   sprintf(szBuffer,"MakeTextWindow,\"%s\",\"%s\",\"%s\",%s", szOutputFile, szModuleName, szAppName, szOptions);
   PostRequestMessage(szBuffer, szBuffer);
   }
else
	{
   fputs("NOFRAME\n",output);
   fputs("TEXT \"System is invalid, cannot compute data.\" .1 .5\n",output);
   fclose(output);
   sprintf(szBuffer,"MakeGraphicWindow,\"%s\",\"%s\",\"%s\",1,%s", szOutputFile, szModuleName, szAppName, szOptions);
   PostRequestMessage(szBuffer, szBuffer);
   }
}

void CenterWindow(HWND hwnd)
{
RECT rect;
int newx, newy;
GetWindowRect(hwnd, &rect);
newx = (GetSystemMetrics(SM_CXSCREEN) - (rect.right  - rect.left))/2;
newy = (GetSystemMetrics(SM_CYSCREEN) - (rect.bottom -  rect.top))/2;
SetWindowPos(hwnd, HWND_TOP, newx, newy, 0, 0, SWP_NOSIZE);
}

void Get_2_5_10(double cmax, double *cscale)
{
int i;
double temp;
if (cmax <= 0)
	{
	*cscale = .00001;
	return;
	}
*cscale = log10(cmax);
i = 0;
for (; *cscale < 0; i--) *cscale = *cscale + 1;
for (; *cscale > 1; i++) *cscale = *cscale - 1;
temp = 10;
if (*cscale < log10(5.0)) temp = 5;
if (*cscale < log10(2.0)) temp = 2;
*cscale = temp * pow(10, (double) i );
}

void remove_quotes(char *s)
{
int i=0;
/* remove the first quote if it exists */
if (s[0] == '"')
	{
	while (s[i])
		{
		s[i] = s[i+1];
		i++;
		}
	}
/* remove the last quote if it exists */
if (strlen(s) > 0)
	{
	if (s[strlen(s)-1] == '"') s[strlen(s)-1] = '\0';
	}
}
