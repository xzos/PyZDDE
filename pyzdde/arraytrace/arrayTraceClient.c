// The code here has been adapted from the C programs zclient.c and ArrayDemo.c,
// which were originally written by Kenneth Moore, and they are shipped with Zemax.
// zclient.c
// Originally written by Kenneth Moore June 1997
// Copyright 1997-2006 Kenneth Moore
// ArrayDemo.c sample program
// Written by Kenneth Moore March 1999
// The original zclient.c and ArrayDemo.c files are also available in the same 
// directory for reference
//
// How to call these functions from Python ? see
// http://scipy.github.io/old-wiki/pages/Cookbook/Ctypes#NumPy.27s_ndpointer_with_ctypes_argtypes

#include "arrayTraceClient.h"

/* global variables used by the client code */
char szAppName[] = "ZemaxClient";
int GotData, ngNumRays, ZEMAX_INSTANCE = 0;
char szGlobalBuffer[5000], szCommandLine[260];
HINSTANCE globalhInstance;
HWND hwndServer, hwndClient;
DDERAYDATA *rdpGRD = NULL;
DDERAYDATA *gPtr2RD = NULL;  /* used for passing the ray data array to the user function */
unsigned int DdeTimeout;
int RETVAL = 0;              /* Return value to Python indicating general error conditions*/
                             /* 0 = SUCCESS, -999 = Couldn't communicate with Zemax, -998 = timeout reached, etc*/

BOOL APIENTRY DllMain(HINSTANCE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        globalhInstance = (HINSTANCE)hModule;
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}


void rayTraceFunction(void)
{
    static char szBuffer[5000];
    int ret = 0;
    ret = PostArrayTraceMessage(szBuffer, gPtr2RD);
    /* clear the pointer */
    gPtr2RD = NULL;
}

/* ----------------------------------------------------------------------------------
  ArrayTrace functions that accept DDERAYDATA as argument
   ----------------------------------------------------------------------------------
 */

int __stdcall arrayTrace(DDERAYDATA * pRAD, unsigned int timeout)
{
    HWND hwnd;  /* handle to client window */
    MSG msg;
    WNDCLASSEX wndclass;

    wndclass.cbSize = sizeof(wndclass);
    wndclass.style = CS_HREDRAW | CS_VREDRAW;
    wndclass.lpfnWndProc = WndProc;
    wndclass.cbClsExtra = 0;
    wndclass.cbWndExtra = 0;
    wndclass.hInstance = globalhInstance;
    wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
    wndclass.lpszMenuName = NULL;
    wndclass.lpszClassName = szAppName;
    wndclass.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    RegisterClassEx(&wndclass);

    /* assign the pointer to ray data a global pointer so that rayTraceFunction() can access it*/
    gPtr2RD = pRAD;

    if (timeout > DDE_TIMEOUT)
        DdeTimeout = timeout;
    else
        DdeTimeout = DDE_TIMEOUT;

    hwnd = CreateWindow(szAppName, "ZEMAX Client", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT,
             CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, globalhInstance, NULL);
    UpdateWindow(hwnd);
    SendMessage(hwnd, WM_USER_INITIATE, 0, 0L);
  
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return RETVAL;
}

/* ----------------------------------------------------------------------------------
  ArrayTrace functions that accept Numpy arrays as arguments
  avoids large overhead times in Python wrapper functions
   ----------------------------------------------------------------------------------
 */

// ----------------------------------------------------------------------------------
// Mode 0: similar to GetTrace (rays defined by field and pupil coordinates)
int __stdcall numpyGetTrace(int nrays, double hx[], double hy[], double px[], double py[],
   double intensity[], int wave_num[], int mode, int surf, int error[], int vigcode[], 
   double pos[][3], double dir[][3], double normal[][3], unsigned int timeout)
{
    int i;

    // allocate memory for list of structures expected by ZEMAX
    DDERAYDATA* RD = calloc(nrays+1,sizeof(*RD));
    
    // set parameters for raytrace (in 0'th element)
    // see Zemax manual for meaning of the fields (do not infer from field names!)
    RD[0].opd=0;         // set type 0 (GetTrace)
    RD[0].wave=mode;     // 0 for real rays, 1 for paraxial rays
    RD[0].error=nrays;
    RD[0].vigcode=0;
    RD[0].want_opd=surf; // surface to trace to, -1 for image, or any valid surface number

    // initialize ray-structure with initial sampling
    for (i=0; i<nrays; i++) {
      RD[i+1].x = hx[i];
      RD[i+1].y = hy[i];
      RD[i+1].z = px[i];
      RD[i+1].l = py[i];
      RD[i+1].intensity = intensity[i];
      RD[i+1].wave = wave_num[i];
      //RD[i+1].error= 0;     // already initialized to 0
      //RD[i+1].vigcode=0;
      //RD[i+1].want_opd=0;
    }

    // arrayTrace
    if(arrayTrace(RD,timeout)==0) {

    // was successful, fill return values
    for (i=0; i<nrays; i++) {
      pos[i][0] = RD[i+1].x;
      pos[i][1] = RD[i+1].y;
      pos[i][2] = RD[i+1].z;
      dir[i][0] = RD[i+1].l;
      dir[i][1] = RD[i+1].m;
      dir[i][2] = RD[i+1].n;
      normal[i][0]=RD[i+1].Exr;
      normal[i][1]=RD[i+1].Eyr;
      normal[i][2]=RD[i+1].Ezr;
      intensity[i]=RD[i+1].intensity;
      error[i]    =RD[i+1].error;
      vigcode[i]  =RD[i+1].vigcode;    
    }
    } // end-if array trace suceeded
    free(RD);
    return RETVAL;
}


// ----------------------------------------------------------------------------------
// Mode 1: similar to GetTraceDirect (rays defined by position and direction cosines on any starting surface)
int __stdcall numpyGetTraceDirect(int nrays, double startpos[][3], double startdir[][3], 
   double intensity[], int wave_num[], int mode, int startsurf, int lastsurf, int error[], 
   int vigcode[], double pos[][3], double dir[][3], double normal[][3], unsigned int timeout)
{
    int i;

    // allocate memory for list of structures expected by ZEMAX
    DDERAYDATA* RD = calloc(nrays+1,sizeof(*RD));

    // set parameters for raytrace (in 0'th element)
    // see Zemax manual for meaning of the fields (do not infer from field names!)
    RD[0].opd=1;         // set type 1 (GetTrace)
    RD[0].wave=mode;     // 0 for real rays, 1 for paraxial rays
    RD[0].error=nrays;
    RD[0].vigcode=startsurf; // the surface on which the coordinates start
    RD[0].want_opd=lastsurf; // surface to trace to, -1 for image, or any valid surface number

    // initialize ray-structure with initial sampling
    for (i=0; i<nrays; i++) {
      RD[i+1].x = startpos[i][0];
      RD[i+1].y = startpos[i][1];
      RD[i+1].z = startpos[i][2];
      RD[i+1].l = startdir[i][0];
      RD[i+1].m = startdir[i][1];
      RD[i+1].n = startdir[i][2];
      RD[i+1].intensity = intensity[i];
      RD[i+1].wave = wave_num[i];
      //RD[i+1].error= 0;           // already initialized to 0
      //RD[i+1].vigcode=0;
    }

    // arrayTrace
    if(arrayTrace(RD,timeout)==0) {

    // was successful, fill return values
    for (i=0; i<nrays; i++) {
      pos[i][0] = RD[i+1].x;
      pos[i][1] = RD[i+1].y;
      pos[i][2] = RD[i+1].z;
      dir[i][0] = RD[i+1].l;
      dir[i][1] = RD[i+1].m;
      dir[i][2] = RD[i+1].n;
      normal[i][0]=RD[i+1].Exr;
      normal[i][1]=RD[i+1].Eyr;
      normal[i][2]=RD[i+1].Ezr;
      intensity[i]=RD[i+1].intensity;
      error[i]    =RD[i+1].error;
      vigcode[i]  =RD[i+1].vigcode;    
    }
    } // end-if array trace suceeded
    free(RD);
    return RETVAL;
}


// ----------------------------------------------------------------------------------
// Calculate OPD for all rays indicated by normalized field and pupil coordinates
// as well as a list of wavenumbers. The return values are multidimensional arrays
// which can be indexed as opd[wave][field][pupil]
int __stdcall numpyOpticalPathDifference(int nField, double hx[], double hy[], 
   int nPupil, double px[], double py[], int nWave, int wave_num[], 
   int error[], int vigcode[], double opd[], double pos[][3], 
   double dir[][3], double intensity[], unsigned int timeout)
{
    int i,iField,iPupil,iWave;
    int nrays = nWave*nField*nPupil;

    // allocate memory for list of structures expected by ZEMAX
    DDERAYDATA* RD = calloc(nrays+1,sizeof(*RD));
    
    // set parameters for raytrace (in 0'th element)
    // see Zemax manual for meaning of the fields (do not infer from field names!)
    RD[0].opd=0;         // set type 0 (GetTrace)
    RD[0].wave=0;        // real ray trace, will be overwritten anyway by Zemax
    RD[0].error=nrays;
    RD[0].vigcode=0;
    RD[0].want_opd=-1;   // trace to image surface

    // initialize ray-structure with initial sampling
    i=1;
    for (iWave=0; iWave<nWave; iWave++) {
      for (iField=0; iField<nField; iField++) {
        for (iPupil=0; iPupil<nPupil; iPupil++) {
            RD[i].x = hx[iField];
            RD[i].y = hy[iField];
            RD[i].z = px[iPupil];
            RD[i].l = py[iPupil];
            RD[i].intensity = 1.0;
            RD[i].wave = wave_num[iWave];
            RD[i].want_opd=1;
            // request Zemax to trace chief ray as reference for OPD
            if (iPupil==0) RD[i].want_opd=-1;
            //RD[i+1].error= 0;     // already initialized to 0
            //RD[i+1].vigcode=0;
            i++;
        }
      }
    }

    // arrayTrace
    if(arrayTrace(RD,timeout)==0) {

    // was successful, fill return values
    for (i=0; i<nrays; i++) {
      pos[i][0] = RD[i+1].x;
      pos[i][1] = RD[i+1].y;
      pos[i][2] = RD[i+1].z;
      dir[i][0] = RD[i+1].l;
      dir[i][1] = RD[i+1].m;
      dir[i][2] = RD[i+1].n;
      // normal is not calculated if want_opd<>0 !
      opd[i]    =RD[i+1].opd;
      intensity[i]=RD[i+1].intensity;
      error[i]  =RD[i+1].error;
      vigcode[i]=RD[i+1].vigcode;    
    }
    } // end-if array trace suceeded
    free(RD);
    return RETVAL;
}


/* ----------------------------------------------------------------------------------
  DDE Communication (from Zemax examples)
   ----------------------------------------------------------------------------------
 */

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
    ATOM          aApp, aTop, aItem;
    DDEACK        DdeAck;
    DDEDATA      *pDdeData;
    GLOBALHANDLE  hDdeData;
    WORD          wStatus;
    UINT_PTR      uiLow, uiHi;

    switch (iMsg)
    {
    case WM_CREATE:
        hwndServer = 0;
        return 0;

    case WM_USER_INITIATE:
        /* find ZEMAX */

        // code added September 1, 2006 to identify which ZEMAX is calling us, in case more than 1 copy of ZEMAX is running
        // this currently is only supported by user defined operands.
        // aApp = GlobalAddAtom ("ZEMAX");
        if (1)
        {
            char szSub[500];
            GetString(szCommandLine, 5, szSub);
            if (strcmp(szSub, "ZEMAX1") == 0 || strcmp(szSub, "ZEMAX2") == 0)
            {
                if (strcmp(szSub, "ZEMAX1") == 0)
                {
                    aApp = GlobalAddAtom("ZEMAX1");
                    ZEMAX_INSTANCE = 1;
                }
                if (strcmp(szSub, "ZEMAX2") == 0)
                {
                    aApp = GlobalAddAtom("ZEMAX2");
                    ZEMAX_INSTANCE = 2;
                }
            }
            else
            {
                aApp = GlobalAddAtom("ZEMAX");
                ZEMAX_INSTANCE = 0;
            }
        }

        aTop = GlobalAddAtom("RayData");

        //SendMessage(HWND_BROADCAST, WM_DDE_INITIATE, (WPARAM)hwnd, MAKELONG(aApp, aTop));    
        SendMessage(HWND_BROADCAST, WM_DDE_INITIATE, (WPARAM)hwnd, PackDDElParam(WM_DDE_INITIATE, aApp, aTop));

        /* delete the atoms */
        GlobalDeleteAtom(aApp);
        GlobalDeleteAtom(aTop);

        /* If no response, terminate */
        if (hwndServer == NULL)
        {
            printf("\nCannot communicate with ZEMAX! Zemax may not be running!\n"); 
            //MessageBox(hwnd, "Cannot communicate with ZEMAX!", "ERROR", MB_ICONEXCLAMATION | MB_OK);
            DestroyWindow(hwnd);
            RETVAL = -999;
            return 0;
        }

        hwndClient = hwnd;

        rayTraceFunction();

        /* terminate the DDE connection */
        PostMessage(hwndServer, WM_DDE_TERMINATE, (WPARAM)hwnd, 0L);
        hwndServer = NULL;

        /* now TERMINATE! */
        DestroyWindow(hwnd);
        return 0;

    case WM_DDE_DATA:
        /* here comes the data! */
        // wParam -- sending window handle
        // lParam -- DDEDATA memory handle & item atom
        UnpackDDElParam(WM_DDE_DATA, lParam, &uiLow, &uiHi);
        FreeDDElParam(WM_DDE_DATA, lParam);
        hDdeData = (GLOBALHANDLE)uiLow;
        pDdeData = (DDEDATA *)GlobalLock(hDdeData);
        aItem = (ATOM)uiHi;

        // Initialize DdeAck structure
        DdeAck.bAppReturnCode = 0;
        DdeAck.reserved = 0;
        DdeAck.fBusy = FALSE;
        DdeAck.fAck = FALSE;

        // Check for matching format, put the data in the buffer
        if (pDdeData->cfFormat == CF_TEXT)
        {
            /* get the data back into RD */
            if (rdpGRD) 
                memcpy(rdpGRD, (DDERAYDATA *)pDdeData->Value, (ngNumRays + 1)*sizeof(DDERAYDATA));
            else 
                strcpy(szGlobalBuffer, (char *)pDdeData->Value);
        }

        GotData = 1;
        GlobalDeleteAtom(aItem);

        // Acknowledge if necessary
        if (pDdeData->fAckReq == TRUE)
        {
            wStatus = *((WORD *)&DdeAck);
            if (!PostMessage((HWND)wParam, WM_DDE_ACK, (WPARAM)hwnd, PackDDElParam(WM_DDE_ACK, wStatus, aItem)))
            {
                if (hDdeData)
                {
                    GlobalUnlock(hDdeData);
                    GlobalFree(hDdeData);
                }
                return 0;
            }
        }

        // Clean up
        GlobalUnlock(hDdeData);
        if (pDdeData->fRelease == TRUE || DdeAck.fAck == FALSE) GlobalFree(hDdeData);
        return 0;

    case WM_DDE_ACK:
        /* we are receiving an acknowledgement */
        /* the only one we care about is in response to the WM_DDE_INITIATE; otherwise free just the memory */
        if (hwndServer == NULL)
        {
            uiLow = (UINT_PTR)NULL;
            uiHi = (UINT_PTR)NULL;
            //UnpackDDElParam(WM_DDE_ACK, lParam, &uiLow, &uiHi);    /* Was causing memory error right after the SendMessage() call from case WM_USER_INITIATE:*/
            //FreeDDElParam(WM_DDE_ACK, lParam);                     /* Was causing memory error right after the SendMessage() call from case WM_USER_INITIATE:*/
            uiLow = (UINT_PTR)(((UINT)lParam) & 0xffff);
            uiHi = (UINT_PTR)((((UINT)lParam) >> 16) & 0xffff);
            hwndServer = (HWND)wParam;
            if (uiLow) GlobalDeleteAtom((ATOM)uiLow);
            if (uiHi) GlobalDeleteAtom((ATOM)uiHi);
        }
        else
        {
            HWND dummy;
            uiLow = (UINT_PTR)NULL;
            uiHi = (UINT_PTR)NULL;
            //UnpackDDElParam(WM_DDE_ACK, lParam, &uiLow, &uiHi);
            //FreeDDElParam(WM_DDE_ACK, lParam); 
            uiLow = (UINT_PTR)(((UINT)lParam) & 0xffff);
            uiHi = (UINT_PTR)((((UINT)lParam) >> 16) & 0xffff);
            dummy = (HWND)wParam;
            if (uiLow) GlobalDeleteAtom((ATOM)uiLow);
            if (uiHi) GlobalDeleteAtom((ATOM)uiHi);
        }
        return 0;

    case WM_DDE_TERMINATE:
        PostMessage(hwndServer, WM_DDE_TERMINATE, (WPARAM)hwnd, 0L);
        hwndServer = NULL;
        return 0;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        BeginPaint(hwnd, &ps);
        EndPaint(hwnd, &ps);
    }
    return 0;

    case WM_CLOSE:
        PostMessage(hwndServer, WM_DDE_TERMINATE, (WPARAM)hwnd, 0L);
        break;             // for default processing

    case WM_DESTROY:
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
    dwTime = GetTickCount();

    sleep_count = 0;

    while (!GotData)
    {
        while (PeekMessage(&msg, hwnd, WM_DDE_FIRST, WM_DDE_LAST, PM_REMOVE))
        {
            DispatchMessage(&msg);
        }
        /* Give the server a chance to respond: wait for 100ms */
        Sleep(100);
        sleep_count++;
        if (sleep_count > 10)   /* check every second if timeout is reached */
        {
            if (GetTickCount() - dwTime > DdeTimeout)
            { 
                printf("\nTimeout reached!\n");      /* will be visible in stdout*/
                RETVAL = -998;                     /* indicate to python calling function */
                return; 
            }  
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
    while (szBuffer[i] && (k <= n))
    {
        szTest[j] = szBuffer[i];

        if (szBuffer[i] == '"')
        {
            i++;
            j++;
            szTest[j] = szBuffer[i];

            /* we have a double quote; keep reading until EOF or another double quote */
            while (szBuffer[i] != '"' && szBuffer[i])
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

    numbytes = (1 + numrays)*sizeof(DDERAYDATA);
    hPokeData = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, (LONG) sizeof(DDEPOKE) + numbytes);
    lpPokeData = (DDEPOKE *)GlobalLock(hPokeData);
    lpPokeData->fRelease = TRUE;
    lpPokeData->cfFormat = CF_TEXT;
    memcpy(lpPokeData->Value, RD, numbytes);

    /* clear the buffers */
    szGlobalBuffer[0] = '\0';
    szBuffer[0] = '\0';

    aItem = GlobalAddAtom("RayArrayData");
    GlobalUnlock(hPokeData);

    GotData = 0;

    if (!PostMessage(hwndServer, WM_DDE_POKE, (WPARAM)hwndClient, PackDDElParam(WM_DDE_POKE, (UINT)hPokeData, aItem)))
    {
        //MessageBox(hwndClient, "Cannot communicate with ZEMAX!", "Hello?", MB_ICONEXCLAMATION | MB_OK);
        printf("\nCannot communicate with ZEMAX! Cannot trace rays!\n");
        GlobalDeleteAtom(aItem);
        GlobalFree(hPokeData);
        RETVAL = -999;
        return -1;
    }
    GlobalDeleteAtom(aItem);
    WaitForData(hwndClient);
    strcpy(szBuffer, szGlobalBuffer);

    /* clear the pointer */
    rdpGRD = NULL;

    if (GotData) 
        return 0;
    else 
        return -1;
}
