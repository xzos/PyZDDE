#include <windows.h>
#include <dde.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <math.h>

#define DLL_EXPORT __declspec(dllexport)

#define WM_USER_INITIATE (WM_USER + 1)
#define DDE_TIMEOUT 50000
#pragma warning ( disable : 4996 ) // functions like strcpy are now deprecated for security reasons; this disables the warning

#ifdef __cplusplus
extern "C" {
#endif

typedef struct
{
    double x, y, z, l, m, n, opd, intensity;
    double Exr, Exi, Eyr, Eyi, Ezr, Ezi;
    int wave, error, vigcode, want_opd;
}DDERAYDATA;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void WaitForData(HWND hwnd);
char *GetString(char *szBuffer, int n, char *szSubString);
int  PostArrayTraceMessage(char *szBuffer, DDERAYDATA *RD);
void rayTraceFunction();
DLL_EXPORT int __stdcall arrayTrace(DDERAYDATA * pRAD, unsigned int timeout);

#ifdef __cplusplus
}
#endif
