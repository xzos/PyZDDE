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
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "gdi32.lib")

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
DLL_EXPORT int __stdcall numpyGetTrace(int nrays, double field[][2], double pupil[][2], 
  double intensity[], int wave_num[], int mode, int surf, int want_opd, 
  int error[], int vigcode[], double pos[][3], double dir[][3], double normal[][3], 
  double opd[], unsigned int timeout);

#ifdef __cplusplus
}
#endif
