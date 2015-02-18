/**************************************************************************/
/* The next few lines are standard "boiler plate" code all UDF's have     */
/**************************************************************************/

#include <windows.h>
#include <dde.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <math.h>

/* boiler plate declarations for functions in zclient we will call */
typedef struct
	{
   double x, y, z, l, m, n, opd, intensity;
	double Exr, Exi, Eyr, Eyi, Ezr, Ezi;
	int wave, error, vigcode, want_opd;
	}DDERAYDATA;

void WaitForData(HWND hwnd);
char *GetString(char *szBuffer, int n, char *szSubString);
void remove_quotes(char *s);
void PostRequestMessage(char *szItem, char *szDataOut);
void PostArrayTraceMessage(char *szBuffer, DDERAYDATA *RD);
void CenterWindow(HWND hwnd);
void UserFunction(char *szCommandLine);
void MakeEmptyWindow(int text, char *szAppName, char *szOptions);
void Get_2_5_10(double cmax, double *cscale);

/* boiler plate global variables used by the client code */
extern HINSTANCE globalhInstance;
extern HWND hwndClient;


/**************************************************/
/* Below is the code specific to this UDF program */
/**************************************************/

// ArrayDemo sample program
// Written by Kenneth Moore March 1999
// Version 1.0

// prints intensity modified by user defined surfaces - June 1999

// This sample program illustrates using ZCLIENT calls to trace large numbers
// of rays with the GetTRaceArray family of DDE calls. It is really easy to do!
// Most of this code is just for defining the rays to trace and printing the results.


/* okay, now the actual code that creates the UDF data for display in a ZEMAX window */
/* this function MUST be called UserFunction and take the command line argument      */

void UserFunction(char *szCommandLine)
{
char szModuleName[300], szOutputFile[300];
FILE *output;
int  i, j, k, show_settings;
static char szBuffer[5000], szSub[256], szAppName[] = "Array Demo";
DDERAYDATA RD[1000];

/* extract the command line arguments */
show_settings = atoi(GetString(szCommandLine, 1, szSub));
/* this tells us where to put the output data */
GetString(szCommandLine, 2, szOutputFile);
remove_quotes(szOutputFile);

//MessageBox(hwndClient, szCommandLine, "szCommandLine:", MB_OK|MB_ICONEXCLAMATION|MB_APPLMODAL);
//MessageBox(hwndClient, szOutputFile, "Output file:", MB_OK|MB_ICONEXCLAMATION|MB_APPLMODAL);

/* first, update the lens so we have the latest data; and then test to make sure the system is valid */
PostRequestMessage("GetUpdate", szBuffer);

if (atoi(GetString(szBuffer, 0, szSub)))
	{
   /* the current lens cannot be traced! */
   /* some features may be able to create data without tracing rays; some can't */
   /* If data cannot be created return "empty" data displays */
   sprintf(szBuffer,"???");
   MakeEmptyWindow(1, szAppName, szBuffer);
   return;
   }

if (show_settings) MessageBox (hwndClient, "This window has no options.", "ZEMAX Client Window", MB_ICONEXCLAMATION | MB_OK | MB_SYSTEMMODAL);

/* Fill RD with data to trace some rays */
RD[0].x = 0.0;
RD[0].y = 0.0;
RD[0].z = 0.0;
RD[0].l = 0.0;
RD[0].m = 0.0;
RD[0].n = 0.0;
RD[0].opd = 0.0; /* mode 0, like GetTrace */
RD[0].intensity = 0.0;
RD[0].wave = 0;
RD[0].error = 0; /* trace a bunch of rays */
RD[0].vigcode = 0;
RD[0].want_opd = -1;

/* define the rays */
k = 0;
for (i = -10; i <= 10; i++)
	{
	for (j = -10; j <= 10; j++)
		{
      k++;
      RD[k].x = 0.0;
      RD[k].y = 0.0;
      RD[k].z = (double) i / 20.0;
      RD[k].l = (double) j / 20.0;
      RD[k].m = 0.0;
      RD[k].n = 0.0;
      RD[k].opd = 0.0;
      RD[k].intensity = 1.0;
      RD[k].wave = 1;
      RD[k].error = 0;
      RD[k].vigcode = 0;
      RD[k].want_opd = 0;
      }
   }
RD[0].error = k; /* trace the k rays */

/* Now go get the data */
PostArrayTraceMessage(szBuffer, RD);
/* Okay, we got the data! There, wasn't that easy! */

/* open a file for output */
if ((output = fopen(szOutputFile, "wt")) == NULL)
	{
   /* can't open the file!! */
   return;
   }

/* this windows function gives us the name of our own executable; we pass this back to ZEMAX */
GetModuleFileName(globalhInstance, szModuleName, 255);

/* ok, make a text listing */
fputs("Listing of Array trace data\n",output);

fputs("     px      py error            xout            yout   trans\n", output);

k = 0;
for (i = -10; i <= 10; i++)
	{
	for (j = -10; j <= 10; j++)
		{
      k++;

      sprintf(szBuffer, "%7.3f %7.3f %5i %15.6E %15.6E %7.4f\n", (double) i / 20.0, (double) j / 20.0, RD[k].error, RD[k].x, RD[k].y, RD[k].intensity);
		fputs(szBuffer, output);
      }
   }
/* close the file! */
fclose(output);

/* create a text window. Note we pass back the filename and module name. */

//MessageBox(hwndClient, szOutputFile, "Output file:", MB_OK|MB_ICONEXCLAMATION|MB_APPLMODAL);

sprintf(szBuffer,"MakeTextWindow,\"%s\",\"%s\",\"%s\"", szOutputFile, szModuleName, szAppName);

//MessageBox(hwndClient, szBuffer, "szBuffer:", MB_OK|MB_ICONEXCLAMATION|MB_APPLMODAL);

PostRequestMessage(szBuffer, szBuffer);  // insr: this statement sends the data back to Zemax for showing the text window 
                                         // along with the previous sprintf(szBuffer, "Mak....) statement
}



