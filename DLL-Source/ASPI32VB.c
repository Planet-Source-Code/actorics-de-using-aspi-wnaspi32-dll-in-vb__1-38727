/*
***********************************************************************
'C file: ASPIshim.c
'
'Wrapper DLL for the Adaptec WNASPI32.DLL. 
'The adaptec dll is declared as _decl calling convention, 
'however visual basic uses the _stdcall format.
'
'(c) Jon F. Zahornacky - 2002
' E-mail: jonzeke@yahoo.com
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation; either
'version 2.1 of the License, or (at your option) any later version.
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'	
***********************************************************************
*/

#define STRICT
#include <windows.h>
#include "wnaspi32.h"

// this dll's and WNASPI's instance
HINSTANCE hInst, hInstAspi;

DWORD (*pfnGetASPI32SupportInfo)	(void);
DWORD (*pfnSendASPI32Command)		(LPSRB);
DWORD (*pfnGetASPI32Buffer)			(PASPI32BUFF);
DWORD (*pfnFreeASPI32Buffer)		(PASPI32BUFF);
BOOL  (*pfnTranslateASPI32Address)	(PDWORD, PDWORD);


// Main Entry Point For DLL
BOOL WINAPI DllMain (HINSTANCE hInstA, DWORD dwReason, LPVOID lpvReserved)
	{
	switch (dwReason)
		{
		case DLL_PROCESS_ATTACH:
			// record this dll's instance
			hInst = hInstA;
			
			hInstAspi = LoadLibrary("WNASPI32");
			if (!hInstAspi)
				{
				hInstA = 0;
				exit(-1);
				}
    
			// Store address for ASPI calls
			(FARPROC)pfnGetASPI32SupportInfo   = GetProcAddress(hInstAspi, "GetASPI32SupportInfo");
			(FARPROC)pfnSendASPI32Command	   = GetProcAddress(hInstAspi, "SendASPI32Command");
			(FARPROC)pfnGetASPI32Buffer		   = GetProcAddress(hInstAspi, "GetASPI32Buffer");
			(FARPROC)pfnFreeASPI32Buffer	   = GetProcAddress(hInstAspi, "FreeASPI32Buffer");
			(FARPROC)pfnTranslateASPI32Address = GetProcAddress(hInstAspi, "TranslateASPI32Adress");
			break;
   
		case DLL_THREAD_ATTACH:
			// Thread created
			break;

		case DLL_THREAD_DETACH:
			// Thread exiting (cleanly)
			break;

		case DLL_PROCESS_DETACH:
			// dll is being	free'd
			FreeLibrary (hInstAspi);
			hInst = 0;
			break;

		} return TRUE;
	}


// The Translation Code...
DWORD __stdcall GetASPI32SupportInfoEx(void)
	{
	//this will return the Support Info
	return pfnGetASPI32SupportInfo();
	}

DWORD __stdcall SendASPI32CommandEx(LPSRB psrb)
	{
	//This will return the from Visual Basic sended command
	return pfnSendASPI32Command(psrb);
	}

BOOL  __stdcall GetASPI32BufferEx(PASPI32BUFF pab)
	{
	//this will return the Buffer
	return pfnGetASPI32Buffer(pab);
	}

BOOL __stdcall FreeASPI32BufferEx(PASPI32BUFF pab)
	{
	//This will free the Buffer
	return pfnFreeASPI32Buffer(pab);
	}

BOOL __stdcall TranslateASPI32AddressEx( PDWORD pdwPath, PDWORD pdwDEVNODE)
	{
	//This will translate a address (CD-ROM, HDD, FDD...)
	 return pfnTranslateASPI32Address(pdwPath, pdwDEVNODE);
	}
