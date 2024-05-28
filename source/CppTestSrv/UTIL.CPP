/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
//
// util.cpp - Common utilities for printing out messages
//
//
#include <objbase.h>
#include <cstdio>    //sprintf
#include <cstdlib>
#include <cassert>
// #include <tchar.h>

#include "util.h"

// We are building a local server.
	// Listbox window handle
	extern HWND g_hWndListBox ;

	static inline void output(const char* sz)
	{
		size_t newsize = strlen(sz) + 1;
		wchar_t* wcstring = new wchar_t[newsize];
		size_t convertedChars = 0;
		mbstowcs_s(&convertedChars, wcstring, newsize, sz, _TRUNCATE);
		::SendMessage(g_hWndListBox, LB_ADDSTRING, 0, (LPARAM)wcstring) ;
		delete []wcstring;
	}

//
// Utilities
//
namespace Util
{

//
// Print out a message with a label.
//
void Trace(const char* szLabel, const char* szText, HRESULT hr)
{
	char buf[256] ;
	sprintf(buf, "%s: \t%s", szLabel, szText) ;
	output(buf) ;

	if (FAILED(hr))
	{
		ErrorMessage(hr) ;
	}
}

//
// Print out the COM/OLE error string for an HRESULT.
//
void ErrorMessage(HRESULT hr)
{
	LPTSTR pMsgBuf = NULL;
 
	::FormatMessage( 
		FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
		(LPTSTR)&pMsgBuf,
		0,
		NULL 
	) ;

	char buf[256] ;
	int iLength = wcslen(pMsgBuf)+1 ;
	char* psz = new char[iLength] ;
	wcstombs(psz, pMsgBuf, iLength) ;
	sprintf(buf, "Error (%x): %s", hr, psz) ;
	output(buf) ;
		
	// Free the buffer.
	LocalFree(pMsgBuf) ;
}

} ; // End Namespace Util


//
// Overloaded ostream insertion operator
// Converts from wchar_t to char
//
std::ostream& operator<< ( std::ostream& os, const wchar_t* wsz )
{
	// Length of incoming string
	int iLength = wcslen(wsz)+1 ;

	// Allocate buffer for converted string.
	char* psz = new char[iLength] ;

	// Convert from wchar_t to char.
	wcstombs(psz, wsz, iLength) ;

	// Send it out.
	os << psz ;

	// cleanup
	delete [] psz ;
	return os ;
}
