/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

#ifndef __Util_h__
#define __Util_h__

//
// Util.h - Shared utilities
//
#include <strstream>

namespace Util
{
	void Trace(const char* szLabel, const char* szText, HRESULT hr) ;

	void ErrorMessage(HRESULT hr) ;
} ;


//
// Overloaded insertion operator for converting from
// Unicode (wchar_t) to non-Unicode.
//
std::ostream& operator<< ( std::ostream& os, const wchar_t* wsz ) ;

#endif // __Util_h__