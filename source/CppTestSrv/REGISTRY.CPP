/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// Registry.cpp
//

#include <objbase.h>
#include <cassert>
#include <cstdio>    // sprintf
#include <cstdlib>   // splitpath

#include "Iface.h"
#include "Registry.h"

////////////////////////////////////////////////////////
//
// Internal helper functions prototypes
//

// Set the given key and its value.
BOOL setKeyAndValue(LPCWSTR pszPath,
                    LPCWSTR szSubkey,
                    LPCWSTR szValue) ;

// Convert a GUID into a char string.
void GUIDtochar(const GUID& guid, 
                 LPWSTR szGUID,
                 int length) ;

// Determine if a particular subkey exists.
BOOL SubkeyExists(LPCWSTR pszPath,
                  LPCWSTR szSubkey) ;

// Delete szKeyChild and all of its descendents.
LONG recursiveDeleteKey(HKEY hKeyParent, LPCWSTR szKeyChild) ;

//
// Type library name
//
const wchar_t szTypeLibName[] = L"server.tlb" ;

////////////////////////////////////////////////////////
//
// Constants
//

// Size of a GUID as a string
const int GUID_STRING_SIZE = 39 ;

/////////////////////////////////////////////////////////
//
// Public function implementation
//

//
// Register the component in the registry.
//
HRESULT RegisterServer(HMODULE hModule,        // DLL module handle
                       const CLSID& clsid,     // Class ID
                       LPCWSTR szFriendlyName, // Friendly Name
                       LPCWSTR szVerIndProgID, // Programmatic
                       LPCWSTR szProgID,       // IDs
					   const GUID& libid)      // Library ID
{
	// Get server location.
	wchar_t szModule[512] ;
	DWORD dwResult =
		::GetModuleFileName(hModule, 
		                    szModule,
		                    sizeof(szModule)/sizeof(wchar_t)) ;
	assert(dwResult != 0) ;

	// Convert the CLSID into a char.
	wchar_t szCLSID[GUID_STRING_SIZE] ;
	GUIDtochar(clsid, szCLSID, GUID_STRING_SIZE) ;

	// Build the key CLSID\\{...}
	wchar_t szKey[64] ;
	wcscpy(szKey, L"CLSID\\") ;
	wcscat(szKey, szCLSID) ;
  
	// Add the CLSID to the registry.
	setKeyAndValue(szKey, NULL, szFriendlyName) ;

	// Add the server filename subkey under the CLSID key.
	setKeyAndValue(szKey, L"LocalServer32", szModule) ;

	// Add the ProgID subkey under the CLSID key.
	setKeyAndValue(szKey, L"ProgID", szProgID) ;

	// Add the version-independent ProgID subkey under CLSID key.
	setKeyAndValue(szKey, L"VersionIndependentProgID",
	               szVerIndProgID) ;

	// Add the Type Library ID subkey under the CLSID key.
	wchar_t szLIBID[GUID_STRING_SIZE] ;
	GUIDtochar(libid, szLIBID, GUID_STRING_SIZE) ;
	setKeyAndValue(szKey, L"TypeLib", szLIBID) ;

	
	// Add the version-independent ProgID subkey under HKEY_CLASSES_ROOT.
	setKeyAndValue(szVerIndProgID, NULL, szFriendlyName) ; 
	setKeyAndValue(szVerIndProgID, L"CLSID", szCLSID) ;
	setKeyAndValue(szVerIndProgID, L"CurVer", szProgID) ;

	// Add the versioned ProgID subkey under HKEY_CLASSES_ROOT.
	setKeyAndValue(szProgID, NULL, szFriendlyName) ; 
	setKeyAndValue(szProgID, L"CLSID", szCLSID) ;

	// Register the Type Library.
	ITypeLib* pITypeLib = NULL ;
	// Try to load it from the path.
	wchar_t szDrive[_MAX_DRIVE];
	wchar_t szDir[_MAX_DIR];
	// Split the fullname to get the pathname.
	_wsplitpath(szModule, szDrive, szDir, NULL, NULL) ;

	// Append name of registry.
	wchar_t szTypeLibFullName[_MAX_PATH];
	_swprintf(szTypeLibFullName,
			L"%s%s%s",
			szDrive,
			szDir,
			szTypeLibName) ;

	// convert to wide char
	// wchar_t wszTypeLibFullName[_MAX_PATH] ;
	// mbstowcs(wszTypeLibFullName, szTypeLibFullName, _MAX_PATH) ;

	// if LoadTypeLib succeeds, it will have registered
	// the type library for us.
	// for the next time.  
	HRESULT hr = ::LoadTypeLib(szTypeLibFullName,
									&pITypeLib) ;
	assert(hr == S_OK) ;

	// Ensure that the type library is registered.
	hr = RegisterTypeLib(pITypeLib, szTypeLibFullName, NULL) ;
	assert(hr == S_OK) ;

	return S_OK ;
}

//
// Remove the component from the registry.
//
LONG UnregisterServer(const CLSID& clsid,         // Class ID
                      LPCWSTR szVerIndProgID, // Programmatic
                      LPCWSTR szProgID)       //   IDs
{
	// Convert the CLSID into a char.
	wchar_t szCLSID[GUID_STRING_SIZE] ;
	GUIDtochar(clsid, szCLSID, GUID_STRING_SIZE) ;

	// Build the key CLSID\\{...}
	wchar_t szKey[80] ;
	wcscpy(szKey, L"CLSID\\") ;
	wcscat(szKey, szCLSID) ;

	// Check for a another server for this component.
	if (SubkeyExists(szKey, L"InprocServer32"))
	{
		// Delete only the path for this server.
		wcscat(szKey, L"\\LocalServer32") ;
		LONG lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szKey) ;
		assert(lResult == ERROR_SUCCESS) ;
	}
	else
	{
		// Delete all related keys.
		// Delete the CLSID Key - CLSID\{...}
		LONG lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szKey) ;
		assert((lResult == ERROR_SUCCESS) ||
		       (lResult == ERROR_FILE_NOT_FOUND)) ; // Subkey may not exist.

		// Delete the version-independent ProgID Key.
		lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szVerIndProgID) ;
		assert((lResult == ERROR_SUCCESS) ||
		       (lResult == ERROR_FILE_NOT_FOUND)) ; // Subkey may not exist.

		// Delete the ProgID key.
		lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szProgID) ;
		assert((lResult == ERROR_SUCCESS) ||
		       (lResult == ERROR_FILE_NOT_FOUND)) ; // Subkey may not exist.
	}

	// Unregister the Type Library.
	HRESULT hr = UnRegisterTypeLib(LIBID_ComtypesCppTestSrvLib, 
									1, 0, // Major/Minor version numbers
									0x00, 
									SYS_WIN64) ;
	assert(hr == S_OK) ;

	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Internal helper functions
//

// Convert a GUID to a wchar_t string.
void GUIDtochar(const GUID& guid,
                 LPWSTR szGUID,
                 int length)
{
	assert(length >= GUID_STRING_SIZE) ;
	// Get wide string version.
	LPOLESTR wszGUID = NULL ;
	HRESULT hr = StringFromCLSID(guid, &wszGUID) ;
	assert(SUCCEEDED(hr)) ;

	// Covert from wide characters to non-wide.
	// wcstombs(szGUID, wszGUID, length) ;
	// Copy the retrieved string.
	wcscpy_s(szGUID, length, wszGUID) ;


	// Free memory.
	CoTaskMemFree(wszGUID) ;
}

//
// Delete a key and all of its descendents.
//
LONG recursiveDeleteKey(HKEY hKeyParent,           // Parent of key to delete
                        LPCWSTR lpszKeyChild)      // Key to delete
{
	// Open the child.
	HKEY hKeyChild ;
	LONG lRes = RegOpenKeyEx(hKeyParent, lpszKeyChild, 0,
	                         KEY_ALL_ACCESS, &hKeyChild) ;
	if (lRes != ERROR_SUCCESS)
	{
		return lRes ;
	}

	// Enumerate all of the decendents of this child.
	FILETIME time ;
	wchar_t szBuffer[256] ;
	DWORD dwSize = 256 ;
	while (RegEnumKeyEx(hKeyChild, 0, szBuffer, &dwSize, NULL,
	                    NULL, NULL, &time) == S_OK)
	{
		// Delete the decendents of this child.
		lRes = recursiveDeleteKey(hKeyChild, szBuffer) ;
		if (lRes != ERROR_SUCCESS)
		{
			// Cleanup before exiting.
			RegCloseKey(hKeyChild) ;
			return lRes;
		}
		dwSize = 256 ;
	}

	// Close the child.
	RegCloseKey(hKeyChild) ;

	// Delete this child.
	return RegDeleteKey(hKeyParent, lpszKeyChild) ;
}

//
// Determine if a particular subkey exists.
//
BOOL SubkeyExists(LPCWSTR pszPath,    // Path of key to check
                  LPCWSTR szSubkey)   // Key to check
{
	HKEY hKey ;
	wchar_t szKeyBuf[80] ;

	// Copy keyname into buffer.
	wcscpy(szKeyBuf, pszPath) ;

	// Add subkey name to buffer.
	if (szSubkey != NULL)
	{
		wcscat(szKeyBuf, L"\\") ;
		wcscat(szKeyBuf, szSubkey ) ;
	}

	// Determine if key exists by trying to open it.
	LONG lResult = ::RegOpenKeyEx(HKEY_CLASSES_ROOT, 
	                              szKeyBuf,
	                              0,
	                              KEY_ALL_ACCESS,
	                              &hKey) ;
	if (lResult == ERROR_SUCCESS)
	{
		RegCloseKey(hKey) ;
		return TRUE ;
	}
	return FALSE ;
}

//
// Create a key and set its value.
//   - This helper function was borrowed and modifed from
//     Kraig Brockschmidt's book Inside OLE.
//
BOOL setKeyAndValue(LPCWSTR szKey,
                    LPCWSTR szSubkey,
                    LPCWSTR szValue)
{
	HKEY hKey;
	wchar_t szKeyBuf[1024] ;

	// Copy keyname into buffer.
	wcscpy(szKeyBuf, szKey) ;

	// Add subkey name to buffer.
	if (szSubkey != NULL)
	{
		wcscat(szKeyBuf, L"\\") ;
		wcscat(szKeyBuf, szSubkey ) ;
	}

	// Create and open key and subkey.
	long lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT ,
	                              szKeyBuf, 
	                              0, NULL, REG_OPTION_NON_VOLATILE,
	                              KEY_ALL_ACCESS, NULL, 
	                              &hKey, NULL) ;
	if (lResult != ERROR_SUCCESS)
	{
		return FALSE ;
	}

	// Set the Value.
	if (szValue != NULL)
	{
		RegSetValueEx(hKey, NULL, 0, REG_SZ, 
		              (BYTE *)szValue, 
		              sizeof(wchar_t) * (wcslen(szValue)+1)) ;
	}

	RegCloseKey(hKey) ;
	return TRUE ;
}
