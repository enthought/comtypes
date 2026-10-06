/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// CoComtypesDispSafearrayParamTest.cpp - Component
//
#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoComtypesDispSafearrayParamTest.h"

// We need to put this declaration here because we explicitely expose a dispinterface
// in parallel to the dual interface but dispinterfaces don't appear in the
// MIDL-generated header file.
EXTERN_C const IID DIID_IDispSafearrayParamTest;

static inline void trace(const char* msg)
	{ Util::Trace("CoComtypesDispSafearrayParamTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoComtypesDispSafearrayParamTest", msg, hr) ;}

///////////////////////////////////////////////////////////
//
// Interface IDualSafearrayParamTest - Implementation
//

HRESULT __stdcall CB::InitArray(SAFEARRAY* *pptest_array)
{
	int i ;
	double *pdata = NULL ;
	HRESULT hr ;
	std::ostringstream sout ;

	// Display the contents of the received SAFEARRAY.
	hr = SafeArrayAccessData(*pptest_array, reinterpret_cast<void**>(&pdata)) ;
	if (FAILED(hr)) 
	{
		return E_FAIL ;
	}
	sout << "Received SAFEARRAY contains:" << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	for (i = 0; i < (*pptest_array)->rgsabound[0].cElements; i++)
	{
		sout << "\n\t\t" << "Element# " << i << ": " << pdata[i] << std::ends ;
		trace(sout.str().c_str()) ;
		sout.str("") ;
	}

	// Modify the SAFEARRAY.
	sout << "Modifying SAFEARRAY contents." << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	for (i = 0; i < (*pptest_array)->rgsabound[0].cElements; i++)
	{
		pdata[i] = (double)(i) ;
	}

	// Display the contents of the modifyied SAFEARRAY.
	sout << "Modifyied SAFEARRAY now contains:" << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	for (i = 0; i < (*pptest_array)->rgsabound[0].cElements; i++)
	{
		sout << "\n\t\t" << "Element# " << i << ": " << pdata[i] << std::ends ;
		trace(sout.str().c_str()) ;
		sout.str("") ;
	}
	SafeArrayUnaccessData(*pptest_array) ;

	return S_OK ;
}

HRESULT __stdcall CB::VerifyArray(SAFEARRAY* ptest_array,
									VARIANT_BOOL* result)
{
	long i ;
	double *pdata = NULL ;
	HRESULT hr ;
	std::ostringstream sout ;

	// Display and verify the contents of the received SAFEARRAY.
	*result = VARIANT_TRUE ;
	hr = SafeArrayAccessData(ptest_array, reinterpret_cast<void**>(&pdata)) ;
	if (FAILED(hr)) 
	{
		return E_FAIL ;
	}
	sout << "Received SAFEARRAY contains:  " << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	for (i = 0; i < ptest_array->rgsabound[0].cElements; i++)
	{
		if (pdata[i] != (double)(i))
		{
			*result = VARIANT_FALSE ;
		}
		sout << "\n\t\t" << "Element# " << i << ": " << pdata[i] << std::ends ;
		trace(sout.str().c_str()) ;
		sout.str("") ;
	}

	// Modify the SAFEARRAY.
	sout << "Modifying SAFEARRAY contents." << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	for (i = 0; i < ptest_array->rgsabound[0].cElements; i++)
	{
		pdata[i] = (double)(0.0) ;
	}

	// Display the contents of the modifyied SAFEARRAY.
	sout << "Modified SAFEARRAY now contains:" << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	for (i = 0; i < ptest_array->rgsabound[0].cElements; i++)
	{
		sout << "\n\t\t" << "Element# " << i << ": " << pdata[i] << std::ends ;
		trace(sout.str().c_str()) ;
		sout.str("") ;
	}
	SafeArrayUnaccessData(ptest_array) ;

	return S_OK ;
}


// Return native arrays with fixed values, independent of Python/NumPy input.
// Two-dimensional arrays have two rows and are filled in column-major order.
template <class T>
static HRESULT CreateNumericArray(VARTYPE vartype, short dimensions, long count,
                                 const T (&values)[3], SAFEARRAY** result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = NULL ;
	if ((dimensions != 1 && dimensions != 2) || count <= 0 ||
	    (dimensions == 2 && count % 2 != 0))
	{
		return E_INVALIDARG ;
	}
	SAFEARRAYBOUND bounds[2] = {{static_cast<ULONG>(count), -2}, {1, 5}} ;
	if (dimensions == 2)
	{
		bounds[0].cElements = 2 ;
		bounds[1].cElements = count / 2 ;
	}
	SAFEARRAY* array = SafeArrayCreate(vartype, dimensions, bounds) ;
	if (array == NULL)
	{
		return E_OUTOFMEMORY ;
	}
	T* data = NULL ;
	HRESULT hr = SafeArrayAccessData(array, reinterpret_cast<void**>(&data)) ;
	if (FAILED(hr))
	{
		SafeArrayDestroy(array) ;
		return hr ;
	}
	for (long i = 0; i < count; ++i)
	{
		data[i] = values[i % 3] ;
	}
	hr = SafeArrayUnaccessData(array) ;
	if (FAILED(hr))
	{
		SafeArrayDestroy(array) ;
		return hr ;
	}
	*result = array ;
	return S_OK ;
}

HRESULT __stdcall CB::GetUint8Array(short dimensions, long count, SAFEARRAY** result)
{
	const BYTE values[3] = {0, 42, 255} ;
	return CreateNumericArray(VT_UI1, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetInt16Array(short dimensions, long count, SAFEARRAY** result)
{
	const SHORT values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I2, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetUint16Array(short dimensions, long count, SAFEARRAY** result)
{
	const USHORT values[3] = {0, 42, 65535} ;
	return CreateNumericArray(VT_UI2, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetInt32Array(short dimensions, long count, SAFEARRAY** result)
{
	const LONG values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I4, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetUint32Array(short dimensions, long count, SAFEARRAY** result)
{
	const ULONG values[3] = {0, 42, 4294967295UL} ;
	return CreateNumericArray(VT_UI4, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetInt64Array(short dimensions, long count, SAFEARRAY** result)
{
	const LONGLONG values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I8, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetUint64Array(short dimensions, long count, SAFEARRAY** result)
{
	const ULONGLONG values[3] = {0, 42, 4294967297ULL} ;
	return CreateNumericArray(VT_UI8, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetFloat32Array(short dimensions, long count, SAFEARRAY** result)
{
	const FLOAT values[3] = {1.25f, -2.5f, 3.75f} ;
	return CreateNumericArray(VT_R4, dimensions, count, values, result) ;
}

HRESULT __stdcall CB::GetFloat64Array(short dimensions, long count, SAFEARRAY** result)
{
	const DOUBLE values[3] = {1.25, -2.5, 3.75} ;
	return CreateNumericArray(VT_R8, dimensions, count, values, result) ;
}

//
// Constructor
//
CB::CB(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter), 
  m_pITypeInfo(NULL)
{
	// Empty
}

//
// Destructor
//
CB::~CB()
{
	if (m_pITypeInfo != NULL)
	{
		m_pITypeInfo->Release() ;
	}

	trace("Destroy self.") ;
}

//
// NondelegatingQueryInterface implementation
//
HRESULT __stdcall CB::NondelegatingQueryInterface(const IID& iid,
                                                  void** ppv)
{ 	
	if (iid == IID_IDualSafearrayParamTest)
	{
		return FinishQI(static_cast<IDualSafearrayParamTest*>(this), ppv) ;
	}
	else if (iid == IID_INumericSafearrayTest)
	{
		return FinishQI(static_cast<INumericSafearrayTest*>(this), ppv) ;
	}
	else 	if (iid == DIID_IDispSafearrayParamTest)
	{
		trace("Queried for IDispSafearrayParamTest.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else 	if (iid == IID_IDispatch)
	{
		trace("Queried for IDispatch.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else
	{
		return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
	}
}

//
// Load and register the type library.
//
HRESULT CB::Init()
{
	HRESULT hr ;

	// Load TypeInfo on demand if we haven't already loaded it.
	if (m_pITypeInfo == NULL)
	{
		ITypeLib* pITypeLib = NULL ;
		hr = ::LoadRegTypeLib(LIBID_ComtypesCppTestSrvLib, 
		                      1, 0, // Major/Minor version numbers
		                      0x00, 
		                      &pITypeLib) ;
		if (FAILED(hr)) 
		{
			trace("LoadRegTypeLib Failed.", hr) ;
			return hr ;   
		}

		// Get type information for the interface of the object.
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualSafearrayParamTest,
		                                  &m_pITypeInfo) ;
		pITypeLib->Release() ;
		if (FAILED(hr))  
		{ 
			trace("GetTypeInfoOfGuid failed.", hr) ;
			return hr ;
		}   
	}
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Creation function used by CFactory
//
HRESULT CB::CreateInstance(IUnknown* pUnknownOuter,
                           CUnknown** ppNewComponent ) 
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CB(pUnknownOuter) ;
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// IDispatch implementation
//
HRESULT __stdcall CB::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CB::GetTypeInfo(
	UINT iTypeInfo,
	LCID,          // This object does not support localization.
	ITypeInfo** ppITypeInfo)
{    
	*ppITypeInfo = NULL ;

	if(iTypeInfo != 0)
	{
		trace("GetTypeInfo call failed -- bad iTypeInfo index.") ;
		return DISP_E_BADINDEX ; 
	}

	trace("GetTypeInfo call succeeded.") ;

	// Call AddRef and return the pointer.
	m_pITypeInfo->AddRef() ; 
	*ppITypeInfo = m_pITypeInfo ;
	return S_OK ;
}

HRESULT __stdcall CB::GetIDsOfNames(  
	const IID& iid,
	OLECHAR** arrayNames,
	UINT countNames,
	LCID,          // Localization is not supported.
	DISPID* arrayDispIDs)
{
	if (iid != IID_NULL)
	{
		trace("GetIDsOfNames call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}

	trace("GetIDsOfNames call succeeded.") ;
	HRESULT hr = m_pITypeInfo->GetIDsOfNames(arrayNames,
	                                         countNames,
	                                         arrayDispIDs) ;
	return hr ;
}

HRESULT __stdcall CB::Invoke(   
      DISPID dispidMember,
      const IID& iid,
      LCID,          // Localization is not supported.
      WORD wFlags,
      DISPPARAMS* pDispParams,
      VARIANT* pvarResult,
      EXCEPINFO* pExcepInfo,
      UINT* pArgErr)
{        
	if (iid != IID_NULL)
	{
		trace("Invoke call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}

	::SetErrorInfo(0, NULL) ;

	trace("Invoke call succeeded.") ;
	HRESULT hr = m_pITypeInfo->Invoke(
		static_cast<IDispatch*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ; 
	return hr ;
}
