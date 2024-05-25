/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// Cmpnt.cpp - Component
//
#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "Cmpnt.h"

// We need to put this declaration here because we explicitely expose a dispinterface
// in parallel to the dual interface but dispinterfaces don't appear in the
// MIDL-generated header file.
EXTERN_C const IID DIID_IComtypesTest;

static inline void trace(const char* msg)
	{ Util::Trace("Component", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("Component", msg, hr) ;}

///////////////////////////////////////////////////////////
//
// Interface IDualComtypesTest - Implementation
//

HRESULT __stdcall CA::InitRecord(T_TEST_RECORD* test_record)
{
	// Display the received T_TEST_RECORD structure.
	if (test_record->question == NULL){
		test_record->question = ::SysAllocString(L"") ;
	}
	std::ostringstream sout ;
	sout << "Received T_TEST_RECORD structure contains:  " << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	sout << "\n\t\t" << "question: " << test_record->question << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	sout << "\n\t\t" << "answer: " << test_record->answer << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	sout << "\n\t\t" << "needs_clarification: " << test_record->needs_clarification << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;

	if (! ::SysReAllocString(&(test_record->question), L"The meaning of life, the universe and everything?"))
	{
		return E_OUTOFMEMORY ;
	}
	test_record->answer = 42 ;
	test_record->needs_clarification = VARIANT_TRUE ;

	return S_OK ;
}

HRESULT __stdcall CA::VerifyRecord(T_TEST_RECORD* test_record,
																   VARIANT_BOOL* result)
{
	// Display the received T_TEST_RECORD structure.
	if (test_record->question == NULL){
		test_record->question = ::SysAllocString(L"") ;
	}
	std::ostringstream sout ;
	sout << "Received T_TEST_RECORD structure contains:  " << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	sout << "\n\t\t" << "question: " << test_record->question << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	sout << "\n\t\t" << "answer: " << test_record->answer << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;
	sout << "\n\t\t" << "needs_clarification: " << test_record->needs_clarification << std::ends ;
	trace(sout.str().c_str()) ;
	sout.str("") ;

	// Check if we received an initialization record.
	if (_wcsicmp(test_record->question, L"The meaning of life, the universe and everything?") == 0
		&& test_record->answer == 42
		&& test_record->needs_clarification == VARIANT_TRUE){
		*result = VARIANT_TRUE ;
	}
	else {
		*result =  VARIANT_FALSE ;
	}

	// Modify the received record.
	// This modification should not change the record on the client side
	// because it is just an [in] parameter and not passed with VT_BYREF.
	test_record->answer == 12 ;

	return S_OK ;
}

//
// Constructor
//
CA::CA(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter), 
  m_pITypeInfo(NULL)
{
	// Empty
}

//
// Destructor
//
CA::~CA()
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
HRESULT __stdcall CA::NondelegatingQueryInterface(const IID& iid,
                                                  void** ppv)
{ 	
	if (iid == IID_IDualComtypesTest)
	{
		return FinishQI(static_cast<IDualComtypesTest*>(this), ppv) ;
	}
	else 	if (iid == DIID_IComtypesTest)
	{
		trace("Queried for IComtypesTest.") ;
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
HRESULT CA::Init()
{
	HRESULT hr ;

	// Load TypeInfo on demand if we haven't already loaded it.
	if (m_pITypeInfo == NULL)
	{
		ITypeLib* pITypeLib = NULL ;
		hr = ::LoadRegTypeLib(LIBID_ComtypesTestLib, 
		                      1, 0, // Major/Minor version numbers
		                      0x00, 
		                      &pITypeLib) ;
		if (FAILED(hr)) 
		{
			trace("LoadRegTypeLib Failed.", hr) ;
			return hr ;   
		}

		// Get type information for the interface of the object.
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualComtypesTest,
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
HRESULT CA::CreateInstance(IUnknown* pUnknownOuter,
                           CUnknown** ppNewComponent ) 
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CA(pUnknownOuter) ;
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// IDispatch implementation
//
HRESULT __stdcall CA::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CA::GetTypeInfo(
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

HRESULT __stdcall CA::GetIDsOfNames(  
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

HRESULT __stdcall CA::Invoke(   
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
