/*
 * CoComtypesDispPutRefTest.cpp - Component for dispinterface propputref tests
 *
 * The members of IDualPutRefTest are copied into the dispinterface
 * IDispPutRefTest, so comtypes builds them through its dispinterface member
 * generator.  Item exercises a code path that parameterized properties do
 * not reach:
 *
 *   * Item is a propputref property.  comtypes invokes its setter with
 *     DISPATCH_PROPERTYPUTREF.
 */

#include <objbase.h>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoComtypesDispPutRefTest.h"

// We need to put this declaration here because we explicitly expose a
// dispinterface in parallel to the dual interface, but dispinterfaces don't
// appear in the MIDL-generated header file.
EXTERN_C const IID DIID_IDispPutRefTest;

static inline void trace(const char* msg)
	{ Util::Trace("CoComtypesDispPutRefTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoComtypesDispPutRefTest", msg, hr) ;}

///////////////////////////////////////////////////////////
//
// Constructor / Destructor
//
CComtypesDispPutRefTest::CComtypesDispPutRefTest(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter),
  m_pITypeInfo(NULL)
{
	// Empty
}

CComtypesDispPutRefTest::~CComtypesDispPutRefTest()
{
	for (std::map<LONG, IUnknown*>::iterator it = m_refs.begin() ;
	     it != m_refs.end() ; ++it)
	{
		if (it->second != NULL)
		{
			it->second->Release() ;
		}
	}

	if (m_pITypeInfo != NULL)
	{
		m_pITypeInfo->Release() ;
	}

	trace("Destroy self.") ;
}

///////////////////////////////////////////////////////////
//
// Creation function used by CFactory
//
HRESULT CComtypesDispPutRefTest::CreateInstance(IUnknown* pUnknownOuter,
                                                CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CComtypesDispPutRefTest(pUnknownOuter) ;
	return S_OK ;
}

//
// NondelegatingQueryInterface implementation
//
HRESULT __stdcall CComtypesDispPutRefTest::NondelegatingQueryInterface(
	const IID& iid, void** ppv)
{
	if (iid == IID_IDualPutRefTest)
	{
		return FinishQI(static_cast<IDualPutRefTest*>(this), ppv) ;
	}
	else if (iid == DIID_IDispPutRefTest)
	{
		trace("Queried for IDispPutRefTest.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else if (iid == IID_IDispatch)
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
HRESULT CComtypesDispPutRefTest::Init()
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
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualPutRefTest,
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
// IDispatch Implementation
//
HRESULT __stdcall CComtypesDispPutRefTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CComtypesDispPutRefTest::GetTypeInfo(UINT iTypeInfo, LCID,
                                                       ITypeInfo** ppITypeInfo)
{
	*ppITypeInfo = NULL ;

	if (iTypeInfo != 0)
	{
		trace("GetTypeInfo call failed -- bad iTypeInfo index.") ;
		return DISP_E_BADINDEX ;
	}

	trace("GetTypeInfo call succeeded.") ;
	m_pITypeInfo->AddRef() ;
	*ppITypeInfo = m_pITypeInfo ;
	return S_OK ;
}

HRESULT __stdcall CComtypesDispPutRefTest::GetIDsOfNames(const IID& iid,
                                                         OLECHAR** arrayNames,
                                                         UINT countNames, LCID,
                                                         DISPID* arrayDispIDs)
{
	if (iid != IID_NULL)
	{
		trace("GetIDsOfNames call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}

	trace("GetIDsOfNames call succeeded.") ;
	HRESULT hr = m_pITypeInfo->GetIDsOfNames(arrayNames, countNames, arrayDispIDs) ;
	return hr ;
}

HRESULT __stdcall CComtypesDispPutRefTest::Invoke(DISPID dispidMember,
                                                  const IID& iid, LCID,
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
		static_cast<IDualPutRefTest*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ;
	return hr ;
}

///////////////////////////////////////////////////////////
//
// Interface IDualPutRefTest - Implementation
//

// Reads back the pointer stored by putref_Item.
HRESULT __stdcall CComtypesDispPutRefTest::get_Item(LONG index, IUnknown** value)
{
	trace("get_Item called") ;
	if (value == NULL)
	{
		return E_POINTER ;
	}

	*value = NULL ;

	std::map<LONG, IUnknown*>::const_iterator it = m_refs.find(index) ;
	if (it == m_refs.end() || it->second == NULL)
	{
		return S_OK ;
	}

	*value = it->second ;
	it->second->AddRef() ;
	return S_OK ;
}

// Stores the object reference carried by value.  comtypes reaches this
// through DISPATCH_PROPERTYPUTREF.  The value is typed as IUnknown, so the
// object reference arrives as a COM pointer rather than a VARIANT.
HRESULT __stdcall CComtypesDispPutRefTest::putref_Item(LONG index,
                                                       IUnknown* value)
{
	trace("putref_Item called") ;

	std::map<LONG, IUnknown*>::iterator it = m_refs.find(index) ;
	if (it != m_refs.end())
	{
		if (it->second != NULL)
		{
			it->second->Release() ;
		}
		m_refs.erase(it) ;
	}

	if (value != NULL)
	{
		value->AddRef() ;
		m_refs[index] = value ;
	}

	return S_OK ;
}
