/*
 * CoCoClassParamTest.cpp - Component for coclass return type tests
 *
 * The native logic is intentionally trivial: GetItemAsCC returns a new
 * CCoClassParamTestItem object, whose default interface is
 * ICoClassParamTestItem.  This mirrors the Excel type library, where a
 * property such as _Application.Application is declared to return the
 * Application coclass.
 */

#include <objbase.h>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoCoClassParamTest.h"

// We need to put this declaration here because we explicitly expose a
// dispinterface in parallel to the dual interface, but dispinterfaces don't
// appear in the MIDL-generated header file.
EXTERN_C const IID DIID_IDispCoClassParamTest;

static inline void trace(const char* msg)
	{ Util::Trace("CoCoClassParamTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoCoClassParamTest", msg, hr) ;}

///////////////////////////////////////////////////////////
//
// CCoClassParamTestItem - Implementation
//

HRESULT __stdcall CCoClassParamTestItem::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
	if (iid == IID_ICoClassParamTestItem)
	{
		return FinishQI(static_cast<ICoClassParamTestItem*>(this), ppv) ;
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
}

// Return a fixed value so that a caller can tell the returned object apart
// from any other object.
HRESULT __stdcall CCoClassParamTestItem::GetValue(long* result)
{
	trace("GetValue called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = 42 ;
	return S_OK ;
}

// Return this object itself, AddRef'd, typed as the ICoClassParamTestItem
// interface.  AddRef is required for the same reason QueryInterface always
// AddRefs: the caller receives a new reference that it owns.
HRESULT __stdcall CCoClassParamTestItem::get_SelfAsItf(ICoClassParamTestItem** result)
{
	trace("get_SelfAsItf called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = static_cast<ICoClassParamTestItem*>(this) ;
	AddRef() ;
	return S_OK ;
}

// Return this object itself, AddRef'd, typed as the CoCoClassParamTestItem
// coclass.  The only difference from get_SelfAsItf is the IDL typing: a
// coclass return type makes comtypes QueryInterface the result for the
// coclass' default interface.
HRESULT __stdcall CCoClassParamTestItem::get_SelfAsCC(CoCoClassParamTestItem** result)
{
	trace("get_SelfAsCC called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = reinterpret_cast<CoCoClassParamTestItem*>(static_cast<ICoClassParamTestItem*>(this)) ;
	AddRef() ;
	return S_OK ;
}

// Compare two interface pointers for COM identity.  The COM identity rule
// says that QueryInterface(IUnknown) returns the same pointer for the same
// object, so comparing the IUnknown pointers is the correct way to tell
// whether two interface pointers refer to the same object.  This is
// independent of which interface each pointer was typed as.
static HRESULT IsSameComObject(IUnknown* pLeft, IUnknown* pRight,
                               VARIANT_BOOL* result)
{
	IUnknown* pLeftUnknown = NULL ;
	HRESULT hr = pLeft->QueryInterface(IID_IUnknown,
	                                   reinterpret_cast<void**>(&pLeftUnknown)) ;
	if (FAILED(hr))
	{
		return hr ;
	}

	IUnknown* pRightUnknown = NULL ;
	hr = pRight->QueryInterface(IID_IUnknown,
	                            reinterpret_cast<void**>(&pRightUnknown)) ;
	if (FAILED(hr))
	{
		pLeftUnknown->Release() ;
		return hr ;
	}

	*result = (pLeftUnknown == pRightUnknown) ? VARIANT_TRUE : VARIANT_FALSE ;
	pLeftUnknown->Release() ;
	pRightUnknown->Release() ;
	return S_OK ;
}

// Return whether item is the same COM object as this one.  The argument is
// typed as the CoCoClassParamTestItem coclass, so the type library declares
// the parameter as a coclass pointer.  A coclass pointer is the coclass'
// default interface pointer, so it is reinterpreted as
// ICoClassParamTestItem before the identity comparison.
HRESULT __stdcall CCoClassParamTestItem::IsSameItemAsCC(CoCoClassParamTestItem* item,
                                                        VARIANT_BOOL* result)
{
	trace("IsSameItemAsCC called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = VARIANT_FALSE ;
	if (item == NULL)
	{
		return S_OK ;
	}
	return IsSameComObject(reinterpret_cast<ICoClassParamTestItem*>(item),
	                       static_cast<ICoClassParamTestItem*>(this), result) ;
}

// Return whether item is the same COM object as this one.  The argument is
// typed as the ICoClassParamTestItem interface, so the type library declares
// the parameter as an interface pointer.
HRESULT __stdcall CCoClassParamTestItem::IsSameItemAsItf(ICoClassParamTestItem* item,
                                                         VARIANT_BOOL* result)
{
	trace("IsSameItemAsItf called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = VARIANT_FALSE ;
	if (item == NULL)
	{
		return S_OK ;
	}
	return IsSameComObject(item, static_cast<ICoClassParamTestItem*>(this), result) ;
}

///////////////////////////////////////////////////////////
//
// Constructor / Destructor
//
CCoClassParamTest::CCoClassParamTest(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter),
  m_pITypeInfo(NULL)
{
	// Empty
}

CCoClassParamTest::~CCoClassParamTest()
{
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
HRESULT CCoClassParamTest::CreateInstance(IUnknown* pUnknownOuter,
                                          CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CCoClassParamTest(pUnknownOuter) ;
	return S_OK ;
}

//
// NondelegatingQueryInterface implementation
//
HRESULT __stdcall CCoClassParamTest::NondelegatingQueryInterface(const IID& iid,
                                                                 void** ppv)
{
	if (iid == IID_IDualCoClassParamTest)
	{
		return FinishQI(static_cast<IDualCoClassParamTest*>(this), ppv) ;
	}
	else if (iid == DIID_IDispCoClassParamTest)
	{
		trace("Queried for IDispCoClassParamTest.") ;
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
HRESULT CCoClassParamTest::Init()
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
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualCoClassParamTest,
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
HRESULT __stdcall CCoClassParamTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CCoClassParamTest::GetTypeInfo(UINT iTypeInfo, LCID,
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

HRESULT __stdcall CCoClassParamTest::GetIDsOfNames(const IID& iid,
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

HRESULT __stdcall CCoClassParamTest::Invoke(DISPID dispidMember, const IID& iid,
                                            LCID, WORD wFlags,
                                            DISPPARAMS* pDispParams,
                                            VARIANT* pvarResult,
                                            EXCEPINFO* pExcepInfo, UINT* pArgErr)
{
	if (iid != IID_NULL)
	{
		trace("Invoke call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}

	::SetErrorInfo(0, NULL) ;
	trace("Invoke call succeeded.") ;
	HRESULT hr = m_pITypeInfo->Invoke(
		static_cast<IDualCoClassParamTest*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ;
	return hr ;
}

///////////////////////////////////////////////////////////
//
// Interface IDualCoClassParamTest - Implementation
//

// Return a new CCoClassParamTestItem object.  The type library declares the
// return type as the CoCoClassParamTestItem coclass, so the returned pointer
// is the default interface of CoCoClassParamTestItem
// (ICoClassParamTestItem), cast to the incomplete CoCoClassParamTestItem
// type.  The caller owns the returned reference.
HRESULT __stdcall CCoClassParamTest::GetItemAsCC(CoCoClassParamTestItem** result)
{
	trace("GetItemAsCC called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = NULL ;

	CCoClassParamTestItem* pResult = new CCoClassParamTestItem(NULL) ;
	if (pResult == NULL)
	{
		return E_OUTOFMEMORY ;
	}

	*result = reinterpret_cast<CoCoClassParamTestItem*>(static_cast<ICoClassParamTestItem*>(pResult)) ;
	return S_OK ;
}

// Return a new CCoClassParamTestItem object typed as the
// ICoClassParamTestItem interface.  Unlike GetItemAsCC, the type library
// declares the return type as ICoClassParamTestItem, so comtypes generates
// POINTER(ICoClassParamTestItem) for the dispinterface return type.
// POINTER(ICoClassParamTestItem) has __com_interface__, so the dispinterface
// method QueryInterfaces the result for ICoClassParamTestItem.  The caller
// owns the returned reference.
HRESULT __stdcall CCoClassParamTest::GetItemAsItf(ICoClassParamTestItem** result)
{
	trace("GetItemAsItf called") ;
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = NULL ;

	CCoClassParamTestItem* pResult = new CCoClassParamTestItem(NULL) ;
	if (pResult == NULL)
	{
		return E_OUTOFMEMORY ;
	}

	*result = static_cast<ICoClassParamTestItem*>(pResult) ;
	return S_OK ;
}
