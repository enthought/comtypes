/*
 * CoNamedPropertyTest.cpp - Component for parameterized (named) property tests
 *
 * The properties on IDualNamedPropertyTest take arguments, so comtypes
 * exposes them through its `named_property` descriptor.  The native logic is
 * intentionally trivial: the read/write properties simply store and return
 * the value that was assigned.
 *
 * The getters return a VARIANT.  Reading an index that has not been assigned
 * reports DISP_E_BADINDEX through the VARIANT's scode instead of inventing a
 * value, so a caller can tell "not set" apart from a stored value.
 */

#include <objbase.h>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoNamedPropertyTest.h"

// We need to put this declaration here because we explicitly expose a
// dispinterface in parallel to the dual interface, but dispinterfaces don't
// appear in the MIDL-generated header file.
EXTERN_C const IID DIID_IDispNamedPropertyTest;

static inline void trace(const char* msg)
	{ Util::Trace("CoNamedPropertyTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoNamedPropertyTest", msg, hr) ;}

// Set *value to a VT_I4 VARIANT holding v.
static void SetLongVariant(VARIANT* value, LONG v)
{
	VariantInit(value) ;
	value->vt = VT_I4 ;
	value->lVal = v ;
}

// Set *value to a VT_ERROR VARIANT whose scode is DISP_E_BADINDEX.
static void SetBadIndexVariant(VARIANT* value)
{
	VariantInit(value) ;
	value->vt = VT_ERROR ;
	value->scode = DISP_E_BADINDEX ;
}

// Build a VARIANT holding a SAFEARRAY(VARIANT) with one element per index in
// [start, end].  An index that has no stored value is filled with a VT_NULL
// VARIANT, so a caller can tell "not set" apart from a stored value.
static HRESULT BuildRangeVariant(const std::map<LONG, LONG>& values,
	                                LONG start, LONG end,
	                                VARIANT* result)
{
	VariantInit(result) ;

	if (end < start)
	{
		return E_INVALIDARG ;
	}

	LONG count = end - start + 1 ;
	SAFEARRAYBOUND bound = { static_cast<ULONG>(count), 0 } ;
	SAFEARRAY* array = SafeArrayCreate(VT_VARIANT, 1, &bound) ;
	if (array == NULL)
	{
		return E_OUTOFMEMORY ;
	}

	for (LONG i = 0 ; i < count ; ++i)
	{
		VARIANT element ;
		VariantInit(&element) ;

		std::map<LONG, LONG>::const_iterator it = values.find(start + i) ;
		if (it != values.end())
		{
			SetLongVariant(&element, it->second) ;
		}
		else
		{
			element.vt = VT_NULL ;
		}

		SafeArrayPutElement(array, &i, &element) ;
		VariantClear(&element) ;
	}

	result->vt = VT_ARRAY | VT_VARIANT ;
	result->parray = array ;
	return S_OK ;
}

// Assign the elements of value to the indices [start, end].
// value must be a SAFEARRAY(VARIANT) whose element count equals
// end - start + 1; each element must be a VT_I4.
static HRESULT AssignRangeFromVariant(const VARIANT& value,
	                                     LONG start, LONG end,
	                                     std::map<LONG, LONG>& target)
{
	if (end < start)
	{
		return E_INVALIDARG ;
	}

	if ((value.vt & VT_ARRAY) == 0 || (value.vt & ~VT_ARRAY) != VT_VARIANT ||
	    value.parray == NULL)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	SAFEARRAY* array = value.parray ;
	if (SafeArrayGetDim(array) != 1)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	LONG lower = 0 ;
	LONG upper = -1 ;
	SafeArrayGetLBound(array, 1, &lower) ;
	SafeArrayGetUBound(array, 1, &upper) ;

	LONG count = end - start + 1 ;
	if (upper - lower + 1 != count)
	{
		return DISP_E_BADPARAMCOUNT ;
	}

	for (LONG i = 0 ; i < count ; ++i)
	{
		VARIANT element ;
		VariantInit(&element) ;
		LONG index = lower + i ;
		HRESULT hr = SafeArrayGetElement(array, &index, &element) ;
		if (FAILED(hr))
		{
			VariantClear(&element) ;
			return hr ;
		}

		if (element.vt != VT_I4)
		{
			VariantClear(&element) ;
			return DISP_E_TYPEMISMATCH ;
		}

		target[start + i] = element.lVal ;
		VariantClear(&element) ;
	}

	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Constructor / Destructor
//
CNamedPropertyTest::CNamedPropertyTest(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter),
  m_pITypeInfo(NULL)
{
	// Empty
}

CNamedPropertyTest::~CNamedPropertyTest()
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
HRESULT CNamedPropertyTest::CreateInstance(IUnknown* pUnknownOuter,
                                           CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CNamedPropertyTest(pUnknownOuter) ;
	return S_OK ;
}

//
// NondelegatingQueryInterface implementation
//
HRESULT __stdcall CNamedPropertyTest::NondelegatingQueryInterface(const IID& iid,
                                                                  void** ppv)
{
	if (iid == IID_IDualNamedPropertyTest)
	{
		return FinishQI(static_cast<IDualNamedPropertyTest*>(this), ppv) ;
	}
	else if (iid == DIID_IDispNamedPropertyTest)
	{
		trace("Queried for IDispNamedPropertyTest.") ;
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
HRESULT CNamedPropertyTest::Init()
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
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualNamedPropertyTest,
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
HRESULT __stdcall CNamedPropertyTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CNamedPropertyTest::GetTypeInfo(UINT iTypeInfo, LCID,
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

HRESULT __stdcall CNamedPropertyTest::GetIDsOfNames(const IID& iid,
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

HRESULT __stdcall CNamedPropertyTest::Invoke(DISPID dispidMember, const IID& iid,
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
		static_cast<IDualNamedPropertyTest*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ;
	return hr ;
}

///////////////////////////////////////////////////////////
//
// Interface IDualNamedPropertyTest - Implementation
//

// Read/write parameterized property with a single index.
// Returns the value previously stored with put_Item, or a VARIANT whose scode
// is DISP_E_BADINDEX when nothing has been stored for the index.
HRESULT __stdcall CNamedPropertyTest::get_Item(LONG index, VARIANT* value)
{
	trace("get_Item called") ;
	if (value == NULL)
	{
		return E_POINTER ;
	}

	std::map<LONG, LONG>::const_iterator it = m_itemValues.find(index) ;
	if (it != m_itemValues.end())
	{
		SetLongVariant(value, it->second) ;
	}
	else
	{
		SetBadIndexVariant(value) ;
	}
	return S_OK ;
}

HRESULT __stdcall CNamedPropertyTest::put_Item(LONG index, LONG value)
{
	trace("put_Item called") ;
	m_itemValues[index] = value ;
	return S_OK ;
}

// Read/write parameterized property with two indices that denote a range.
// Reading Range[start, end] returns a SAFEARRAY(VARIANT) with one element per
// index in [start, end]; an index that has not been assigned is filled with a
// VT_NULL VARIANT.
HRESULT __stdcall CNamedPropertyTest::get_Range(LONG start, LONG end,
                                                VARIANT* value)
{
	trace("get_Range called") ;
	if (value == NULL)
	{
		return E_POINTER ;
	}

	return BuildRangeVariant(m_rangeValues, start, end, value) ;
}

// Writing Range[start, end] = values assigns values[i] to index start + i.
// The number of values must equal end - start + 1.
HRESULT __stdcall CNamedPropertyTest::put_Range(LONG start, LONG end,
                                                VARIANT value)
{
	trace("put_Range called") ;
	return AssignRangeFromVariant(value, start, end, m_rangeValues) ;
}
