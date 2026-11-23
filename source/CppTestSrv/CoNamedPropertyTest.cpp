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

// Build a VARIANT holding a SAFEARRAY(VARIANT) of (key, value) pairs.
// Each element is itself a two-element SAFEARRAY(VARIANT) [key, value], so
// comtypes unpacks the result into a tuple of (key, value) tuples.
static HRESULT BuildPairsVariant(const std::map<LONG, LONG>& values,
                                 VARIANT* result)
{
	VariantInit(result) ;

	SAFEARRAYBOUND outerBound = { static_cast<ULONG>(values.size()), 0 } ;
	SAFEARRAY* outer = SafeArrayCreate(VT_VARIANT, 1, &outerBound) ;
	if (outer == NULL)
	{
		return E_OUTOFMEMORY ;
	}

	LONG index = 0 ;
	for (std::map<LONG, LONG>::const_iterator it = values.begin() ;
	     it != values.end() ; ++it, ++index)
	{
		SAFEARRAYBOUND innerBound = { 2, 0 } ;
		SAFEARRAY* inner = SafeArrayCreate(VT_VARIANT, 1, &innerBound) ;
		if (inner == NULL)
		{
			SafeArrayDestroy(outer) ;
			return E_OUTOFMEMORY ;
		}

		VARIANT key ;
		SetLongVariant(&key, it->first) ;
		VARIANT val ;
		SetLongVariant(&val, it->second) ;
		LONG keyIndex = 0 ;
		LONG valueIndex = 1 ;
		SafeArrayPutElement(inner, &keyIndex, &key) ;
		SafeArrayPutElement(inner, &valueIndex, &val) ;
		VariantClear(&key) ;
		VariantClear(&val) ;

		VARIANT element ;
		VariantInit(&element) ;
		element.vt = VT_ARRAY | VT_VARIANT ;
		element.parray = inner ;
		SafeArrayPutElement(outer, &index, &element) ;
		// SafeArrayPutElement copies the element, so release our copy.
		VariantClear(&element) ;
	}

	result->vt = VT_ARRAY | VT_VARIANT ;
	result->parray = outer ;
	return S_OK ;
}

// Replace target with the (key, value) pairs described by value.
// value must be a SAFEARRAY(VARIANT) of two-element SAFEARRAY(VARIANT)s, or
// VT_NULL/VT_EMPTY to clear the map.
static HRESULT ReplaceMapFromVariant(const VARIANT& value,
                                     std::map<LONG, LONG>& target)
{
	if (value.vt == VT_NULL || value.vt == VT_EMPTY)
	{
		target.clear() ;
		return S_OK ;
	}

	if ((value.vt & VT_ARRAY) == 0 || (value.vt & ~VT_ARRAY) != VT_VARIANT ||
	    value.parray == NULL)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	SAFEARRAY* outer = value.parray ;
	if (SafeArrayGetDim(outer) != 1)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	LONG lower = 0 ;
	LONG upper = -1 ;
	SafeArrayGetLBound(outer, 1, &lower) ;
	SafeArrayGetUBound(outer, 1, &upper) ;

	std::map<LONG, LONG> parsed ;
	for (LONG i = lower ; i <= upper ; ++i)
	{
		VARIANT element ;
		VariantInit(&element) ;
		HRESULT hr = SafeArrayGetElement(outer, &i, &element) ;
		if (FAILED(hr))
		{
			VariantClear(&element) ;
			return hr ;
		}

		if ((element.vt & VT_ARRAY) == 0 ||
		    (element.vt & ~VT_ARRAY) != VT_VARIANT ||
		    element.parray == NULL)
		{
			VariantClear(&element) ;
			return DISP_E_TYPEMISMATCH ;
		}

		SAFEARRAY* inner = element.parray ;
		if (SafeArrayGetDim(inner) != 1)
		{
			VariantClear(&element) ;
			return DISP_E_TYPEMISMATCH ;
		}

		LONG innerLower = 0 ;
		LONG innerUpper = -1 ;
		SafeArrayGetLBound(inner, 1, &innerLower) ;
		SafeArrayGetUBound(inner, 1, &innerUpper) ;
		if (innerUpper - innerLower + 1 != 2)
		{
			VariantClear(&element) ;
			return DISP_E_TYPEMISMATCH ;
		}

		VARIANT key ;
		VariantInit(&key) ;
		VARIANT val ;
		VariantInit(&val) ;
		LONG keyIndex = innerLower ;
		LONG valueIndex = innerLower + 1 ;
		SafeArrayGetElement(inner, &keyIndex, &key) ;
		SafeArrayGetElement(inner, &valueIndex, &val) ;

		if (key.vt != VT_I4 || val.vt != VT_I4)
		{
			VariantClear(&key) ;
			VariantClear(&val) ;
			VariantClear(&element) ;
			return DISP_E_TYPEMISMATCH ;
		}

		parsed[key.lVal] = val.lVal ;
		VariantClear(&key) ;
		VariantClear(&val) ;
		VariantClear(&element) ;
	}

	target.swap(parsed) ;
	return S_OK ;
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

// Build a VARIANT holding an empty SAFEARRAY(VARIANT).
static HRESULT BuildEmptyVariantArray(VARIANT* result)
{
	VariantInit(result) ;

	SAFEARRAYBOUND bound = { 0, 0 } ;
	SAFEARRAY* array = SafeArrayCreate(VT_VARIANT, 1, &bound) ;
	if (array == NULL)
	{
		return E_OUTOFMEMORY ;
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

// Read-only parameterized property.  It reads the same map as Item.
HRESULT __stdcall CNamedPropertyTest::get_ReadOnlyItem(LONG index, VARIANT* value)
{
	trace("get_ReadOnlyItem called") ;
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

// Write-only parameterized property.  It writes the same map as Item.
HRESULT __stdcall CNamedPropertyTest::put_WriteOnlyItem(LONG index, LONG value)
{
	trace("put_WriteOnlyItem called") ;
	m_itemValues[index] = value ;
	return S_OK ;
}

// Parameterized property whose only argument is optional.
// With no index (VT_ERROR), the getter returns the whole map as a sequence of
// (key, value) pairs.  With an index, it behaves like Item.
HRESULT __stdcall CNamedPropertyTest::get_OptionalItem(VARIANT index, VARIANT* value)
{
	trace("get_OptionalItem called") ;
	if (value == NULL)
	{
		return E_POINTER ;
	}

	if (index.vt == VT_ERROR)
	{
		return BuildPairsVariant(m_optionalValues, value) ;
	}

	std::map<LONG, LONG>::const_iterator it = m_optionalValues.find(index.lVal) ;
	if (it != m_optionalValues.end())
	{
		SetLongVariant(value, it->second) ;
	}
	else
	{
		SetBadIndexVariant(value) ;
	}
	return S_OK ;
}

// Setter for the optional-index property.
// With no index (VT_ERROR), the whole map is replaced from a sequence of
// (key, value) pairs.  With an index, a single entry is stored.
HRESULT __stdcall CNamedPropertyTest::put_OptionalItem(VARIANT index, VARIANT value)
{
	trace("put_OptionalItem called") ;

	if (index.vt == VT_ERROR)
	{
		return ReplaceMapFromVariant(value, m_optionalValues) ;
	}

	if (value.vt != VT_I4)
	{
		return DISP_E_TYPEMISMATCH ;
	}
	m_optionalValues[index.lVal] = value.lVal ;
	return S_OK ;
}

// Parameterized property that mixes a default-valued argument with an
// optional argument.  `start` has the default value 1.  When `end` is
// omitted (VT_ERROR), the getter reads from `start` to the last assigned
// index; when nothing is assigned in that range it returns an empty array.
// When `end` is given, it behaves like Range.
HRESULT __stdcall CNamedPropertyTest::get_OptionalRange(LONG start, VARIANT end,
	                                                      VARIANT* value)
{
	trace("get_OptionalRange called") ;
	if (value == NULL)
	{
		return E_POINTER ;
	}

	LONG endIndex ;
	if (end.vt == VT_ERROR)
	{
		if (m_optionalRangeValues.empty())
		{
			return BuildEmptyVariantArray(value) ;
		}
		endIndex = m_optionalRangeValues.rbegin()->first ;
		if (endIndex < start)
		{
			return BuildEmptyVariantArray(value) ;
		}
	}
	else
	{
		endIndex = end.lVal ;
	}

	return BuildRangeVariant(m_optionalRangeValues, start, endIndex, value) ;
}

// Setter for the default/optional-index property.  When `end` is omitted
// (VT_ERROR), the values are written from `start` to the end of the supplied
// sequence.  When `end` is given, it behaves like Range.
HRESULT __stdcall CNamedPropertyTest::put_OptionalRange(LONG start, VARIANT end,
	                                                      VARIANT value)
{
	trace("put_OptionalRange called") ;

	LONG endIndex ;
	if (end.vt == VT_ERROR)
	{
		if ((value.vt & VT_ARRAY) == 0 || (value.vt & ~VT_ARRAY) != VT_VARIANT ||
		    value.parray == NULL)
		{
			return DISP_E_TYPEMISMATCH ;
		}

		LONG lower = 0 ;
		LONG upper = -1 ;
		SafeArrayGetLBound(value.parray, 1, &lower) ;
		SafeArrayGetUBound(value.parray, 1, &upper) ;
		LONG count = upper - lower + 1 ;
		if (count <= 0)
		{
			return DISP_E_BADPARAMCOUNT ;
		}
		endIndex = start + count - 1 ;
	}
	else
	{
		endIndex = end.lVal ;
	}

	return AssignRangeFromVariant(value, start, endIndex, m_optionalRangeValues) ;
}
