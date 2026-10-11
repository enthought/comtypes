#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h"
#include "CoStringSafearrayInsideVariantTest.h"

static inline void trace(const char* msg)
	{ Util::Trace("CoStringSafearrayInsideVariantTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoStringSafearrayInsideVariantTest", msg, hr) ;}

static HRESULT Extract1DSafeArray(const VARIANT& var, SAFEARRAY** ppsa, VARTYPE* pvtElem)
{
	SAFEARRAY* psa = NULL ;
	if (var.vt & VT_BYREF)
	{
		if ((var.vt & ~VT_BYREF) == (VT_ARRAY | VT_BSTR) ||
		    (var.vt & ~VT_BYREF) == (VT_ARRAY | VT_VARIANT) ||
		    ((var.vt & ~VT_BYREF) & VT_ARRAY))
		{
			psa = *(var.pparray) ;
		}
		else
		{
			return DISP_E_TYPEMISMATCH ;
		}
	}
	else if (var.vt & VT_ARRAY)
	{
		psa = var.parray ;
	}
	else
	{
		return DISP_E_TYPEMISMATCH ;
	}

	if (psa == NULL)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	if (SafeArrayGetDim(psa) != 1)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	VARTYPE vtElem = 0 ;
	HRESULT hr = SafeArrayGetVartype(psa, &vtElem) ;
	if (FAILED(hr) || vtElem == VT_EMPTY)
	{
		vtElem = var.vt & ~(VT_ARRAY | VT_BYREF) ;
	}

	*ppsa = psa ;
	*pvtElem = vtElem ;
	return S_OK ;
}

static HRESULT CheckBstrArrayValues(SAFEARRAY* psa, VARIANT_BOOL* result)
{
	long lBound = 0 ;
	long uBound = -1 ;
	SafeArrayGetLBound(psa, 1, &lBound) ;
	SafeArrayGetUBound(psa, 1, &uBound) ;
	long count = uBound - lBound + 1 ;

	if (count != 2)
	{
		*result = VARIANT_FALSE ;
		return S_OK ;
	}

	BSTR* pbstrData = NULL ;
	HRESULT hr = SafeArrayAccessData(psa, reinterpret_cast<void**>(&pbstrData)) ;
	if (FAILED(hr))
	{
		return hr ;
	}

	if (pbstrData[0] != NULL && wcscmp(pbstrData[0], L"foo") == 0 &&
	    pbstrData[1] != NULL && wcscmp(pbstrData[1], L"bar") == 0)
	{
		*result = VARIANT_TRUE ;
	}
	else
	{
		*result = VARIANT_FALSE ;
	}

	SafeArrayUnaccessData(psa) ;
	return S_OK ;
}

static HRESULT CheckVariantArrayValues(SAFEARRAY* psa, VARIANT_BOOL* result)
{
	long lBound = 0 ;
	long uBound = -1 ;
	SafeArrayGetLBound(psa, 1, &lBound) ;
	SafeArrayGetUBound(psa, 1, &uBound) ;
	long count = uBound - lBound + 1 ;

	if (count != 2)
	{
		*result = VARIANT_FALSE ;
		return S_OK ;
	}

	VARIANT* pvarData = NULL ;
	HRESULT hr = SafeArrayAccessData(psa, reinterpret_cast<void**>(&pvarData)) ;
	if (FAILED(hr))
	{
		return hr ;
	}

	if (pvarData[0].vt == VT_BSTR && pvarData[0].bstrVal != NULL && wcscmp(pvarData[0].bstrVal, L"foo") == 0 &&
	    pvarData[1].vt == VT_BSTR && pvarData[1].bstrVal != NULL && wcscmp(pvarData[1].bstrVal, L"bar") == 0)
	{
		*result = VARIANT_TRUE ;
	}
	else
	{
		*result = VARIANT_FALSE ;
	}

	SafeArrayUnaccessData(psa) ;
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Interface IStringSafearrayInsideVariantTest - Implementation
//

static HRESULT GetStringArrayLength(SAFEARRAY* arr, VARTYPE elementType,
	ULONG* length)
{
	if (arr == NULL || length == NULL || SafeArrayGetDim(arr) != 1)
	{
		return E_INVALIDARG ;
	}
	VARTYPE actual = 0 ;
	if (FAILED(SafeArrayGetVartype(arr, &actual)) || actual != elementType)
	{
		return DISP_E_TYPEMISMATCH ;
	}
	LONG lo = 0, hi = -1 ; SafeArrayGetLBound(arr, 1, &lo); SafeArrayGetUBound(arr, 1, &hi);
	*length = hi >= lo ? static_cast<ULONG>(hi - lo + 1) : 0;
	return S_OK ;
}

static HRESULT CreateStringArray(VARTYPE elementType, ULONG length,
	SAFEARRAY** result)
{
	SAFEARRAYBOUND bound = {length, 0};
	*result = SafeArrayCreate(elementType, 1, &bound);
	if (*result == NULL && bound.cElements != 0)
	{
		return E_OUTOFMEMORY;
	}
	return S_OK ;
}

static HRESULT PrepareRepeatedStringArray(SAFEARRAY* arr, VARTYPE elementType,
	LONG repeat, SAFEARRAY** result, ULONG* length)
{
	if (result == NULL || repeat < 0)
	{
		return E_INVALIDARG ;
	}
	HRESULT hr = GetStringArrayLength(arr, elementType, length) ;
	if (FAILED(hr))
	{
		return hr ;
	}
	SAFEARRAYBOUND bound = {*length * static_cast<ULONG>(repeat), 0};
	return CreateStringArray(elementType, bound.cElements, result) ;
}

static HRESULT VariantArrayToBstrArrayImpl(const VARIANT& var, SAFEARRAY** result)
{
	SAFEARRAY* arr = NULL; VARTYPE elementType = 0;
	HRESULT hr = Extract1DSafeArray(var, &arr, &elementType);
	if (FAILED(hr) || elementType != VT_VARIANT) return DISP_E_TYPEMISMATCH;
	ULONG length = 0 ;
	hr = GetStringArrayLength(arr, VT_VARIANT, &length) ;
	if (FAILED(hr) || result == NULL)
	{
		return FAILED(hr) ? hr : E_INVALIDARG ;
	}
	hr = CreateStringArray(VT_BSTR, length, result);
	if (FAILED(hr))
	{
		return hr;
	}
	SAFEARRAY* out = *result;
	if (length != 0)
	{
		void* src = NULL; void* dst = NULL; hr = SafeArrayAccessData(arr, &src);
		if (FAILED(hr))
		{
			return hr;
		}
		hr = SafeArrayAccessData(out, &dst);
		if (FAILED(hr))
		{
			SafeArrayUnaccessData(arr); SafeArrayDestroy(out);
			return hr;
		}
		for (ULONG i = 0; i < length; ++i)
		{
			static_cast<BSTR*>(dst)[i] = SysAllocString(static_cast<VARIANT*>(src)[i].bstrVal);
		}
		SafeArrayUnaccessData(out); SafeArrayUnaccessData(arr);
	}
	return S_OK;
}

static HRESULT BstrArrayToVariantArrayImpl(const VARIANT& var, SAFEARRAY** result)
{
	SAFEARRAY* arr = NULL; VARTYPE elementType = 0;
	HRESULT hr = Extract1DSafeArray(var, &arr, &elementType);
	if (FAILED(hr) || elementType != VT_BSTR) return DISP_E_TYPEMISMATCH;
	ULONG length = 0 ;
	hr = GetStringArrayLength(arr, VT_BSTR, &length) ;
	if (FAILED(hr) || result == NULL)
	{
		return FAILED(hr) ? hr : E_INVALIDARG ;
	}
	hr = CreateStringArray(VT_VARIANT, length, result);
	if (FAILED(hr))
	{
		return hr;
	}
	SAFEARRAY* out = *result;
	if (length != 0)
	{
		void* src = NULL; void* dst = NULL; hr = SafeArrayAccessData(arr, &src);
		if (FAILED(hr))
		{
			return hr;
		}
		hr = SafeArrayAccessData(out, &dst);
		if (FAILED(hr))
		{
			SafeArrayUnaccessData(arr); SafeArrayDestroy(out);
			return hr;
		}
		for (ULONG i = 0; i < length; ++i)
		{
			VariantInit(&static_cast<VARIANT*>(dst)[i]);
			static_cast<VARIANT*>(dst)[i].vt = VT_BSTR;
			static_cast<VARIANT*>(dst)[i].bstrVal = SysAllocString(static_cast<BSTR*>(src)[i]);
		}
		SafeArrayUnaccessData(out); SafeArrayUnaccessData(arr);
	}
	return S_OK;
}

static HRESULT RepeatVariantArrayImpl(const VARIANT& var, LONG repeat,
	SAFEARRAY** result)
{
	SAFEARRAY* arr = NULL; VARTYPE elementType = 0;
	HRESULT hr = Extract1DSafeArray(var, &arr, &elementType);
	if (FAILED(hr) || elementType != VT_VARIANT) return DISP_E_TYPEMISMATCH;
	ULONG length = 0 ;
	hr = PrepareRepeatedStringArray(arr, VT_VARIANT, repeat, result, &length) ;
	if (FAILED(hr))
	{
		return hr;
	}
	SAFEARRAY* out = *result;
	SAFEARRAYBOUND bound = {length * static_cast<ULONG>(repeat), 0};
	if (bound.cElements != 0)
	{
		void* src = NULL; void* dst = NULL; hr = SafeArrayAccessData(arr, &src);
		if (FAILED(hr))
		{
			return hr;
		}
		hr = SafeArrayAccessData(out, &dst);
		if (FAILED(hr))
		{
			SafeArrayUnaccessData(arr); SafeArrayDestroy(out);
			return hr;
		}
		for (ULONG i = 0; i < bound.cElements; ++i)
		{
			ULONG j = i % length;
			VariantInit(&static_cast<VARIANT*>(dst)[i]);
			VariantCopy(&static_cast<VARIANT*>(dst)[i], &static_cast<VARIANT*>(src)[j]);
		}
		SafeArrayUnaccessData(out); SafeArrayUnaccessData(arr);
	}
	return S_OK;
}

static HRESULT RepeatBstrArrayImpl(const VARIANT& var, LONG repeat,
	SAFEARRAY** result)
{
	SAFEARRAY* arr = NULL; VARTYPE elementType = 0;
	HRESULT hr = Extract1DSafeArray(var, &arr, &elementType);
	if (FAILED(hr) || elementType != VT_BSTR) return DISP_E_TYPEMISMATCH;
	ULONG length = 0 ;
	hr = PrepareRepeatedStringArray(arr, VT_BSTR, repeat, result, &length) ;
	if (FAILED(hr))
	{
		return hr;
	}
	SAFEARRAY* out = *result;
	SAFEARRAYBOUND bound = {length * static_cast<ULONG>(repeat), 0};
	if (bound.cElements != 0)
	{
		void* src = NULL; void* dst = NULL; hr = SafeArrayAccessData(arr, &src);
		if (FAILED(hr))
		{
			return hr;
		}
		hr = SafeArrayAccessData(out, &dst);
		if (FAILED(hr))
		{
			SafeArrayUnaccessData(arr); SafeArrayDestroy(out);
			return hr;
		}
		for (ULONG i = 0; i < bound.cElements; ++i)
		{
			ULONG j = i % length;
			static_cast<BSTR*>(dst)[i] = SysAllocString(static_cast<BSTR*>(src)[j]);
		}
		SafeArrayUnaccessData(out); SafeArrayUnaccessData(arr);
	}
	return S_OK;
}

static HRESULT WrapStringArrayResult(
	HRESULT hr, SAFEARRAY* array, VARTYPE elementType, VARIANT* result)
{
	if (FAILED(hr))
	{
		if (array != NULL) SafeArrayDestroy(array);
		return hr;
	}
	if (result == NULL)
	{
		if (array != NULL) SafeArrayDestroy(array);
		return E_INVALIDARG;
	}
	VariantInit(result);
	result->vt = VT_ARRAY | elementType;
	result->parray = array;
	return S_OK;
}

HRESULT __stdcall CStringSafearrayInsideVariantTest::VariantArrayToBstrArray(VARIANT a, VARIANT* r)
{
	SAFEARRAY* array = NULL;
	return WrapStringArrayResult(VariantArrayToBstrArrayImpl(a, &array), array, VT_BSTR, r);
}
HRESULT __stdcall CStringSafearrayInsideVariantTest::BstrArrayToVariantArray(VARIANT a, VARIANT* r)
{
	SAFEARRAY* array = NULL;
	return WrapStringArrayResult(BstrArrayToVariantArrayImpl(a, &array), array, VT_VARIANT, r);
}
HRESULT __stdcall CStringSafearrayInsideVariantTest::RepeatVariantArray(VARIANT a, LONG n, VARIANT* r)
{
	SAFEARRAY* array = NULL;
	return WrapStringArrayResult(RepeatVariantArrayImpl(a, n, &array), array, VT_VARIANT, r);
}
HRESULT __stdcall CStringSafearrayInsideVariantTest::RepeatBstrArray(VARIANT a, LONG n, VARIANT* r)
{
	SAFEARRAY* array = NULL;
	return WrapStringArrayResult(RepeatBstrArrayImpl(a, n, &array), array, VT_BSTR, r);
}

///////////////////////////////////////////////////////////
//
// Constructor / Destructor / QI
//

CStringSafearrayInsideVariantTest::CStringSafearrayInsideVariantTest(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter),
  m_pITypeInfo(NULL)
{
}

CStringSafearrayInsideVariantTest::~CStringSafearrayInsideVariantTest()
{
	if (m_pITypeInfo != NULL)
	{
		m_pITypeInfo->Release() ;
	}
	trace("Destroy self.") ;
}

HRESULT __stdcall CStringSafearrayInsideVariantTest::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
	if (iid == IID_IStringSafearrayInsideVariantTest)
	{
		return FinishQI(static_cast<IStringSafearrayInsideVariantTest*>(this), ppv) ;
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

HRESULT CStringSafearrayInsideVariantTest::Init()
{
	HRESULT hr ;
	if (m_pITypeInfo == NULL)
	{
		ITypeLib* pITypeLib = NULL ;
		hr = ::LoadRegTypeLib(LIBID_ComtypesCppTestSrvLib,
		                      1, 0,
		                      0x00,
		                      &pITypeLib) ;
		if (FAILED(hr))
		{
			trace("LoadRegTypeLib Failed.", hr) ;
			return hr ;
		}

		hr = pITypeLib->GetTypeInfoOfGuid(IID_IStringSafearrayInsideVariantTest,
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

HRESULT CStringSafearrayInsideVariantTest::CreateInstance(IUnknown* pUnknownOuter,
                                          CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		return CLASS_E_NOAGGREGATION ;
	}
	*ppNewComponent = new CStringSafearrayInsideVariantTest(pUnknownOuter) ;
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// IDispatch implementation
//

HRESULT __stdcall CStringSafearrayInsideVariantTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CStringSafearrayInsideVariantTest::GetTypeInfo(
	UINT iTypeInfo,
	LCID,
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

HRESULT __stdcall CStringSafearrayInsideVariantTest::GetIDsOfNames(
	const IID& iid,
	OLECHAR** arrayNames,
	UINT countNames,
	LCID,
	DISPID* arrayDispIDs)
{
	if (iid != IID_NULL)
	{
		trace("GetIDsOfNames call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}
	trace("GetIDsOfNames call succeeded.") ;
	return m_pITypeInfo->GetIDsOfNames(arrayNames, countNames, arrayDispIDs) ;
}

HRESULT __stdcall CStringSafearrayInsideVariantTest::Invoke(
	DISPID dispidMember,
	const IID& iid,
	LCID,
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
	return m_pITypeInfo->Invoke(
		static_cast<IDispatch*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ;
}
