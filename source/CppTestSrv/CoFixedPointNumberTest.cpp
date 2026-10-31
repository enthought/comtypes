// CoFixedPointNumberTest.cpp
// Test component for converting strings to VT_DECIMAL and VT_CY VARIANTs.
// ToDecVariant: takes a BSTR string and calls VarDecFromStr to return a VT_DECIMAL VARIANT.
// ToCyVariant: takes a BSTR string and calls VarCyFromStr to return a VT_CY VARIANT.
// InitCyVariantWith3dot14: [in, out] VARIANT*. Discards the incoming content and reflects a VT_CY VARIANT holding 3.14.
//                          If the incoming VARIANT is VT_BYREF | VT_CY, writes 3.14 through the reference instead.
// InitDecVariantWith3dot14: [in, out] VARIANT*. Discards the incoming content and reflects a VT_DECIMAL VARIANT holding 3.14.
//                           If the incoming VARIANT is VT_BYREF | VT_DECIMAL, writes 3.14 through the reference instead.
#include <objbase.h>
#include <oleauto.h>
#include <string.h>
#include "CoFixedPointNumberTest.h"

HRESULT CFixedPointNumberTest::CreateInstance(IUnknown* pUnknownOuter,
                                              CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		return CLASS_E_NOAGGREGATION ;
	}
	*ppNewComponent = new CFixedPointNumberTest(pUnknownOuter) ;
	return S_OK ;
}

HRESULT __stdcall CFixedPointNumberTest::NondelegatingQueryInterface(const IID& iid,
                                                                     void** ppv)
{
	if (iid == IID_IFixedPointNumberTest)
	{
		return FinishQI(static_cast<IFixedPointNumberTest*>(this), ppv) ;
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
}

// ToDecVariant
// Takes a BSTR string and calls VarDecFromStr to return a VT_DECIMAL VARIANT.
HRESULT __stdcall CFixedPointNumberTest::ToDecVariant(BSTR strValue, VARIANT* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}

	VariantInit(result) ;

	if (strValue == NULL)
	{
		return E_INVALIDARG ;
	}

	// VarDecFromStr converts a string to a DECIMAL (VT_DECIMAL)
	// LOCALE_NOUSEROVERRIDE ensures consistent parsing regardless of user locale settings
	HRESULT hr = VarDecFromStr(strValue, LOCALE_NOUSEROVERRIDE, 0, &result->decVal) ;
	if (SUCCEEDED(hr))
	{
		result->vt = VT_DECIMAL ;
	}

	return hr ;
}

// ToCyVariant
// Takes a BSTR string and calls VarCyFromStr to return a VT_CY VARIANT.
HRESULT __stdcall CFixedPointNumberTest::ToCyVariant(BSTR strValue, VARIANT* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}

	VariantInit(result) ;

	if (strValue == NULL)
	{
		return E_INVALIDARG ;
	}

	// VarCyFromStr converts a string to a CURRENCY (VT_CY)
	// LOCALE_NOUSEROVERRIDE ensures consistent parsing regardless of user locale settings
	HRESULT hr = VarCyFromStr(strValue, LOCALE_NOUSEROVERRIDE, 0, &result->cyVal) ;
	if (SUCCEEDED(hr))
	{
		result->vt = VT_CY ;
	}

	return hr ;
}

// InitCyVariantWith3dot14
// [in, out] VARIANT*: discards whatever the caller passed in and reflects a
// VT_CY VARIANT holding 3.14 into it.
// If the caller passed a reference to a CURRENCY (VT_BYREF | VT_CY), 3.14 is
// written through the reference and the VARIANT itself is left as it is.
// If the caller passed a reference to a VARIANT (VT_BYREF | VT_VARIANT), 3.14 is
// written through the reference and the outer VARIANT is left as it is.
HRESULT __stdcall CFixedPointNumberTest::InitCyVariantWith3dot14(VARIANT* value)
{
	if (value == NULL)
	{
		return E_POINTER ;
	}

	if (value->vt == (VT_BYREF | VT_CY))
	{
		if (value->pcyVal == NULL)
		{
			return E_POINTER ;
		}
		value->pcyVal->int64 = 31400 ;
		return S_OK ;
	}

	if (value->vt == (VT_BYREF | VT_VARIANT))
	{
		if (value->pvarVal == NULL)
		{
			return E_POINTER ;
		}
		// Write through to the referenced VARIANT
		HRESULT hr = VariantClear(value->pvarVal) ;
		if (FAILED(hr))
		{
			return hr ;
		}
		value->pvarVal->cyVal.int64 = 31400 ;
		value->pvarVal->vt = VT_CY ;
		return S_OK ;
	}

	// The incoming VARIANT may own resources (e.g. a BSTR), so release them first.
	HRESULT hr = VariantClear(value) ;
	if (FAILED(hr))
	{
		return hr ;
	}

	// CY is a 64-bit integer scaled by 10000 (4 decimal places): 3.14 -> 31400
	value->cyVal.int64 = 31400 ;
	value->vt = VT_CY ;

	return S_OK ;
}

// InitDecVariantWith3dot14
// [in, out] VARIANT*: discards whatever the caller passed in and reflects a
// VT_DECIMAL VARIANT holding 3.14 into it.
// If the caller passed a reference to a DECIMAL (VT_BYREF | VT_DECIMAL), 3.14 is
// written through the reference and the VARIANT itself is left as it is.
// If the caller passed a reference to a VARIANT (VT_BYREF | VT_VARIANT), 3.14 is
// written through the reference and the outer VARIANT is left as it is.
HRESULT __stdcall CFixedPointNumberTest::InitDecVariantWith3dot14(VARIANT* value)
{
	if (value == NULL)
	{
		return E_POINTER ;
	}

	if (value->vt == (VT_BYREF | VT_DECIMAL))
	{
		if (value->pdecVal == NULL)
		{
			return E_POINTER ;
		}
		// Not a VARIANT-overlaid DECIMAL here, so the field order does not matter.
		memset(value->pdecVal, 0, sizeof(DECIMAL)) ;
		value->pdecVal->scale = 2 ;
		value->pdecVal->Lo64 = 314 ;
		return S_OK ;
	}

	if (value->vt == (VT_BYREF | VT_VARIANT))
	{
		if (value->pvarVal == NULL)
		{
			return E_POINTER ;
		}
		// Write through to the referenced VARIANT
		HRESULT hr = VariantClear(value->pvarVal) ;
		if (FAILED(hr))
		{
			return hr ;
		}
		// DECIMAL overlays the whole VARIANT: DECIMAL::wReserved shares its storage
		// with VARIANT::vt. Build the DECIMAL first, then set vt afterwards.
		// 3.14 = 314 * 10^-2 -> scale 2, sign 0, 96-bit integer 314
		DECIMAL dec ;
		memset(&dec, 0, sizeof(dec)) ;
		dec.scale = 2 ;
		dec.sign = 0 ;
		dec.Hi32 = 0 ;
		dec.Lo64 = 314 ;
		value->pvarVal->decVal = dec ;
		value->pvarVal->vt = VT_DECIMAL ;
		return S_OK ;
	}

	// The incoming VARIANT may own resources (e.g. a BSTR), so release them first.
	HRESULT hr = VariantClear(value) ;
	if (FAILED(hr))
	{
		return hr ;
	}

	// DECIMAL overlays the whole VARIANT: DECIMAL::wReserved shares its storage
	// with VARIANT::vt. Build the DECIMAL first, then set vt afterwards.
	// 3.14 = 314 * 10^-2 -> scale 2, sign 0, 96-bit integer 314
	DECIMAL dec ;
	memset(&dec, 0, sizeof(dec)) ;
	dec.scale = 2 ;
	dec.sign = 0 ;
	dec.Hi32 = 0 ;
	dec.Lo64 = 314 ;
	value->decVal = dec ;
	value->vt = VT_DECIMAL ;

	return S_OK ;
}