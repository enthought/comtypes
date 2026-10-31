// CoFixedPointNumberTest.cpp
// Test component for converting strings to VT_DECIMAL and VT_CY VARIANTs.
// ToDecVariant: takes a BSTR string and calls VarDecFromStr to return a VT_DECIMAL VARIANT.
// ToCyVariant: takes a BSTR string and calls VarCyFromStr to return a VT_CY VARIANT.
// DecomposeCyVariant: takes a VT_CY VARIANT and returns the int64 value.
// DecomposeDecVariant: takes a VT_DECIMAL VARIANT and returns sign (uint8_t), digits (SAFEARRAY of uint8_t), and exponent (int32_t).
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

// DecomposeCyVariant
// Takes a VT_CY VARIANT and returns the int64 value.
HRESULT __stdcall CFixedPointNumberTest::DecomposeCyVariant(VARIANT value, LONGLONG* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}

	if (value.vt != VT_CY)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	// CY is a 64-bit integer scaled by 10000 (4 decimal places)
	// The int64 value is stored in value.cyVal.int64
	*result = value.cyVal.int64 ;

	return S_OK ;
}

// DecomposeDecVariant
// Takes a VT_DECIMAL VARIANT and returns sign (uint8_t), digits (SAFEARRAY of uint8_t), and exponent (int32_t).
// This replicates Python's Decimal tuple representation: (sign, digits, exponent)
// where sign is 0 for positive, 1 for negative; digits is a tuple of individual digits; exponent is negative scale.
HRESULT __stdcall CFixedPointNumberTest::DecomposeDecVariant(VARIANT value, BYTE* sign, SAFEARRAY** digits, LONG* exponent)
{
	if (sign == NULL || digits == NULL || exponent == NULL)
	{
		return E_POINTER ;
	}

	if (value.vt != VT_DECIMAL)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	// DECIMAL structure (tagDEC):
	// wReserved (16 bits) - contains scale (lo byte) and sign (hi byte)
	// Hi32 (32 bits) - high 32 bits of 96-bit integer
	// Lo64 (64 bits) - low 64 bits of 96-bit integer
	//
	// Scale: number of digits to the right of decimal point (0-28)
	// Sign: 0x00 for positive, 0x80 for negative
	// The 96-bit integer is (Hi32 << 64) | Lo64

	DECIMAL dec = value.decVal ;

	// Extract sign (0 for positive, 1 for negative)
	*sign = (dec.sign & 0x80) ? 1 : 0 ;

	// Extract scale (number of decimal places)
	BYTE scale = dec.scale ;

	// Combine Hi32 and Lo64 into a 96-bit integer
	// We need to convert this to individual decimal digits
	// Since we can't easily work with 96-bit integers in C++, we'll use a string approach
	// by converting the DECIMAL to a string first, then parsing it

	// Use VarBstrFromDec to get a string representation
	BSTR bstrValue = NULL ;
	HRESULT hr = VarBstrFromDec(&dec, LOCALE_NOUSEROVERRIDE, 0, &bstrValue) ;
	if (FAILED(hr))
	{
		return hr ;
	}

	// Convert BSTR to char array for processing
	int len = SysStringLen(bstrValue) ;
	char* str = new char[len + 1] ;
	for (int i = 0; i < len; i++)
	{
		str[i] = static_cast<char>(bstrValue[i]) ;
	}
	str[len] = '\0' ;

	SysFreeString(bstrValue) ;

	// Parse the string to extract digits
	// Format could be like "12345", "-12345", "123.45", "-123.45", "0.00123", etc.
	char* p = str ;
	bool negative = false ;
	if (*p == '-')
	{
		negative = true ;
		p++ ;
	}
	else if (*p == '+')
	{
		p++ ;
	}

	// Count digits and find decimal point
	int digitCount = 0 ;
	int decimalPos = -1 ;
	for (int i = 0; i < len; i++)
	{
		if (p[i] >= '0' && p[i] <= '9')
		{
			digitCount++ ;
		}
		else if (p[i] == '.')
		{
			decimalPos = i ;
		}
	}

	// Create SAFEARRAY for digits
	SAFEARRAYBOUND bound ;
	bound.lLbound = 0 ;
	bound.cElements = digitCount ;
	*digits = SafeArrayCreate(VT_UI1, 1, &bound) ;
	if (*digits == NULL)
	{
		delete[] str ;
		return E_OUTOFMEMORY ;
	}

	// Lock the array and fill with digits
	BYTE* pDigits = NULL ;
	hr = SafeArrayAccessData(*digits, (void**)&pDigits) ;
	if (FAILED(hr))
	{
		SafeArrayDestroy(*digits) ;
		*digits = NULL ;
		delete[] str ;
		return hr ;
	}

	int digitIndex = 0 ;
	for (int i = 0; i < len; i++)
	{
		if (p[i] >= '0' && p[i] <= '9')
		{
			pDigits[digitIndex++] = static_cast<BYTE>(p[i] - '0') ;
		}
	}

	SafeArrayUnaccessData(*digits) ;
	delete[] str ;

	// Calculate exponent: negative of scale (number of decimal places)
	// If there was a decimal point, the scale tells us how many digits after it
	// The exponent in Python's Decimal tuple is negative scale
	*exponent = -static_cast<LONG>(scale) ;

	return S_OK ;
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