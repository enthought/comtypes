// CoVariantErrorTest.cpp
// Test double for VT_ERROR VARIANT marshalling.
//
// CreateVariantError: constructs a VARIANT entirely in native C++ — sets vt
//   to VT_ERROR and scode to errorCode directly.  This matches what VBA's 
//   CVErr() produces.
// GetErrorCode: reads the SCODE from a VT_ERROR VARIANT and returns it as a
//   signed LONG.  This matches what Excel VBA's CLng(myErrCell.Value).
#include <objbase.h>
#include "CoVariantErrorTest.h"


HRESULT CVariantErrorTest::CreateInstance(IUnknown* pUnknownOuter,
                                          CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		return CLASS_E_NOAGGREGATION ;
	}
	*ppNewComponent = new CVariantErrorTest(pUnknownOuter) ;
	return S_OK ;
}

HRESULT __stdcall CVariantErrorTest::NondelegatingQueryInterface(const IID& iid,
                                                                 void** ppv)
{
	if (iid == IID_IVariantErrorTest)
	{
		return FinishQI(static_cast<IVariantErrorTest*>(this), ppv) ;
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
}

// CreateVariantError
// Sets *result to VT_ERROR with the given SCODE.
HRESULT __stdcall CVariantErrorTest::CreateVariantError(LONG errorCode, VARIANT* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}

	VariantInit(result) ;

	result->vt    = VT_ERROR ;
	result->scode = static_cast<SCODE>(errorCode) ;

	return S_OK ;
}

// GetErrorCode
// Extracts the SCODE from a VT_ERROR VARIANT.
// Returns DISP_E_TYPEMISMATCH if the variant is not VT_ERROR.
HRESULT __stdcall CVariantErrorTest::GetErrorCode(VARIANT value, LONG* errorCode)
{
	if (errorCode == NULL)
	{
		return E_POINTER ;
	}

	if (value.vt != VT_ERROR)
	{
		return DISP_E_TYPEMISMATCH ;
	}

	*errorCode = value.scode ;
	return S_OK ;
}
