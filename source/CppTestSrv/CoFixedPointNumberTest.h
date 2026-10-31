// CoFixedPointNumberTest.h
// Test component for converting strings to VT_DECIMAL and VT_CY VARIANTs.
// ToDecVariant: takes a BSTR string and calls VarDecFromStr to return a VT_DECIMAL VARIANT.
// ToCyVariant: takes a BSTR string and calls VarCyFromStr to return a VT_CY VARIANT.
// DecomposeCyVariant: takes a VT_CY VARIANT and returns the int64 value.
// DecomposeDecVariant: takes a VT_DECIMAL VARIANT and returns sign (uint8_t), digits (SAFEARRAY of uint8_t), and exponent (int32_t).
// InitCyVariantWith3dot14: [in, out] VARIANT*. Discards the incoming content and reflects a VT_CY VARIANT holding 3.14.
//                          If the incoming VARIANT is VT_BYREF | VT_CY, writes 3.14 through the reference instead.
// InitDecVariantWith3dot14: [in, out] VARIANT*. Discards the incoming content and reflects a VT_DECIMAL VARIANT holding 3.14.
//                           If the incoming VARIANT is VT_BYREF | VT_DECIMAL, writes 3.14 through the reference instead.
#include "CUnknown.h"
#include "Iface.h"

class CFixedPointNumberTest : public CUnknown, public IFixedPointNumberTest
{
public:
	static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent) ;

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv) ;

	// Interface IFixedPointNumberTest
	virtual HRESULT __stdcall ToDecVariant(BSTR strValue, VARIANT* result) ;
	virtual HRESULT __stdcall ToCyVariant(BSTR strValue, VARIANT* result) ;
	virtual HRESULT __stdcall DecomposeCyVariant(VARIANT value, LONGLONG* result) ;
	virtual HRESULT __stdcall DecomposeDecVariant(VARIANT value, BYTE* sign, SAFEARRAY** digits, LONG* exponent) ;
	virtual HRESULT __stdcall InitCyVariantWith3dot14(VARIANT* value) ;
	virtual HRESULT __stdcall InitDecVariantWith3dot14(VARIANT* value) ;

	CFixedPointNumberTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {} ;
} ;