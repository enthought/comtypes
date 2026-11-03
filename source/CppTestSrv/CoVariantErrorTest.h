// CoVariantErrorTest.h
// Test double for VT_ERROR VARIANT marshalling.
#include "CUnknown.h"
#include "Iface.h"

class CVariantErrorTest : public CUnknown, public IVariantErrorTest
{
public:
	static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent) ;

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv) ;

	// Interface IVariantErrorTest
	virtual HRESULT __stdcall CreateVariantError(LONG errorCode, VARIANT* result) ;
	virtual HRESULT __stdcall GetErrorCode(VARIANT value, LONG* errorCode) ;

	CVariantErrorTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {} ;
} ;
