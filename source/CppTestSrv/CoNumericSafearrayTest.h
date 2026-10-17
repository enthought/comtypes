// A focused COM test double for native numeric SAFEARRAY output.
#include "CUnknown.h"
#include "Iface.h"

class CNumericSafearrayTest : public CUnknown, public INumericSafearrayTest
{
public:
	static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent) ;

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv) ;

	// Interface INumericSafearrayTest
	virtual HRESULT __stdcall GetInt8Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetUint8Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetInt16Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetUint16Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetInt32Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetUint32Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetInt64Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetUint64Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetFloat32Array(short dimensions, long count, SAFEARRAY** result) ;
	virtual HRESULT __stdcall GetFloat64Array(short dimensions, long count, SAFEARRAY** result) ;

	CNumericSafearrayTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {} ;
} ;
