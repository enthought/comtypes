// A separate COM component for byte and narrow-string compatibility.
#include "CUnknown.h"
#include "Iface.h"

class CByteCompatibilityTest : public CUnknown, public IByteCompatibilityTest
{
public:
	static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent) ;

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv) ;
	virtual HRESULT __stdcall EchoInt8(char value, char* result) ;
	virtual HRESULT __stdcall EchoUint8(byte value, byte* result) ;
	virtual HRESULT __stdcall EchoLpStr(LPSTR value, LPSTR* result) ;
	virtual HRESULT __stdcall ReadCharPointer(unsigned char* value, byte* result) ;
	virtual HRESULT __stdcall GetCharacterBytes(byte value, unsigned char* result) ;

	CByteCompatibilityTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {} ;
} ;
