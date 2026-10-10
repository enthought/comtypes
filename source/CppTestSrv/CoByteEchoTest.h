// A separate COM component for byte and narrow-string round trips.
#include "CUnknown.h"
#include "Iface.h"

class CByteEchoTest : public CUnknown, public IByteEchoTest
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

	CByteEchoTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {} ;
} ;
