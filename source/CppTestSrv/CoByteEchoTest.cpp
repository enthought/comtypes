#include <objbase.h>
#include <string.h>
#include "CoByteEchoTest.h"

HRESULT CByteEchoTest::CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		return CLASS_E_NOAGGREGATION ;
	}
	*ppNewComponent = new CByteEchoTest(pUnknownOuter) ;
	return S_OK ;
}

HRESULT __stdcall CByteEchoTest::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
	if (iid == IID_IByteEchoTest)
	{
		return FinishQI(static_cast<IByteEchoTest*>(this), ppv) ;
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
}

HRESULT __stdcall CByteEchoTest::EchoInt8(char value, char* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = value ;
	return S_OK ;
}

HRESULT __stdcall CByteEchoTest::EchoUint8(byte value, byte* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = value ;
	return S_OK ;
}

HRESULT __stdcall CByteEchoTest::EchoLpStr(LPSTR value, LPSTR* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = NULL ;
	if (value == NULL)
	{
		return E_INVALIDARG ;
	}
	size_t size = strlen(value) + 1 ;
	*result = static_cast<LPSTR>(CoTaskMemAlloc(size)) ;
	if (*result == NULL)
	{
		return E_OUTOFMEMORY ;
	}
	memcpy(*result, value, size) ;
	return S_OK ;
}

HRESULT __stdcall CByteEchoTest::ReadCharPointer(unsigned char* value, byte* result)
{
	if (value == NULL || result == NULL)
	{
		return E_POINTER ;
	}
	// A raw char* marshals one character, not a NUL-terminated string.
	*result = static_cast<byte>(*value) ;
	return S_OK ;
}
