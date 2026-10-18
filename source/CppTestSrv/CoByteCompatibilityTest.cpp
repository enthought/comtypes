#include <objbase.h>
#include <string.h>
#include "CoByteCompatibilityTest.h"

HRESULT CByteCompatibilityTest::CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		return CLASS_E_NOAGGREGATION ;
	}
	*ppNewComponent = new CByteCompatibilityTest(pUnknownOuter) ;
	return S_OK ;
}

HRESULT __stdcall CByteCompatibilityTest::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
	if (iid == IID_IByteCompatibilityTest)
	{
		return FinishQI(static_cast<IByteCompatibilityTest*>(this), ppv) ;
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
}

HRESULT __stdcall CByteCompatibilityTest::EchoInt8(char value, char* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = value ;
	return S_OK ;
}

HRESULT __stdcall CByteCompatibilityTest::EchoUint8(byte value, byte* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = value ;
	return S_OK ;
}

HRESULT __stdcall CByteCompatibilityTest::EchoLpStr(LPSTR value, LPSTR* result)
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

HRESULT __stdcall CByteCompatibilityTest::ReadCharPointer(unsigned char* value, byte* result)
{
	if (value == NULL || result == NULL)
	{
		return E_POINTER ;
	}
	// A raw char* marshals one character, not a NUL-terminated string.
	*result = static_cast<byte>(*value) ;
	return S_OK ;
}

HRESULT __stdcall CByteCompatibilityTest::GetCharacterBytes(byte value, unsigned char* result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = value ;
	return S_OK ;
}
