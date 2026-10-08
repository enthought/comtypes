#include <objbase.h>
#include "CoNumericSafearrayTest.h"

HRESULT CNumericSafearrayTest::CreateInstance(IUnknown* pUnknownOuter,
                                             CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		return CLASS_E_NOAGGREGATION ;
	}
	*ppNewComponent = new CNumericSafearrayTest(pUnknownOuter) ;
	return S_OK ;
}

HRESULT __stdcall CNumericSafearrayTest::NondelegatingQueryInterface(const IID& iid,
                                                                    void** ppv)
{
	if (iid == IID_INumericSafearrayTest)
	{
		return FinishQI(static_cast<INumericSafearrayTest*>(this), ppv) ;
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
}

// Return native arrays with fixed values, independent of Python/NumPy input.
// Two-dimensional arrays have two rows and are filled in column-major order.
template <class T>
static HRESULT CreateNumericArray(VARTYPE vartype, short dimensions, long count,
                                 const T (&values)[3], SAFEARRAY** result)
{
	if (result == NULL)
	{
		return E_POINTER ;
	}
	*result = NULL ;
	if ((dimensions != 1 && dimensions != 2) || count <= 0 ||
	    (dimensions == 2 && count % 2 != 0))
	{
		return E_INVALIDARG ;
	}
	SAFEARRAYBOUND bounds[2] = {{static_cast<ULONG>(count), -2}, {1, 5}} ;
	if (dimensions == 2)
	{
		bounds[0].cElements = 2 ;
		bounds[1].cElements = count / 2 ;
	}
	SAFEARRAY* array = SafeArrayCreate(vartype, dimensions, bounds) ;
	if (array == NULL)
	{
		return E_OUTOFMEMORY ;
	}
	T* data = NULL ;
	HRESULT hr = SafeArrayAccessData(array, reinterpret_cast<void**>(&data)) ;
	if (FAILED(hr))
	{
		SafeArrayDestroy(array) ;
		return hr ;
	}
	for (long i = 0; i < count; ++i)
	{
		data[i] = values[i % 3] ;
	}
	hr = SafeArrayUnaccessData(array) ;
	if (FAILED(hr))
	{
		SafeArrayDestroy(array) ;
		return hr ;
	}
	*result = array ;
	return S_OK ;
}

HRESULT __stdcall CNumericSafearrayTest::GetInt8Array(short dimensions, long count, SAFEARRAY** result)
{
	const CHAR values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I1, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetUint8Array(short dimensions, long count, SAFEARRAY** result)
{
	const BYTE values[3] = {0, 42, 255} ;
	return CreateNumericArray(VT_UI1, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetInt16Array(short dimensions, long count, SAFEARRAY** result)
{
	const SHORT values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I2, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetUint16Array(short dimensions, long count, SAFEARRAY** result)
{
	const USHORT values[3] = {0, 42, 65535} ;
	return CreateNumericArray(VT_UI2, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetInt32Array(short dimensions, long count, SAFEARRAY** result)
{
	const LONG values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I4, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetUint32Array(short dimensions, long count, SAFEARRAY** result)
{
	const ULONG values[3] = {0, 42, 4294967295UL} ;
	return CreateNumericArray(VT_UI4, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetInt64Array(short dimensions, long count, SAFEARRAY** result)
{
	const LONGLONG values[3] = {-7, 0, 42} ;
	return CreateNumericArray(VT_I8, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetUint64Array(short dimensions, long count, SAFEARRAY** result)
{
	const ULONGLONG values[3] = {0, 42, 4294967297ULL} ;
	return CreateNumericArray(VT_UI8, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetFloat32Array(short dimensions, long count, SAFEARRAY** result)
{
	const FLOAT values[3] = {1.25f, -2.5f, 3.75f} ;
	return CreateNumericArray(VT_R4, dimensions, count, values, result) ;
}

HRESULT __stdcall CNumericSafearrayTest::GetFloat64Array(short dimensions, long count, SAFEARRAY** result)
{
	const DOUBLE values[3] = {1.25, -2.5, 3.75} ;
	return CreateNumericArray(VT_R8, dimensions, count, values, result) ;
}
