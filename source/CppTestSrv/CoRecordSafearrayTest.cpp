/*
 * CoRecordSafearrayTest.cpp - Component for record SAFEARRAY round-trip tests
 *
 * This component implements two test pathways:
 * 1. IDispatch dual vtable (IDualRecordSafearrayTest)
 * 2. IDispatch dispinterface Invoke (IDispRecordSafearrayTest)
 */

#include <objbase.h>
#include <string.h>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h"
#include "CoRecordSafearrayTest.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispRecordSafearrayTest;

// Explicit GUIDs for record structs: MIDL does not emit __declspec(uuid(...)) for
// typedef'd structs, so __uuidof() cannot be used directly.
static const GUID GUID_StructValueRecordSafearrayTest =
    { 0x6B298B7B, 0xCA0E, 0x431D, { 0x98, 0x8C, 0xFE, 0xD3, 0xCE, 0xB5, 0x97, 0x29 } };
static const GUID GUID_StructBstrRecordSafearrayTest =
    { 0x5F72F918, 0xEC40, 0x42C1, { 0xA0, 0x34, 0xF6, 0x60, 0x89, 0x02, 0xEC, 0xB3 } };

static inline void trace(const char* msg)
    { Util::Trace("CoRecordSafearrayTest", msg, S_OK); }
static inline void trace(const char* msg, HRESULT hr)
    { Util::Trace("CoRecordSafearrayTest", msg, hr); }

// ========================================================================
// Helper Functions
// ========================================================================

// Return a reversed deep copy of a record SAFEARRAY.
static HRESULT ReverseRecords(
    SAFEARRAY* records, const GUID& recGuid, SAFEARRAY** reversed)
{
    if (reversed == NULL)
        return E_POINTER;
    *reversed = NULL;
    if (records == NULL)
        return E_POINTER;

    if (::SafeArrayGetDim(records) != 1)
        return DISP_E_TYPEMISMATCH;

    VARTYPE vt;
    HRESULT hr = ::SafeArrayGetVartype(records, &vt);
    if (FAILED(hr) || vt != VT_RECORD)
        return DISP_E_TYPEMISMATCH;

    IRecordInfo* pRecordInfo = NULL;
    hr = ::SafeArrayGetRecordInfo(records, &pRecordInfo);
    if (FAILED(hr) || pRecordInfo == NULL)
        return DISP_E_TYPEMISMATCH;

    GUID actual_guid;
    hr = pRecordInfo->GetGuid(&actual_guid);
    if (FAILED(hr) || !IsEqualGUID(actual_guid, recGuid))
    {
        pRecordInfo->Release();
        return DISP_E_TYPEMISMATCH;
    }

    LONG lbound, ubound;
    hr = ::SafeArrayGetLBound(records, 1, &lbound);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }
    hr = ::SafeArrayGetUBound(records, 1, &ubound);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }

    LONG count = ubound - lbound + 1;
    if (count <= 0)
    {
        pRecordInfo->Release();
        return E_INVALIDARG;
    }

    ULONG size = 0;
    hr = pRecordInfo->GetSize(&size);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }

    SAFEARRAY* psa = ::SafeArrayCreateVectorEx(VT_RECORD, 0, count, pRecordInfo);
    if (psa == NULL)
    {
        pRecordInfo->Release();
        return E_OUTOFMEMORY;
    }

    void* pvSrc = NULL;
    void* pvDst = NULL;
    hr = ::SafeArrayAccessData(records, &pvSrc);
    if (FAILED(hr))
    {
        ::SafeArrayDestroy(psa);
        pRecordInfo->Release();
        return hr;
    }
    hr = ::SafeArrayAccessData(psa, &pvDst);
    if (FAILED(hr))
    {
        ::SafeArrayUnaccessData(records);
        ::SafeArrayDestroy(psa);
        pRecordInfo->Release();
        return hr;
    }

    BYTE* src = static_cast<BYTE*>(pvSrc);
    BYTE* dst = static_cast<BYTE*>(pvDst);
    for (LONG i = 0; i < count && SUCCEEDED(hr); ++i)
    {
        BYTE* dstElem = dst + (size_t)i * size;
        BYTE* srcElem = src + (size_t)(count - 1 - i) * size;
        hr = pRecordInfo->RecordInit(dstElem);
        if (SUCCEEDED(hr))
            hr = pRecordInfo->RecordCopy(srcElem, dstElem);
    }

    ::SafeArrayUnaccessData(psa);
    ::SafeArrayUnaccessData(records);
    pRecordInfo->Release();

    if (FAILED(hr))
    {
        ::SafeArrayDestroy(psa);
        return hr;
    }

    *reversed = psa;
    return S_OK;
}

// ========================================================================
// CRecordSafearrayTest Implementation
// ========================================================================

CRecordSafearrayTest::CRecordSafearrayTest(IUnknown* pUnknownOuter)
    : CUnknown(pUnknownOuter), m_pITypeInfo(NULL)
{
}

CRecordSafearrayTest::~CRecordSafearrayTest()
{
    if (m_pITypeInfo != NULL)
        m_pITypeInfo->Release();
    trace("Destroy self.");
}

HRESULT CRecordSafearrayTest::CreateInstance(
    IUnknown* pUnknownOuter, CUnknown** ppNewComponent)
{
    if (pUnknownOuter != NULL)
        return CLASS_E_NOAGGREGATION;

    *ppNewComponent = new CRecordSafearrayTest(pUnknownOuter);
    return S_OK;
}

HRESULT __stdcall CRecordSafearrayTest::NondelegatingQueryInterface(
    const IID& iid, void** ppv)
{
    if (iid == IID_IDualRecordSafearrayTest)
    {
        return FinishQI(static_cast<IDualRecordSafearrayTest*>(this), ppv);
    }
    else if (iid == DIID_IDispRecordSafearrayTest)
    {
        trace("Queried for IDispRecordSafearrayTest.");
        return FinishQI(static_cast<IDispatch*>(this), ppv);
    }
    else if (iid == IID_IDispatch)
    {
        trace("Queried for IDispatch.");
        return FinishQI(static_cast<IDispatch*>(this), ppv);
    }
    return CUnknown::NondelegatingQueryInterface(iid, ppv);
}

HRESULT CRecordSafearrayTest::Init()
{
    HRESULT hr;

    if (m_pITypeInfo == NULL)
    {
        ITypeLib* pITypeLib = NULL;
        hr = ::LoadRegTypeLib(LIBID_ComtypesCppTestSrvLib,
                              1, 0, 0x00, &pITypeLib);
        if (FAILED(hr))
        {
            trace("LoadRegTypeLib Failed.", hr);
            return hr;
        }

        hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualRecordSafearrayTest,
                                          &m_pITypeInfo);
        pITypeLib->Release();
        if (FAILED(hr))
        {
            trace("GetTypeInfoOfGuid failed.", hr);
            return hr;
        }
    }
    return S_OK;
}

// ------------------------------------------------------------------------
// IDispatch Implementation
// ------------------------------------------------------------------------

HRESULT __stdcall CRecordSafearrayTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
    trace("GetTypeInfoCount call succeeded.");
    *pCountTypeInfo = 1;
    return S_OK;
}

HRESULT __stdcall CRecordSafearrayTest::GetTypeInfo(
    UINT iTypeInfo, LCID, ITypeInfo** ppITypeInfo)
{
    *ppITypeInfo = NULL;

    if (iTypeInfo != 0)
    {
        trace("GetTypeInfo call failed -- bad iTypeInfo index.");
        return DISP_E_BADINDEX;
    }

    trace("GetTypeInfo call succeeded.");
    m_pITypeInfo->AddRef();
    *ppITypeInfo = m_pITypeInfo;
    return S_OK;
}

HRESULT __stdcall CRecordSafearrayTest::GetIDsOfNames(
    const IID& iid, OLECHAR** arrayNames, UINT countNames, LCID, DISPID* arrayDispIDs)
{
    if (iid != IID_NULL)
    {
        trace("GetIDsOfNames call failed -- bad IID.");
        return DISP_E_UNKNOWNINTERFACE;
    }

    trace("GetIDsOfNames call succeeded.");
    HRESULT hr = m_pITypeInfo->GetIDsOfNames(arrayNames, countNames, arrayDispIDs);
    return hr;
}

HRESULT __stdcall CRecordSafearrayTest::Invoke(
    DISPID dispidMember, const IID& iid, LCID, WORD wFlags,
    DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* pArgErr)
{
    if (iid != IID_NULL)
    {
        trace("Invoke call failed -- bad IID.");
        return DISP_E_UNKNOWNINTERFACE;
    }

    ::SetErrorInfo(0, NULL);
    trace("Invoke call succeeded.");
    HRESULT hr = m_pITypeInfo->Invoke(
        static_cast<IDualRecordSafearrayTest*>(this),
        dispidMember, wFlags, pDispParams,
        pvarResult, pExcepInfo, pArgErr);
    return hr;
}

// ------------------------------------------------------------------------
// IDualRecordSafearrayTest Implementation
// ------------------------------------------------------------------------

// Wrap a single value record into a SAFEARRAY of length 1
HRESULT __stdcall CRecordSafearrayTest::WrapValueRecord(
    StructValueRecordSafearrayTest* record, SAFEARRAY** records)
{
    trace("WrapValueRecord called");
    if (records == NULL)
        return E_POINTER;
    *records = NULL;

    if (record == NULL)
        return E_POINTER;

    // Get IRecordInfo from typelib
    IRecordInfo* pRecordInfo = NULL;
    HRESULT hr = ::GetRecordInfoFromGuids(
        LIBID_ComtypesCppTestSrvLib, 1, 0, 0, GUID_StructValueRecordSafearrayTest, &pRecordInfo);
    if (FAILED(hr))
    {
        trace("GetRecordInfoFromGuids failed", hr);
        return hr;
    }

    // Create SAFEARRAY with 1 element
    SAFEARRAY* psa = ::SafeArrayCreateVectorEx(VT_RECORD, 0, 1, pRecordInfo);
    if (psa == NULL)
    {
        pRecordInfo->Release();
        return E_OUTOFMEMORY;
    }

    // Access data and copy the input record using RecordCopy
    void* pvData = NULL;
    hr = ::SafeArrayAccessData(psa, &pvData);
    if (FAILED(hr))
    {
        ::SafeArrayDestroy(psa);
        pRecordInfo->Release();
        return hr;
    }

    // Initialize destination element before RecordCopy (S5/N5)
    // RecordCopy expects destination to be initialized; otherwise it may
    // try to free garbage BSTR/pointer values.
    hr = pRecordInfo->RecordInit(pvData);
    if (FAILED(hr))
    {
        ::SafeArrayUnaccessData(psa);
        ::SafeArrayDestroy(psa);
        pRecordInfo->Release();
        return hr;
    }

    // Use RecordCopy to properly copy the record (handles BSTR/pointers)
    hr = pRecordInfo->RecordCopy(record, pvData);
    ::SafeArrayUnaccessData(psa);
    pRecordInfo->Release();

    if (FAILED(hr))
    {
        ::SafeArrayDestroy(psa);
        return hr;
    }

    *records = psa;
    return S_OK;
}

// Return the first value record from a SAFEARRAY
HRESULT __stdcall CRecordSafearrayTest::FirstValueRecord(
    SAFEARRAY* records, StructValueRecordSafearrayTest* record)
{
    trace("FirstValueRecord called");
    if (record == NULL)
        return E_POINTER;

    // Zero out the output record first (COM out parameter rule)
    memset(record, 0, sizeof(StructValueRecordSafearrayTest));

    if (records == NULL)
        return E_POINTER;

    // Validate dimensions
    UINT dims = ::SafeArrayGetDim(records);
    if (dims != 1)
        return DISP_E_TYPEMISMATCH;

    // Validate VARTYPE
    VARTYPE vt;
    HRESULT hr = ::SafeArrayGetVartype(records, &vt);
    if (FAILED(hr) || vt != VT_RECORD)
        return DISP_E_TYPEMISMATCH;

    // Get IRecordInfo
    IRecordInfo* pRecordInfo = NULL;
    hr = ::SafeArrayGetRecordInfo(records, &pRecordInfo);
    if (FAILED(hr) || pRecordInfo == NULL)
        return DISP_E_TYPEMISMATCH;

    // Validate IID
    GUID actual_guid;
    hr = pRecordInfo->GetGuid(&actual_guid);
    if (FAILED(hr) || !IsEqualGUID(actual_guid, GUID_StructValueRecordSafearrayTest))
    {
        pRecordInfo->Release();
        return DISP_E_TYPEMISMATCH;
    }

    // Get bounds
    LONG lbound, ubound;
    hr = ::SafeArrayGetLBound(records, 1, &lbound);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }
    hr = ::SafeArrayGetUBound(records, 1, &ubound);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }

    LONG count = ubound - lbound + 1;
    if (count <= 0)
    {
        pRecordInfo->Release();
        return E_INVALIDARG;
    }

    // Access data and copy first element using RecordCopy
    void* pvData = NULL;
    hr = ::SafeArrayAccessData(records, &pvData);
    if (FAILED(hr))
    {
        pRecordInfo->Release();
        return hr;
    }

    // Use RecordCopy to properly copy the record (handles BSTR/pointers)
    hr = pRecordInfo->RecordCopy(pvData, record);
    ::SafeArrayUnaccessData(records);
    pRecordInfo->Release();

    return hr;
}

// Return a reversed deep copy of a record SAFEARRAY.
HRESULT __stdcall CRecordSafearrayTest::ReverseValueRecords(
    SAFEARRAY* records, SAFEARRAY** reversed)
{
    trace("ReverseValueRecords called");
    return ReverseRecords(records, GUID_StructValueRecordSafearrayTest, reversed);
}

// Wrap a single BSTR record into a SAFEARRAY of length 1
HRESULT __stdcall CRecordSafearrayTest::WrapBstrRecord(
    StructBstrRecordSafearrayTest* record, SAFEARRAY** records)
{
    trace("WrapBstrRecord called");
    if (records == NULL)
        return E_POINTER;
    *records = NULL;

    if (record == NULL)
        return E_POINTER;

    // Get IRecordInfo from typelib
    IRecordInfo* pRecordInfo = NULL;
    HRESULT hr = ::GetRecordInfoFromGuids(
        LIBID_ComtypesCppTestSrvLib, 1, 0, 0, GUID_StructBstrRecordSafearrayTest, &pRecordInfo);
    if (FAILED(hr))
    {
        trace("GetRecordInfoFromGuids failed", hr);
        return hr;
    }

    // Create SAFEARRAY with 1 element
    SAFEARRAY* psa = ::SafeArrayCreateVectorEx(VT_RECORD, 0, 1, pRecordInfo);
    if (psa == NULL)
    {
        pRecordInfo->Release();
        return E_OUTOFMEMORY;
    }

    // Access data and copy the input record using RecordCopy
    void* pvData = NULL;
    hr = ::SafeArrayAccessData(psa, &pvData);
    if (FAILED(hr))
    {
        ::SafeArrayDestroy(psa);
        pRecordInfo->Release();
        return hr;
    }

    // Initialize destination element before RecordCopy (S5/N5)
    // RecordCopy expects destination to be initialized; otherwise it may
    // try to free garbage BSTR/pointer values.
    hr = pRecordInfo->RecordInit(pvData);
    if (FAILED(hr))
    {
        ::SafeArrayUnaccessData(psa);
        ::SafeArrayDestroy(psa);
        pRecordInfo->Release();
        return hr;
    }

    // Use RecordCopy to properly copy the record (handles BSTR/pointers)
    hr = pRecordInfo->RecordCopy(record, pvData);
    ::SafeArrayUnaccessData(psa);
    pRecordInfo->Release();

    if (FAILED(hr))
    {
        ::SafeArrayDestroy(psa);
        return hr;
    }

    *records = psa;
    return S_OK;
}

// Return the first BSTR record from a SAFEARRAY
HRESULT __stdcall CRecordSafearrayTest::FirstBstrRecord(
    SAFEARRAY* records, StructBstrRecordSafearrayTest* record)
{
    trace("FirstBstrRecord called");
    if (record == NULL)
        return E_POINTER;

    // Zero out the output record first (COM out parameter rule)
    memset(record, 0, sizeof(StructBstrRecordSafearrayTest));

    if (records == NULL)
        return E_POINTER;

    // Validate dimensions
    UINT dims = ::SafeArrayGetDim(records);
    if (dims != 1)
        return DISP_E_TYPEMISMATCH;

    // Validate VARTYPE
    VARTYPE vt;
    HRESULT hr = ::SafeArrayGetVartype(records, &vt);
    if (FAILED(hr) || vt != VT_RECORD)
        return DISP_E_TYPEMISMATCH;

    // Get IRecordInfo
    IRecordInfo* pRecordInfo = NULL;
    hr = ::SafeArrayGetRecordInfo(records, &pRecordInfo);
    if (FAILED(hr) || pRecordInfo == NULL)
        return DISP_E_TYPEMISMATCH;

    // Validate IID
    GUID actual_guid;
    hr = pRecordInfo->GetGuid(&actual_guid);
    if (FAILED(hr) || !IsEqualGUID(actual_guid, GUID_StructBstrRecordSafearrayTest))
    {
        pRecordInfo->Release();
        return DISP_E_TYPEMISMATCH;
    }

    // Get bounds
    LONG lbound, ubound;
    hr = ::SafeArrayGetLBound(records, 1, &lbound);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }
    hr = ::SafeArrayGetUBound(records, 1, &ubound);
    if (FAILED(hr)) { pRecordInfo->Release(); return hr; }

    LONG count = ubound - lbound + 1;
    if (count <= 0)
    {
        pRecordInfo->Release();
        return E_INVALIDARG;
    }

    // Access data and copy first element using RecordCopy
    void* pvData = NULL;
    hr = ::SafeArrayAccessData(records, &pvData);
    if (FAILED(hr))
    {
        pRecordInfo->Release();
        return hr;
    }

    // Use RecordCopy to properly copy the record (handles BSTR/pointers)
    hr = pRecordInfo->RecordCopy(pvData, record);
    ::SafeArrayUnaccessData(records);
    pRecordInfo->Release();

    return hr;
}

// Return a reversed deep copy of a record SAFEARRAY.
HRESULT __stdcall CRecordSafearrayTest::ReverseBstrRecords(
    SAFEARRAY* records, SAFEARRAY** reversed)
{
    trace("ReverseBstrRecords called");
    return ReverseRecords(records, GUID_StructBstrRecordSafearrayTest, reversed);
}
