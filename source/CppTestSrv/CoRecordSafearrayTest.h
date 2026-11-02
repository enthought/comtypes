/*
 * CoRecordSafearrayTest.h - Component for record SAFEARRAY round-trip tests
 *
 * This component implements two test pathways:
 * 1. IDispatch dual vtable (IDualRecordSafearrayTest)
 * 2. IDispatch dispinterface Invoke (IDispRecordSafearrayTest)
 */

#ifndef __CoRecordSafearrayTest_h__
#define __CoRecordSafearrayTest_h__

#include "Iface.h"
#include "CUnknown.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispRecordSafearrayTest;

///////////////////////////////////////////////////////////
//
// Component for record SAFEARRAY round-trip tests
//
class CRecordSafearrayTest : public CUnknown,
                              public IDualRecordSafearrayTest
{
public:
    static HRESULT CreateInstance(IUnknown* pUnknownOuter,
                                  CUnknown** ppNewComponent);

private:
    DECLARE_IUNKNOWN
    virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv);

    // Wrap a single record into a SAFEARRAY of length 1
    virtual HRESULT __stdcall WrapValueRecord(
        StructValueRecordSafearrayTest* record,
        SAFEARRAY** records);

    // Return the first record from a SAFEARRAY
    virtual HRESULT __stdcall FirstValueRecord(
        SAFEARRAY* records,
        StructValueRecordSafearrayTest* record);

    // Wrap a single BSTR record into a SAFEARRAY of length 1
    virtual HRESULT __stdcall WrapBstrRecord(
        StructBstrRecordSafearrayTest* record,
        SAFEARRAY** records);

    // Return the first BSTR record from a SAFEARRAY
    virtual HRESULT __stdcall FirstBstrRecord(
        SAFEARRAY* records,
        StructBstrRecordSafearrayTest* record);

    // IDispatch
    virtual HRESULT __stdcall GetTypeInfoCount(UINT* pCountTypeInfo);
    virtual HRESULT __stdcall GetTypeInfo(UINT iTypeInfo, LCID, ITypeInfo** ppITypeInfo);
    virtual HRESULT __stdcall GetIDsOfNames(const IID& iid, OLECHAR** arrayNames,
                                            UINT countNames, LCID, DISPID* arrayDispIDs);
    virtual HRESULT __stdcall Invoke(DISPID dispidMember, const IID& iid, LCID,
                                     WORD wFlags, DISPPARAMS* pDispParams,
                                     VARIANT* pvarResult, EXCEPINFO* pExcepInfo,
                                     UINT* pArgErr);

    // Initialization
    virtual HRESULT Init();

    CRecordSafearrayTest(IUnknown* pUnknownOuter);
    ~CRecordSafearrayTest();

    ITypeInfo* m_pITypeInfo;
};

#endif
