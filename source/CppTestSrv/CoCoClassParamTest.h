/*
 * CoCoClassParamTest.h - Component for coclass return or argument type tests
 *
 */

#ifndef __CoCoClassParamTest_h__
#define __CoCoClassParamTest_h__

#include "Iface.h"
#include "CUnknown.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispCoClassParamTest;

///////////////////////////////////////////////////////////
//
// CCoClassParamTestItem
//
class CCoClassParamTestItem : public CUnknown, public ICoClassParamTestItem
{
public:
	CCoClassParamTestItem(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {}

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid,
	                                                      void** ppv);

	// Interface ICoClassParamTestItem
	virtual HRESULT __stdcall GetValue(long* result);
	virtual HRESULT __stdcall get_SelfAsItf(ICoClassParamTestItem** result);
	virtual HRESULT __stdcall get_SelfAsCC(CoCoClassParamTestItem** result);
	virtual HRESULT __stdcall IsSameItemAsCC(CoCoClassParamTestItem* item,
	                                          VARIANT_BOOL* result);
	virtual HRESULT __stdcall IsSameItemAsItf(ICoClassParamTestItem* item,
	                                          VARIANT_BOOL* result);
};

///////////////////////////////////////////////////////////
//
// CCoClassParamTest
//
class CCoClassParamTest : public CUnknown, public IDualCoClassParamTest
{
public:
	// Creation
	static HRESULT CreateInstance(IUnknown* pUnknownOuter,
	                              CUnknown** ppNewComponent);

private:
	// Declare the delegating IUnknown.
	DECLARE_IUNKNOWN

	// IUnknown
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid,
	                                                      void** ppv);

	// IDispatch
	virtual HRESULT __stdcall GetTypeInfoCount(UINT* pCountTypeInfo);
	virtual HRESULT __stdcall GetTypeInfo(UINT iTypeInfo, LCID,
	                                      ITypeInfo** ppITypeInfo);
	virtual HRESULT __stdcall GetIDsOfNames(const IID& iid, OLECHAR** arrayNames,
	                                        UINT countNames, LCID,
	                                        DISPID* arrayDispIDs);
	virtual HRESULT __stdcall Invoke(DISPID dispidMember, const IID& iid, LCID,
	                                 WORD wFlags, DISPPARAMS* pDispParams,
	                                 VARIANT* pvarResult, EXCEPINFO* pExcepInfo,
	                                 UINT* pArgErr);

	// Interface IDualCoClassParamTest
	virtual HRESULT __stdcall GetItemAsCC(CoCoClassParamTestItem** result);
	virtual HRESULT __stdcall GetItemAsItf(ICoClassParamTestItem** result);

	// Initialization
	virtual HRESULT Init();

	// Constructor
	CCoClassParamTest(IUnknown* pUnknownOuter);

	// Destructor
	~CCoClassParamTest();

	// Pointer to type information.
	ITypeInfo* m_pITypeInfo;
};

#endif
