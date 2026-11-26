/*
 * CoCoClassParamTest.h - Component for coclass return type tests
 *
 * The type library declares the return type of IDualCoClassParamTest.GetItemAsCC
 * as the CoCoClassParamTestItem coclass.  comtypes generates
 * POINTER(POINTER(CoCoClassParamTestItem)) for the [out, retval] parameter
 * of the dual interface and POINTER(CoCoClassParamTestItem) for the return
 * type of the dispinterface.  A coclass is represented by its default
 * interface, so the method returns an object that implements
 * ICoClassParamTestItem.
 *
 * CCoClassParamTestItem is the object returned by
 * CCoClassParamTest::GetItemAsCC.  It implements ICoClassParamTestItem, the
 * default interface of the CoCoClassParamTestItem coclass.
 */

#ifndef __CoCoClassParamTest_h__
#define __CoCoClassParamTest_h__

#include "Iface.h"
#include "CUnknown.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispCoClassParamTest;

///////////////////////////////////////////////////////////
//
// CCoClassParamTestItem - the object returned by
// CCoClassParamTest::GetItemAsCC.  It implements ICoClassParamTestItem, the
// default interface of the CoCoClassParamTestItem coclass.
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
};

///////////////////////////////////////////////////////////
//
// CCoClassParamTest - the component that returns a coclass.
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
