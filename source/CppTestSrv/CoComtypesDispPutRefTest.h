/*
 * CoComtypesDispPutRefTest.h - Component for dispinterface propputref tests
 *
 * This component implements a dual interface (IDualPutRefTest) whose member
 * exercises a code path in comtypes' dispinterface member generator that
 * parameterized properties do not reach:
 *
 *   * Item is a propputref property.  comtypes invokes its setter with
 *     DISPATCH_PROPERTYPUTREF.
 *
 * The dispinterface IDispPutRefTest is an alias for the dual interface, so
 * the same members are reached through IDispatch::Invoke.
 */

#ifndef __CoComtypesDispPutRefTest_h__
#define __CoComtypesDispPutRefTest_h__

#include <map>

#include "Iface.h"
#include "CUnknown.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispPutRefTest;

///////////////////////////////////////////////////////////
//
// Component for dispinterface propputref tests
//
class CComtypesDispPutRefTest : public CUnknown,
                                public IDualPutRefTest
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

	// Interface IDualPutRefTest
	virtual HRESULT __stdcall get_Item(LONG index, IUnknown** value);
	virtual HRESULT __stdcall putref_Item(LONG index, IUnknown* value);

	// Initialization
	virtual HRESULT Init();

	// Constructor
	CComtypesDispPutRefTest(IUnknown* pUnknownOuter);

	// Destructor
	~CComtypesDispPutRefTest();

	// Pointer to type information.
	ITypeInfo* m_pITypeInfo;

	// Pointers stored by putref_RefItem, keyed by index.
	std::map<LONG, IUnknown*> m_refs;
};

#endif
