/*
 * CoNamedPropertyTest.h - Component for parameterized (named) property tests
 *
 * This component implements a dual interface (IDualNamedPropertyTest) whose
 * properties take arguments.  comtypes exposes such properties through its
 * `named_property` descriptor, which supports `obj.Prop[index]`,
 * `obj.Prop[index] = value`, `obj.Prop(index)` and tuple indexing
 * `obj.Prop[a, b]`.
 *
 * The getters return a VARIANT so that reading an index that has not been
 * assigned can report DISP_E_BADINDEX through the VARIANT's scode.
 *
 * Range takes two integer indices that denote a range.  Reading
 * Range[start, end] returns one element per index in [start, end], filling an
 * index that has not been assigned with a VT_NULL VARIANT.  Writing
 * Range[start, end] = values assigns values[i] to index start + i.
 */

#ifndef __CoNamedPropertyTest_h__
#define __CoNamedPropertyTest_h__

#include <map>

#include "Iface.h"
#include "CUnknown.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispNamedPropertyTest;

///////////////////////////////////////////////////////////
//
// Component for parameterized (named) property tests
//
class CNamedPropertyTest : public CUnknown,
                           public IDualNamedPropertyTest
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

	// Interface IDualNamedPropertyTest
	virtual HRESULT __stdcall get_Item(LONG index, VARIANT* value);
	virtual HRESULT __stdcall put_Item(LONG index, LONG value);
	virtual HRESULT __stdcall get_Range(LONG start, LONG end, VARIANT* value);
	virtual HRESULT __stdcall put_Range(LONG start, LONG end, VARIANT value);

	// Initialization
	virtual HRESULT Init();

	// Constructor
	CNamedPropertyTest(IUnknown* pUnknownOuter);

	// Destructor
	~CNamedPropertyTest();

	// Pointer to type information.
	ITypeInfo* m_pITypeInfo;

	// Storage for the single-index Item property.
	std::map<LONG, LONG> m_itemValues;
	// Storage for the two-index Range property.
	std::map<LONG, LONG> m_rangeValues;
};

#endif
