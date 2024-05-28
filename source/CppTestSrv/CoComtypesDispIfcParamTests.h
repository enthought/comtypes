/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// CoComtypesDispIfcParamTests.cpp - Component
//

#include "Iface.h"
#include "CUnknown.h" 

///////////////////////////////////////////////////////////
//
// Component A
//
class CA : public CUnknown,
           public IDualRecordParamTest
{
public:	
	// Creation
	static HRESULT CreateInstance(IUnknown* pUnknownOuter,
	                              CUnknown** ppNewComponent ) ;

private:
	// Declare the delegating IUnknown.
	DECLARE_IUNKNOWN

	// IUnknown
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid,
	                                                      void** ppv) ;

	// IDispatch
	virtual HRESULT __stdcall GetTypeInfoCount(UINT* pCountTypeInfo) ;

	virtual HRESULT __stdcall GetTypeInfo(
		UINT iTypeInfo,
		LCID,              // Localization is not supported.
		ITypeInfo** ppITypeInfo) ;
	
	virtual HRESULT __stdcall GetIDsOfNames(
		const IID& iid,
		OLECHAR** arrayNames,
		UINT countNames,
		LCID,              // Localization is not supported.
		DISPID* arrayDispIDs) ;

	virtual HRESULT __stdcall Invoke(   
		DISPID dispidMember,
		const IID& iid,
		LCID,              // Localization is not supported.
		WORD wFlags,
		DISPPARAMS* pDispParams,
		VARIANT* pvarResult,
		EXCEPINFO* pExcepInfo,
		UINT* pArgErr) ;

	// Interface IDualRecordParamTest
	virtual HRESULT __stdcall InitRecord(StructRecordParamTest* test_record) ;
	virtual HRESULT __stdcall VerifyRecord(
										 StructRecordParamTest* test_record,
										 VARIANT_BOOL* result) ;

	// Initialization
 	virtual HRESULT Init() ;

	// Constructor
	CA(IUnknown* pUnknownOuter) ;

	// Destructor
	~CA() ;

	// Pointer to type information.
	ITypeInfo* m_pITypeInfo ;
} ;
