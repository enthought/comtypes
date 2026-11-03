#include "Iface.h"
#include "CUnknown.h"

class CStringSafearrayInsideVariantTest : public CUnknown,
                           public IStringSafearrayInsideVariantTest
{
public:
	// Creation
	static HRESULT CreateInstance(IUnknown* pUnknownOuter,
	                              CUnknown** ppNewComponent) ;

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

	// Interface IStringSafearrayInsideVariantTest
	virtual HRESULT __stdcall VariantArrayToBstrArray(VARIANT arr, VARIANT* result) ;
	virtual HRESULT __stdcall BstrArrayToVariantArray(VARIANT arr, VARIANT* result) ;
	virtual HRESULT __stdcall RepeatVariantArray(VARIANT arr, LONG count, VARIANT* result) ;
	virtual HRESULT __stdcall RepeatBstrArray(VARIANT arr, LONG count, VARIANT* result) ;

	// Initialization
	virtual HRESULT Init() ;

	// Constructor
	CStringSafearrayInsideVariantTest(IUnknown* pUnknownOuter) ;

	// Destructor
	~CStringSafearrayInsideVariantTest() ;

	// Pointer to type information.
	ITypeInfo* m_pITypeInfo ;
} ;
