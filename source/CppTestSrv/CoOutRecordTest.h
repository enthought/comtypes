/*
 * CoOutRecordTest.h - Component for [out] record ownership tests
 *
 */

#ifndef __CoOutRecordTest_h__
#define __CoOutRecordTest_h__

#include <memory>

#include "Iface.h"
#include "CUnknown.h"

///////////////////////////////////////////////////////////
//
// Sentinel object
//
// A minimal IUnknown implementation.  It is not an ATL CComObject; it is a
// plain C++ object so that its lifetime is fully under this test double's
// control.  Each sentinel holds a copy of the shared live-count so that the
// count is decremented exactly when the sentinel is destroyed.
//
class CSentinel : public IUnknown
{
public:
	explicit CSentinel(std::shared_ptr<LONG> liveCount);
	virtual ~CSentinel();

	// IUnknown
	virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
	virtual ULONG __stdcall AddRef();
	virtual ULONG __stdcall Release();

private:
	LONG m_cRef;
	std::shared_ptr<LONG> m_liveCount;
};

///////////////////////////////////////////////////////////
//
// Component for [out] record ownership tests
//
class COutRecordTest : public CUnknown,
                       public IOutRecordTest
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

	// Interface IOutRecordTest
	virtual HRESULT __stdcall GetRecord(LONG cchText, StructOutRecordTest* rec);
	virtual HRESULT __stdcall GetRecordRetval(StructOutRecordTest* rec);
	virtual HRESULT __stdcall GetRecordsAround(StructOutRecordTest* first,
	                                           LONG* middle,
	                                           StructOutRecordTest* last);
	virtual HRESULT __stdcall ModifyRecord(StructOutRecordTest* rec);
	virtual HRESULT __stdcall FailGetRecord(StructOutRecordTest* rec);
	virtual HRESULT __stdcall GetLiveSentinelCount(LONG* count);
	virtual HRESULT __stdcall GetSentinel(IUnknown** sentinel);

	// Initialization
	virtual HRESULT Init();

	// Constructor
	COutRecordTest(IUnknown* pUnknownOuter);

	// Destructor
	~COutRecordTest();

	// Create a new sentinel and return it with ownership transferred to the
	// caller (reference count 1).  The caller is responsible for releasing it.
	IUnknown* CreateSentinel();

	// Per-instance count of live sentinels.
	std::shared_ptr<LONG> m_liveSentinelCount;
};

#endif
