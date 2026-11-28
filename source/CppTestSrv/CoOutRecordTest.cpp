/*
 * CoOutRecordTest.cpp - Component for [out] record ownership tests
 *
 * See CoOutRecordTest.h for the purpose of this test double.
 */

#include <objbase.h>
#include <string.h>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoOutRecordTest.h"

static inline void trace(const char* msg)
	{ Util::Trace("CoOutRecordTest", msg, S_OK); }
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoOutRecordTest", msg, hr); }

///////////////////////////////////////////////////////////
//
// CSentinel - Implementation
//

CSentinel::CSentinel(std::shared_ptr<LONG> liveCount)
	: m_cRef(1), m_liveCount(liveCount)
{
	::InterlockedIncrement(m_liveCount.get());
}

CSentinel::~CSentinel()
{
	::InterlockedDecrement(m_liveCount.get());
}

HRESULT __stdcall CSentinel::QueryInterface(const IID& iid, void** ppv)
{
	if (ppv == NULL)
	{
		return E_POINTER;
	}
	if (iid == IID_IUnknown)
	{
		*ppv = static_cast<IUnknown*>(this);
		AddRef();
		return S_OK;
	}
	*ppv = NULL;
	return E_NOINTERFACE;
}

ULONG __stdcall CSentinel::AddRef()
{
	return static_cast<ULONG>(::InterlockedIncrement(&m_cRef));
}

ULONG __stdcall CSentinel::Release()
{
	LONG cRef = ::InterlockedDecrement(&m_cRef);
	if (cRef == 0)
	{
		delete this;
	}
	return static_cast<ULONG>(cRef);
}

///////////////////////////////////////////////////////////
//
// COutRecordTest - Implementation
//

COutRecordTest::COutRecordTest(IUnknown* pUnknownOuter)
	: CUnknown(pUnknownOuter),
	  m_liveSentinelCount(std::make_shared<LONG>(0))
{
	// Empty
}

COutRecordTest::~COutRecordTest()
{
	trace("Destroy self.");
}

HRESULT COutRecordTest::CreateInstance(IUnknown* pUnknownOuter,
                                       CUnknown** ppNewComponent)
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION;
	}

	*ppNewComponent = new COutRecordTest(pUnknownOuter);
	return S_OK;
}

HRESULT __stdcall COutRecordTest::NondelegatingQueryInterface(const IID& iid,
                                                              void** ppv)
{
	if (iid == IID_IOutRecordTest)
	{
		return FinishQI(static_cast<IOutRecordTest*>(this), ppv);
	}
	return CUnknown::NondelegatingQueryInterface(iid, ppv);
}

HRESULT COutRecordTest::Init()
{
	// This interface is not a dual interface, so no type information is
	// needed.
	return S_OK;
}

IUnknown* COutRecordTest::CreateSentinel()
{
	CSentinel* pSentinel = new CSentinel(m_liveSentinelCount);
	// The sentinel is created with a reference count of 1.  That reference
	// is handed to the caller, so no AddRef is needed here.
	return static_cast<IUnknown*>(pSentinel);
}

///////////////////////////////////////////////////////////
//
// Interface IOutRecordTest - Implementation
//

HRESULT __stdcall COutRecordTest::GetRecord(LONG cchText,
                                            StructOutRecordTest* rec)
{
	trace("GetRecord called");
	if (rec == NULL)
	{
		return E_POINTER;
	}
	// Zero the out record first (COM out parameter rule).
	memset(rec, 0, sizeof(StructOutRecordTest));

	if (cchText < 0)
	{
		return E_INVALIDARG;
	}

	// Build a BSTR of cchText 'a' characters.
	BSTR text = ::SysAllocStringLen(NULL, static_cast<UINT>(cchText));
	if (text == NULL)
	{
		return E_OUTOFMEMORY;
	}
	for (LONG i = 0; i < cchText; ++i)
	{
		text[i] = L'a';
	}

	rec->tag = 1;
	rec->text = text;
	rec->unk = CreateSentinel();
	return S_OK;
}

HRESULT __stdcall COutRecordTest::GetRecordRetval(StructOutRecordTest* rec)
{
	trace("GetRecordRetval called");
	if (rec == NULL)
	{
		return E_POINTER;
	}
	memset(rec, 0, sizeof(StructOutRecordTest));

	BSTR text = ::SysAllocString(L"retval");
	if (text == NULL)
	{
		return E_OUTOFMEMORY;
	}

	rec->tag = 2;
	rec->text = text;
	rec->unk = CreateSentinel();
	return S_OK;
}

HRESULT __stdcall COutRecordTest::GetRecordsAround(StructOutRecordTest* first,
                                                   LONG* middle,
                                                   StructOutRecordTest* last)
{
	trace("GetRecordsAround called");
	if (first == NULL || middle == NULL || last == NULL)
	{
		return E_POINTER;
	}
	memset(first, 0, sizeof(StructOutRecordTest));
	memset(last, 0, sizeof(StructOutRecordTest));

	BSTR firstText = ::SysAllocString(L"first");
	if (firstText == NULL)
	{
		return E_OUTOFMEMORY;
	}
	BSTR lastText = ::SysAllocString(L"last");
	if (lastText == NULL)
	{
		::SysFreeString(firstText);
		return E_OUTOFMEMORY;
	}

	first->tag = 10;
	first->text = firstText;
	first->unk = CreateSentinel();

	*middle = 42;

	last->tag = 20;
	last->text = lastText;
	last->unk = CreateSentinel();
	return S_OK;
}

HRESULT __stdcall COutRecordTest::ModifyRecord(StructOutRecordTest* rec)
{
	trace("ModifyRecord called");
	if (rec == NULL)
	{
		return E_POINTER;
	}

	// Release the old contents, if any.
	if (rec->text != NULL)
	{
		::SysFreeString(rec->text);
		rec->text = NULL;
	}
	if (rec->unk != NULL)
	{
		rec->unk->Release();
		rec->unk = NULL;
	}

	BSTR text = ::SysAllocString(L"modified");
	if (text == NULL)
	{
		return E_OUTOFMEMORY;
	}

	rec->tag = rec->tag + 1;
	rec->text = text;
	rec->unk = CreateSentinel();
	return S_OK;
}

HRESULT __stdcall COutRecordTest::FailGetRecord(StructOutRecordTest* rec)
{
	trace("FailGetRecord called");
	// Deliberately do not touch *rec.
	return E_FAIL;
}

HRESULT __stdcall COutRecordTest::GetLiveSentinelCount(LONG* count)
{
	trace("GetLiveSentinelCount called");
	if (count == NULL)
	{
		return E_POINTER;
	}
	*count = *m_liveSentinelCount;
	return S_OK;
}

HRESULT __stdcall COutRecordTest::GetSentinel(IUnknown** sentinel)
{
	trace("GetSentinel called");
	if (sentinel == NULL)
	{
		return E_POINTER;
	}
	// CreateSentinel returns a reference held for the caller.
	*sentinel = CreateSentinel();
	return S_OK;
}
