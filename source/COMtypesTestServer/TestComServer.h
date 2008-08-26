// TestComServer.h : Declaration of the CTestComServer

#pragma once
#include "resource.h"       // main symbols

#include "COMtypesTestServer.h"
#include "_ITestComServerEvents_CP.h"


// CTestComServer

class ATL_NO_VTABLE CTestComServer : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTestComServer, &CLSID_TestComServer>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CTestComServer>,
	public CProxy_ITestComServerEvents<CTestComServer>, 
	public IDispatchImpl<ITestComServer, &IID_ITestComServer, &LIBID_COMtypesTestServerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CTestComServer()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_TESTCOMSERVER)

DECLARE_NOT_AGGREGATABLE(CTestComServer)

BEGIN_COM_MAP(CTestComServer)
	COM_INTERFACE_ENTRY(ITestComServer)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CTestComServer)
	CONNECTION_POINT_ENTRY(__uuidof(_ITestComServerEvents))
END_CONNECTION_POINT_MAP()
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:

};

OBJECT_ENTRY_AUTO(__uuidof(TestComServer), CTestComServer)
