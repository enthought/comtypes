// TestComServer.cpp : Implementation of CTestComServer

#include "stdafx.h"
#include "TestComServer.h"


// CTestComServer

STDMETHODIMP CTestComServer::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ITestComServer
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
