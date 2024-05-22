/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

#include "CFactory.h"
#include "Iface.h"
#include "Cmpnt.h"


///////////////////////////////////////////////////////////
//
// Server.cpp
//
// This file contains the component server code.
// The FactoryDataArray contains the components that 
// can be served.
//

// Each component derived from CUnknown defines a static function
// for creating the component with the following prototype. 
// HRESULT CreateInstance(IUnknown* pUnknownOuter, 
//                        CUnknown** ppNewComponent) ;
// This function is used to create the component.
//

//
// The following array contains the data used by CFactory
// to create components. Each element in the array contains
// the CLSID, the pointer to the creation function, and the name
// of the component to place in the Registry.
//
CFactoryData g_FactoryDataArray[] =
{
	{&CLSID_Component, CA::CreateInstance, 
		L"Comtypes Test automation server component",	// Friendly Name
		L"ComtypesTest.COM.Server.1",									// ProgID
		L"ComtypesTest.COM.Server",										// Version-independent ProgID
		&LIBID_ComtypesTestLib,												// Type Library ID
		NULL, 0}
} ;
int g_cFactoryDataEntries
	= sizeof(g_FactoryDataArray) / sizeof(CFactoryData) ;
