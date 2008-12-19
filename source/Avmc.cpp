// Avmc.cpp : Implementation of CAvmcIfcApp and DLL registration.

#include "stdafx.h"
#include "AvmcIfc.h"
#include "Avmc.h"
#include <comdef.h>

/////////////////////////////////////////////////////////////////////////////
//

const IID DEVICE_INFO_IID = {0x6C7A25CB, 0x7938, 0x4BE0, {0xA2, 0x85, 0x12, 0xC6, 0x16, 0x71, 0x7F, 0xDD} };

STDMETHODIMP Avmc::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IAvmc,
	};

	for (int i=0;i<sizeof(arr)/sizeof(arr[0]);i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP Avmc::FindAllAvmc(SAFEARRAY **avmcList)
{	
	FT_STATUS ftStatus;
	DWORD numDevs;
	DWORD i;

	if (!avmcList)
		return E_POINTER;

	if (*avmcList != NULL) {
		::SafeArrayDestroy(*avmcList);
		*avmcList = NULL;
	}
	numDevs = 2;
	ftStatus = FT_OK;

	devInfo = new FT_DEVICE_LIST_INFO_NODE[numDevs];

	strcpy(devInfo[0].Description, "Avmc");
	devInfo[0].Flags = 12;
	devInfo[0].ID = 13;
	devInfo[0].LocId = 14;
	
	strcpy(devInfo[0].SerialNumber, "1234");
	devInfo[0].Type = 15;

	strcpy(devInfo[1].Description, "Avmc2");
	devInfo[1].Flags = 22;
	devInfo[1].ID = 23;
	devInfo[1].LocId = 24;
	strcpy(devInfo[1].SerialNumber, "5678");
	devInfo[1].Type = 25;

	ftStatus = FT_OK;
	if (ftStatus == FT_OK) {

		// Copy into a new safe array...
		//////////////////////////////////////////////////
		//here starts the actual creation of the array
		//////////////////////////////////////////////////
		IRecordInfo *pUdtRecordInfo = NULL;
		HRESULT hr = GetRecordInfoFromGuids (LIBID_AVMCIFCLib, 1, 0, 0, DEVICE_INFO_IID, &pUdtRecordInfo);
	    if( FAILED( hr ) ) {
			HRESULT hr2 = Error( _T("Can not create Device Info interface") );
			return( hr );
		}
		////////////////// SafeArray Creation and Returning ///////////////
		SAFEARRAYBOUND rgsabound[1];
		rgsabound[0].lLbound = 0;
		rgsabound[0].cElements = numDevs;
		*avmcList = ::SafeArrayCreateEx( VT_RECORD, 1, rgsabound, pUdtRecordInfo );

		pUdtRecordInfo->Release(); // Release the interface
		if( *avmcList == NULL ) {
	        HRESULT hr = Error( _T("Can not create array of Device Info structures") );
			return( hr );
		}

		///////////////////////////////////////////////////////////////////

#ifdef DEBUG_NOW
		strstream s2;
		s2.clear();
		for (i = 0; i < numDevs; i++) {
			s2 << "Dev " << i << endl;
			s2 << " Flags = 0x" << hex << devInfo[i].Flags << endl;
			s2 << " Type =  0x" << hex << devInfo[i].Type << endl;
			s2 << " ID =    0x" << hex << devInfo[i].ID << endl;
			s2 << " LocId = 0x" << hex << devInfo[i].LocId << endl;
			s2 << " SerialNumber = " << devInfo[i].SerialNumber << endl;
			s2 << " Description  = " << devInfo[i].Description << endl;
			s2 << " ftHandle     = 0x" << hex << devInfo[i].ftHandle << endl;
			s2 << "---" << endl;
		}
		s2 << ends;
		MessageBox (0, s2.str(), "Device List", 0);
#endif
		DeviceInfo HUGEP *pD = NULL;
		hr = SafeArrayAccessData (*avmcList, (void HUGEP **)&pD);

		for (i = 0; i < numDevs; i++) {
			pD[i].Flags =			(ULONG) devInfo[i].Flags;
			pD[i].Type =			(ULONG) devInfo[i].Type;
			pD[i].ID =				(ULONG) devInfo[i].ID;
			pD[i].LocId =			(ULONG) devInfo[i].LocId;
			pD[i].SerialNumber =	_com_util::ConvertStringToBSTR(devInfo[i].SerialNumber);
			pD[i].Description =		_com_util::ConvertStringToBSTR(devInfo[i].Description);
			pD[i].ftHandle =		(ULONG)devInfo[i].ftHandle;
		}

		hr = SafeArrayUnaccessData(*avmcList);
	}

	return S_OK;
}