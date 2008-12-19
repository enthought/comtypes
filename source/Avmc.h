// Avmc.h: Definition of the Avmc class
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_AVMC_H__14DBEE48_EB7D_48AD_8BEE_8E15B2D0A780__INCLUDED_)
#define AFX_AVMC_H__14DBEE48_EB7D_48AD_8BEE_8E15B2D0A780__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// Avmc

//#include "FTD2XX.h"
typedef PVOID	FT_HANDLE;
typedef ULONG	FT_STATUS;

enum {
	FT_OK,
	FT_INVALID_HANDLE,
	FT_DEVICE_NOT_FOUND,
	FT_DEVICE_NOT_OPENED,
	FT_IO_ERROR,
	FT_INSUFFICIENT_RESOURCES,
	FT_INVALID_PARAMETER,
	FT_INVALID_BAUD_RATE,

	FT_DEVICE_NOT_OPENED_FOR_ERASE,
	FT_DEVICE_NOT_OPENED_FOR_WRITE,
	FT_FAILED_TO_WRITE_DEVICE,
	FT_EEPROM_READ_FAILED,
	FT_EEPROM_WRITE_FAILED,
	FT_EEPROM_ERASE_FAILED,
	FT_EEPROM_NOT_PRESENT,
	FT_EEPROM_NOT_PROGRAMMED,
	FT_INVALID_ARGS,
	FT_NOT_SUPPORTED,
	FT_OTHER_ERROR
};


//#define FT_SUCCESS(status) ((status) == FT_OK)

typedef struct _ft_device_list_info_node {
	ULONG Flags;
	ULONG Type;
	ULONG ID;
	DWORD LocId;
	char SerialNumber[16];
	char Description[64];
	FT_HANDLE ftHandle;
} FT_DEVICE_LIST_INFO_NODE;

class Avmc : 
	public IDispatchImpl<IAvmc, &IID_IAvmc, &LIBID_AVMCIFCLib>, 
	public ISupportErrorInfo,
	public CComObjectRoot,
	public CComCoClass<Avmc,&CLSID_Avmc>
{
public:
	Avmc() {}
BEGIN_COM_MAP(Avmc)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IAvmc)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()
//DECLARE_NOT_AGGREGATABLE(Avmc) 
// Remove the comment from the line above if you don't want your object to 
// support aggregation. 

DECLARE_REGISTRY_RESOURCEID(IDR_Avmc)
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IAvmc
public:
	STDMETHOD(FindAllAvmc)(/*[out]*/ SAFEARRAY **avmcList);

private:
	void FinalizeCommand(char *command, int cmdLength);
	int  ReadToSimpleArray(int devNum, char *arr);

private:
	FT_DEVICE_LIST_INFO_NODE			*devInfo;
	char								mResArray[512]; // Inconsistant array - just to read results...
};

#endif // !defined(AFX_AVMC_H__14DBEE48_EB7D_48AD_8BEE_8E15B2D0A780__INCLUDED_)
