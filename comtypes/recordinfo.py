import ctypes
from ctypes import (
    POINTER,
    OleDLL,
    byref,
    c_void_p,
    c_wchar_p,
)
from ctypes.wintypes import BOOL, LCID, ULONG
from typing import TYPE_CHECKING, Any, Optional

from comtypes import BSTR, COMMETHOD, GUID, HRESULT, IUnknown, _safearray
from comtypes.automation import VARIANT
from comtypes.typeinfo import ITypeInfo

if TYPE_CHECKING:
    from ctypes import _Pointer

    from comtypes import hints  # type: ignore


class IRecordInfo(IUnknown):
    # C:/vc98/include/OAIDL.H 5974
    _iid_ = GUID("{0000002F-0000-0000-C000-000000000046}")

    def GetFieldNames(self, *args: Any) -> list[Optional[str]]:
        count = ctypes.c_ulong()
        self.__com_GetFieldNames(count, None)  # type: ignore
        array = (BSTR * count.value)()
        self.__com_GetFieldNames(count, array)  # type: ignore
        result = array[:]
        # XXX Should SysFreeString the array contents. How to?
        return result

    if TYPE_CHECKING:
        # def RecordInit
        # def RecordClear
        def RecordCopy(
            self, pvExisting: hints.Incomplete, pvNew: hints.Incomplete
        ) -> hints.Hresult: ...
        def GetGuid(self) -> GUID: ...
        def GetName(self) -> str: ...
        def GetSize(self) -> int: ...
        def GetTypeInfo(self) -> ITypeInfo: ...
        # def GetField
        # def GetFieldNoCopy
        # def PutField
        # def PutFieldNoCopy
        def IsMatchingType(self, value: "IRecordInfo") -> bool: ...
        # def RecordCreate
        def RecordCreateCopy(self, pvSource: hints.Incomplete) -> int: ...
        def RecordDestroy(self, pvRecord: hints.Incomplete) -> hints.Hresult: ...


IRecordInfo._methods_ = [
    COMMETHOD([], HRESULT, "RecordInit", (["in"], c_void_p, "pvNew")),
    COMMETHOD([], HRESULT, "RecordClear", (["in"], c_void_p, "pvExisting")),
    COMMETHOD(
        [],
        HRESULT,
        "RecordCopy",
        (["in"], c_void_p, "pvExisting"),
        (["in"], c_void_p, "pvNew"),
    ),
    COMMETHOD([], HRESULT, "GetGuid", (["out"], POINTER(GUID), "pguid")),
    COMMETHOD([], HRESULT, "GetName", (["out"], POINTER(BSTR), "pbstrName")),
    COMMETHOD([], HRESULT, "GetSize", (["out"], POINTER(ctypes.c_ulong), "pcbSize")),
    COMMETHOD(
        [], HRESULT, "GetTypeInfo", (["out"], POINTER(POINTER(ITypeInfo)), "ppTypeInfo")
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetField",
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["out"], POINTER(VARIANT), "pvarField"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetFieldNoCopy",
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["out"], POINTER(VARIANT), "pvarField"),
        (["out"], POINTER(c_void_p), "ppvDataCArray"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "PutField",
        (["in"], ctypes.c_ulong, "wFlags"),
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["in"], POINTER(VARIANT), "pvarField"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "PutFieldNoCopy",
        (["in"], ctypes.c_ulong, "wFlags"),
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["in"], POINTER(VARIANT), "pvarField"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetFieldNames",
        (["in", "out"], POINTER(ctypes.c_ulong), "pcNames"),
        (["in"], POINTER(BSTR), "rgBstrNames"),
    ),
    COMMETHOD([], BOOL, "IsMatchingType", (["in"], POINTER(IRecordInfo))),
    COMMETHOD([], HRESULT, "RecordCreate"),
    COMMETHOD(
        [],
        HRESULT,
        "RecordCreateCopy",
        (["in"], c_void_p, "pvSource"),
        (["out"], POINTER(c_void_p), "ppvDest"),
    ),
    COMMETHOD([], HRESULT, "RecordDestroy", (["in"], c_void_p, "pvRecord")),
]


################################################################
# functions
_oleaut32 = OleDLL("oleaut32")

_GetRecordInfoFromTypeInfo = _oleaut32.GetRecordInfoFromTypeInfo
_GetRecordInfoFromTypeInfo.argtypes = [
    POINTER(ITypeInfo),
    POINTER(POINTER(IRecordInfo)),
]
_GetRecordInfoFromTypeInfo.restype = HRESULT

_GetRecordInfoFromGuids = _oleaut32.GetRecordInfoFromGuids
_GetRecordInfoFromGuids.argtypes = [
    POINTER(GUID),
    ULONG,
    ULONG,
    LCID,
    POINTER(GUID),
    POINTER(POINTER(IRecordInfo)),
]
_GetRecordInfoFromGuids.restype = HRESULT


def GetRecordInfoFromTypeInfo(tinfo: ITypeInfo) -> IRecordInfo:
    """Return an IRecordInfo pointer to the UDT described in tinfo"""
    ri = POINTER(IRecordInfo)()
    _GetRecordInfoFromTypeInfo(tinfo, byref(ri))
    return ri  # type: ignore


def GetRecordInfoFromGuids(
    rGuidTypeLib: str, verMajor: int, verMinor: int, lcid: int, rGuidTypeInfo: str
) -> IRecordInfo:
    ri = POINTER(IRecordInfo)()
    _GetRecordInfoFromGuids(
        byref(GUID(rGuidTypeLib)),
        verMajor,
        verMinor,
        lcid,
        byref(GUID(rGuidTypeInfo)),
        byref(ri),
    )
    return ri  # type: ignore


_SafeArrayGetRecordInfo = _oleaut32.SafeArrayGetRecordInfo
_SafeArrayGetRecordInfo.restype = HRESULT
_SafeArrayGetRecordInfo.argtypes = [POINTER(_safearray.SAFEARRAY), POINTER(c_void_p)]


def SafeArrayGetRecordInfo(
    pa: "_Pointer[_safearray.tagSAFEARRAY]",
) -> IRecordInfo:
    result = POINTER(IRecordInfo)()
    _SafeArrayGetRecordInfo(pa, byref(result))
    return result  # type: ignore
