import ctypes
from ctypes import HRESULT, POINTER, byref
from typing import TYPE_CHECKING, Any, Optional, Type

import comtypes
import comtypes.client
import comtypes.client.dynamic
from comtypes import GUID, STDMETHOD, IUnknown
from comtypes import RevokeActiveObject as RevokeActiveObject
from comtypes.automation import IDispatch

if TYPE_CHECKING:
    from ctypes import _Pointer

    from comtypes import hints  # type: ignore


################################################################
# Interfaces
class IClassFactory(IUnknown):
    _iid_ = GUID("{00000001-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(
            HRESULT,
            "CreateInstance",
            [POINTER(IUnknown), POINTER(GUID), POINTER(ctypes.c_void_p)],
        ),
        STDMETHOD(HRESULT, "LockServer", [ctypes.c_int]),
    ]

    def CreateInstance(
        self,
        punkouter: Optional[Type["_Pointer[IUnknown]"]] = None,
        interface: Optional[Type[IUnknown]] = None,
        dynamic: bool = False,
    ) -> Any:
        if dynamic:
            if interface is not None:
                raise ValueError("interface and dynamic are mutually exclusive")
            itf = IDispatch
        elif interface is None:
            itf = IUnknown
        else:
            itf = interface
        obj = POINTER(itf)()
        self.__com_CreateInstance(punkouter, itf._iid_, byref(obj))  # type: ignore
        if dynamic:
            return comtypes.client.dynamic.Dispatch(obj)
        elif interface is None:
            # An interface was not specified, so return the best.
            return comtypes.client.GetBestInterface(obj)
        # An interface was specified and obj is already that interface.
        return obj

    if TYPE_CHECKING:

        def LockServer(self, fLock: int) -> hints.Hresult: ...


# class IExternalConnection(IUnknown):
#     _iid_ = GUID("{00000019-0000-0000-C000-000000000046}")
#     _methods_ = [
#         STDMETHOD(HRESULT, "AddConnection", [c_ulong, c_ulong]),
#         STDMETHOD(HRESULT, "ReleaseConnection", [c_ulong, c_ulong, c_ulong])]


def RegisterActiveObject(comobj: comtypes.COMObject, weak: bool = True) -> int:
    """Registers a pointer as the active object for its class and returns the handle."""
    punk = comobj._com_pointers_[IUnknown._iid_]
    clsid = comobj._reg_clsid_
    flags = comtypes.ACTIVEOBJECT_WEAK if weak else comtypes.ACTIVEOBJECT_STRONG
    return comtypes.RegisterActiveObject(punk, clsid, flags)
