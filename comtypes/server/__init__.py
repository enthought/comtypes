import ctypes
from ctypes import HRESULT, POINTER, byref
from typing import TYPE_CHECKING

import comtypes
import comtypes.client
from comtypes import GUID, STDMETHOD, IUnknown

if TYPE_CHECKING:
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

    def CreateInstance(self, punkouter=None, interface=None, dynamic=False):
        if dynamic:
            if interface is not None:
                raise ValueError("interface and dynamic are mutually exclusive")
            realInterface = comtypes.automation.IDispatch
        elif interface is None:
            realInterface = IUnknown
        else:
            realInterface = interface
        obj = POINTER(realInterface)()
        self.__com_CreateInstance(punkouter, realInterface._iid_, byref(obj))
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

# The following code is untested:

ACTIVEOBJECT_STRONG = 0x0
ACTIVEOBJECT_WEAK = 0x1

oleaut32 = ctypes.oledll.oleaut32


def RegisterActiveObject(comobj: comtypes.COMObject, weak: bool = True) -> int:
    punk = comobj._com_pointers_[IUnknown._iid_]
    clsid = comobj._reg_clsid_
    if weak:
        flags = ACTIVEOBJECT_WEAK
    else:
        flags = ACTIVEOBJECT_STRONG
    handle = ctypes.c_ulong()
    oleaut32.RegisterActiveObject(punk, byref(clsid), flags, byref(handle))
    return handle.value


def RevokeActiveObject(handle: int) -> None:
    oleaut32.RevokeActiveObject(handle, None)
