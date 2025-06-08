from ctypes import HRESULT, POINTER, OleDLL, byref
from ctypes.wintypes import LPVOID
from typing import Optional, Type, TypeVar, overload

from comtypes import GUID
from comtypes._post_coinit.unknwn import IUnknown
from comtypes.GUID import REFCLSID

_T_IUnknown = TypeVar("_T_IUnknown", bound=IUnknown)


@overload
def GetActiveObject(clsid: GUID, interface: None = None) -> IUnknown: ...
@overload
def GetActiveObject(clsid: GUID, interface: Type[_T_IUnknown]) -> _T_IUnknown: ...
def GetActiveObject(
    clsid: GUID, interface: Optional[Type[IUnknown]] = None
) -> IUnknown:
    """Retrieves a pointer to a running object"""
    p = POINTER(IUnknown)()
    _GetActiveObject(byref(clsid), None, byref(p))
    if interface is not None:
        p = p.QueryInterface(interface)  # type: ignore
    return p  # type: ignore


_oleaut32 = OleDLL("oleaut32")

_GetActiveObject = _oleaut32.GetActiveObject
_GetActiveObject.argtypes = [REFCLSID, LPVOID, POINTER(POINTER(IUnknown))]
_GetActiveObject.restype = HRESULT
