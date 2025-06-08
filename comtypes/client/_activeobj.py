from typing import Any, Optional, Type, TypeVar, overload
from typing import Union as _UnionT

import comtypes
import comtypes.client.dynamic
from comtypes import GUID, CoClass, IUnknown, automation
from comtypes.client._managing import _manage

_T_IUnknown = TypeVar("_T_IUnknown", bound=IUnknown)


################################################################
#
# Object creation
#
@overload
def GetActiveObject(progid: _UnionT[str, Type[CoClass], GUID]) -> Any: ...
@overload
def GetActiveObject(
    progid: _UnionT[str, Type[CoClass], GUID], interface: Type[_T_IUnknown]
) -> _T_IUnknown: ...
def GetActiveObject(
    progid: _UnionT[str, Type[CoClass], GUID],
    interface: Optional[Type[IUnknown]] = None,
    dynamic: bool = False,
) -> Any:
    """Return a pointer to a running COM object that has been
    registered with COM.

    'progid' may be a string like "Excel.Application",
       a string specifying a clsid, a GUID instance, or an object with
       a _clsid_ attribute which should be any of the above.
    'interface' allows to force a certain interface.
    'dynamic=True' will return a dynamic dispatch object.
    """
    clsid = GUID.from_progid(progid)
    if dynamic:
        if interface is not None:
            raise ValueError("interface and dynamic are mutually exclusive")
        interface = automation.IDispatch
    elif interface is None:
        interface = getattr(progid, "_com_interfaces_", [None])[0]
    obj = comtypes.GetActiveObject(clsid, interface=interface)
    if dynamic:
        return comtypes.client.dynamic.Dispatch(obj)
    return _manage(obj, clsid, interface=interface)


def RegisterActiveObject(
    punk: IUnknown, progid: _UnionT[str, Type[CoClass], GUID], weak: bool = True
) -> int:
    clsid = GUID.from_progid(progid)
    flags = comtypes.ACTIVEOBJECT_WEAK if weak else comtypes.ACTIVEOBJECT_STRONG
    return comtypes.RegisterActiveObject(punk, clsid, flags)
