import logging
from typing import TYPE_CHECKING, Any, Optional, Type, TypeVar, overload
from typing import Union as _UnionT

import comtypes
import comtypes.client.dynamic
from comtypes import COSERVERINFO, GUID, CoClass, IUnknown, automation
from comtypes.client._managing import _manage

if TYPE_CHECKING:
    from comtypes import hints  # type: ignore


_T_IUnknown = TypeVar("_T_IUnknown", bound=IUnknown)
logger = logging.getLogger(__name__)


################################################################
#
# Object creation
#
if TYPE_CHECKING:

    @overload
    def GetClassObject(
        progid: _UnionT[str, Type[CoClass], GUID],
        clsctx: Optional[int] = None,
        pServerInfo: Optional[COSERVERINFO] = None,
        interface: None = None,
    ) -> hints.IClassFactory: ...
    @overload
    def GetClassObject(
        progid: _UnionT[str, Type[CoClass], GUID],
        clsctx: Optional[int] = None,
        pServerInfo: Optional[COSERVERINFO] = None,
        interface: Type[_T_IUnknown] = hints.IClassFactory,
    ) -> _T_IUnknown: ...


def GetClassObject(progid, clsctx=None, pServerInfo=None, interface=None):
    # type: (_UnionT[str, Type[CoClass], GUID], Optional[int], Optional[COSERVERINFO], Optional[Type[IUnknown]]) -> IUnknown
    """Create and return the class factory for a COM object.

    'clsctx' specifies how to create the object, use the CLSCTX_... constants.
    'pServerInfo', if used, must be a pointer to a comtypes.COSERVERINFO instance
    'interface' may be used to request an interface other than IClassFactory
    """
    clsid = GUID.from_progid(progid)
    return comtypes.CoGetClassObject(clsid, clsctx, pServerInfo, interface)


@overload
def CreateObject(progid: _UnionT[str, Type[CoClass], GUID]) -> Any: ...
@overload
def CreateObject(
    progid: _UnionT[str, Type[CoClass], GUID],
    clsctx: Optional[int] = None,
    machine: Optional[str] = None,
    interface: Optional[Type[_T_IUnknown]] = None,
    dynamic: bool = ...,
    pServerInfo: Optional[COSERVERINFO] = None,
) -> _T_IUnknown: ...
def CreateObject(
    progid: _UnionT[str, Type[CoClass], GUID],  # which object to create
    clsctx: Optional[int] = None,  # how to create the object
    machine: Optional[str] = None,  # where to create the object
    interface: Optional[Type[IUnknown]] = None,  # the interface we want
    dynamic: bool = False,  # use dynamic dispatch
    pServerInfo: Optional[COSERVERINFO] = None,  # server info struct for remoting
) -> Any:
    """Create a COM object from 'progid', and try to QueryInterface()
    it to the most useful interface, generating typelib support on
    demand.  A pointer to this interface is returned.

    'progid' may be a string like "InternetExplorer.Application",
       a string specifying a clsid, a GUID instance, or an object with
       a _clsid_ attribute which should be any of the above.
    'clsctx' specifies how to create the object, use the CLSCTX_... constants.
    'machine' allows to specify a remote machine to create the object on.
    'interface' allows to force a certain interface
    'dynamic=True' will return a dynamic dispatch object
    'pServerInfo', if used, must be a pointer to a comtypes.COSERVERINFO instance
        This supercedes 'machine'.

    You can also later request to receive events with GetEvents().
    """
    clsid = GUID.from_progid(progid)
    logger.debug("%s -> %s", progid, clsid)
    if dynamic:
        if interface:
            raise ValueError("interface and dynamic are mutually exclusive")
        interface = automation.IDispatch
    elif interface is None:
        interface = getattr(progid, "_com_interfaces_", [None])[0]
    if machine is None and pServerInfo is None:
        logger.debug(
            "CoCreateInstance(%s, clsctx=%s, interface=%s)", clsid, clsctx, interface
        )
        obj = comtypes.CoCreateInstance(clsid, clsctx=clsctx, interface=interface)
    else:
        logger.debug(
            "CoCreateInstanceEx(%s, clsctx=%s, interface=%s, machine=%s,\
                        pServerInfo=%s)",
            clsid,
            clsctx,
            interface,
            machine,
            pServerInfo,
        )
        if machine is not None and pServerInfo is not None:
            msg = "You cannot set both the machine name and server info."
            raise ValueError(msg)
        obj = comtypes.CoCreateInstanceEx(
            clsid,
            clsctx=clsctx,
            interface=interface,
            machine=machine,
            pServerInfo=pServerInfo,
        )
    if dynamic:
        return comtypes.client.dynamic.Dispatch(obj)
    return _manage(obj, clsid, interface=interface)


@overload
def CoGetObject(displayname: str, interface: Type[_T_IUnknown]) -> _T_IUnknown: ...
@overload
def CoGetObject(
    displayname: str, interface: None = None, dynamic: bool = False
) -> Any: ...
def CoGetObject(
    displayname: str,
    interface: Optional[Type[IUnknown]] = None,
    dynamic: bool = False,
) -> Any:
    """Create an object by calling CoGetObject(displayname).

    Additional parameters have the same meaning as in CreateObject().
    """
    if dynamic:
        if interface is not None:
            raise ValueError("interface and dynamic are mutually exclusive")
        interface = automation.IDispatch
    punk = comtypes.CoGetObject(displayname, interface)
    if dynamic:
        return comtypes.client.dynamic.Dispatch(punk)
    return _manage(punk, clsid=None, interface=interface)
