# comtypes version numbers follow semver (http://semver.org/) and PEP 440
__version__ = "1.4.4"

import atexit
from ctypes import *
from ctypes import _Pointer, _SimpleCData

try:
    from _ctypes import COMError  # noqa
except ImportError as e:
    msg = "\n".join(
        (
            "COM technology not available (maybe it's the wrong platform).",
            "Note that COM is only supported on Windows.",
            "For more details, please check: "
            "https://learn.microsoft.com/en-us/windows/win32/com",
        )
    )
    raise ImportError(msg) from e
import logging
import sys

# fmt: off
from typing import (  # noqa
    Any, ClassVar, overload, TYPE_CHECKING, TypeVar,
    # instead of `builtins`. see PEP585
    Dict, List, Tuple, Type,
    # instead of `collections.abc`. see PEP585
    Callable, Iterable, Iterator,
    # instead of `A | B` and `None | A`. see PEP604
    Optional, Union as _UnionT,  # avoiding confusion with `ctypes.Union`
)
# fmt: on
if TYPE_CHECKING:
    from ctypes import _CData  # only in `typeshed`, private in runtime
    from comtypes import hints as hints  # noqa  # type: ignore
else:
    _CData = _SimpleCData.__mro__[:-1][-1]

from comtypes.GUID import GUID
from comtypes import patcher  # noqa
from comtypes._npsupport import interop as npsupport  # noqa
from comtypes._memberspec import _encode_idl  # noqa
from comtypes._tlib_version_checker import _check_version  # noqa
from comtypes._bstr import BSTR  # noqa
from comtypes._py_instance_method import instancemethod  # noqa
from comtypes._idl_stuff import defaultvalue, helpstring, dispid  # noqa
from comtypes._idl_stuff import STDMETHOD, DISPMETHOD, DISPPROPERTY, COMMETHOD  # noqa

_all_slice = slice(None, None, None)


class NullHandler(logging.Handler):
    """A Handler that does nothing."""

    def emit(self, record):
        pass


logger = logging.getLogger(__name__)

# Add a NULL handler to the comtypes logger.  This prevents getting a
# message like this:
#    No handlers could be found for logger "comtypes"
# when logging is not configured and logger.error() is called.
logger.addHandler(NullHandler())


class ReturnHRESULT(Exception):
    """ReturnHRESULT(hresult, text)

    Return a hresult code from a COM method implementation
    without logging an error.
    """


# class IDLWarning(UserWarning):
#    "Warn about questionable type information"

_GUID = GUID
IID = GUID
DWORD = c_ulong

wireHWND = c_ulong

################################################################
# About COM apartments:
# http://blogs.msdn.com/larryosterman/archive/2004/04/28/122240.aspx
################################################################

################################################################
# constants for object creation
CLSCTX_INPROC_SERVER = 1
CLSCTX_INPROC_HANDLER = 2
CLSCTX_LOCAL_SERVER = 4

CLSCTX_INPROC = 3
CLSCTX_SERVER = 5
CLSCTX_ALL = 7

CLSCTX_INPROC_SERVER16 = 8
CLSCTX_REMOTE_SERVER = 16
CLSCTX_INPROC_HANDLER16 = 32
CLSCTX_RESERVED1 = 64
CLSCTX_RESERVED2 = 128
CLSCTX_RESERVED3 = 256
CLSCTX_RESERVED4 = 512
CLSCTX_NO_CODE_DOWNLOAD = 1024
CLSCTX_RESERVED5 = 2048
CLSCTX_NO_CUSTOM_MARSHAL = 4096
CLSCTX_ENABLE_CODE_DOWNLOAD = 8192
CLSCTX_NO_FAILURE_LOG = 16384
CLSCTX_DISABLE_AAA = 32768
CLSCTX_ENABLE_AAA = 65536
CLSCTX_FROM_DEFAULT_CONTEXT = 131072

tagCLSCTX = c_int  # enum
CLSCTX = tagCLSCTX

# Constants for security setups
SEC_WINNT_AUTH_IDENTITY_UNICODE = 0x2
RPC_C_AUTHN_WINNT = 10
RPC_C_AUTHZ_NONE = 0
RPC_C_AUTHN_LEVEL_CONNECT = 2
RPC_C_IMP_LEVEL_IMPERSONATE = 3
EOAC_NONE = 0


################################################################
# Initialization and shutdown
_ole32 = oledll.ole32
_ole32_nohresult = windll.ole32  # use this for functions that don't return a HRESULT

COINIT_MULTITHREADED = 0x0
COINIT_APARTMENTTHREADED = 0x2
COINIT_DISABLE_OLE1DDE = 0x4
COINIT_SPEED_OVER_MEMORY = 0x8


def CoInitialize():
    return CoInitializeEx(COINIT_APARTMENTTHREADED)


def CoInitializeEx(flags=None):
    if flags is None:
        flags = getattr(sys, "coinit_flags", COINIT_APARTMENTTHREADED)
    logger.debug("CoInitializeEx(None, %s)", flags)
    _ole32.CoInitializeEx(None, flags)


# COM is initialized automatically for the thread that imports this
# module for the first time.  sys.coinit_flags is passed as parameter
# to CoInitializeEx, if defined, otherwise COINIT_APARTMENTTHREADED
# (COINIT_MULTITHREADED on Windows CE) is used.
#
# A shutdown function is registered with atexit, so that
# CoUninitialize is called when Python is shut down.
CoInitializeEx()


# We need to have CoUninitialize for multithreaded model where we have
# to initialize and uninitialize COM for every new thread (except main)
# in which we are using COM
def CoUninitialize():
    logger.debug("CoUninitialize()")
    _ole32_nohresult.CoUninitialize()


################################################################
# global registries.

# allows to find interface classes by guid strings (iid)
com_interface_registry = {}

# allows to find coclasses by guid strings (clsid)
com_coclass_registry = {}


def _is_object(obj):
    """This function determines if the argument is a COM object.  It
    is used in several places to determine whether propputref or
    propput setters have to be used."""
    from comtypes.automation import VARIANT

    # A COM pointer is an 'Object'
    if isinstance(obj, POINTER(IUnknown)):
        return True
    # A COM pointer in a VARIANT is an 'Object', too
    elif isinstance(obj, VARIANT) and isinstance(obj.value, POINTER(IUnknown)):
        return True
    # It may be a dynamic dispatch object.
    return hasattr(obj, "_comobj")


################################################################
# IUnknown, the root of all evil...

_T_IUnknown = TypeVar("_T_IUnknown", bound="IUnknown")

from comtypes.unknwn import IUnknown, _shutdown  # noqa

atexit.register(_shutdown)


################################################################
# IPersist is a trivial interface, which allows to ask an object about
# its clsid.
class IPersist(IUnknown):
    _iid_ = GUID("{0000010C-0000-0000-C000-000000000046}")
    _idlflags_ = []
    _methods_ = [
        COMMETHOD([], HRESULT, "GetClassID", (["out"], POINTER(GUID), "pClassID")),
    ]
    if TYPE_CHECKING:
        # Should this be "normal" method that calls `self._GetClassID`?
        def GetClassID(self) -> GUID:
            """Returns the CLSID that uniquely represents an object class that
            defines the code that can manipulate the object's data.
            """
            ...


class IServiceProvider(IUnknown):
    _iid_ = GUID("{6D5140C1-7436-11CE-8034-00AA006009FA}")
    _QueryService: Callable[[Any, Any, Any], int]
    # Overridden QueryService to make it nicer to use (passing it an
    # interface and it returns a pointer to that interface)
    def QueryService(
        self, serviceIID: GUID, interface: Type[_T_IUnknown]
    ) -> _T_IUnknown:
        p = POINTER(interface)()
        self._QueryService(byref(serviceIID), byref(interface._iid_), byref(p))
        return p  # type: ignore

    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "QueryService",
            (["in"], POINTER(GUID), "guidService"),
            (["in"], POINTER(GUID), "riid"),
            (["in"], POINTER(c_void_p), "ppvObject"),
        )
    ]


################################################################


@overload
def CoGetObject(displayname: str, interface: None) -> IUnknown:
    ...


@overload
def CoGetObject(displayname: str, interface: Type[_T_IUnknown]) -> _T_IUnknown:
    ...


def CoGetObject(displayname: str, interface: Optional[Type[IUnknown]]) -> IUnknown:
    """Convert a displayname to a moniker, then bind and return the object
    identified by the moniker."""
    if interface is None:
        interface = IUnknown
    punk = POINTER(interface)()
    # Do we need a way to specify the BIND_OPTS parameter?
    _ole32.CoGetObject(str(displayname), None, byref(interface._iid_), byref(punk))
    return punk  # type: ignore


_pUnkOuter = Type["_Pointer[IUnknown]"]


@overload
def CoCreateInstance(
    clsid: GUID,
    interface: None = None,
    clsctx: Optional[int] = None,
    punkouter: Optional[_pUnkOuter] = None,
) -> IUnknown:
    ...


@overload
def CoCreateInstance(
    clsid: GUID,
    interface: Type[_T_IUnknown],
    clsctx: Optional[int] = None,
    punkouter: Optional[_pUnkOuter] = None,
) -> _T_IUnknown:
    ...


def CoCreateInstance(
    clsid: GUID,
    interface: Optional[Type[IUnknown]] = None,
    clsctx: Optional[int] = None,
    punkouter: Optional[_pUnkOuter] = None,
) -> IUnknown:
    """The basic windows api to create a COM class object and return a
    pointer to an interface.
    """
    if clsctx is None:
        clsctx = CLSCTX_SERVER
    if interface is None:
        interface = IUnknown
    p = POINTER(interface)()
    iid = interface._iid_
    _ole32.CoCreateInstance(byref(clsid), punkouter, clsctx, byref(iid), byref(p))
    return p  # type: ignore


if TYPE_CHECKING:

    @overload
    def CoGetClassObject(clsid, clsctx=None, pServerInfo=None, interface=None):
        # type: (GUID, Optional[int], Optional[COSERVERINFO], None) -> hints.IClassFactory
        pass

    @overload
    def CoGetClassObject(clsid, clsctx=None, pServerInfo=None, interface=None):
        # type: (GUID, Optional[int], Optional[COSERVERINFO], Type[_T_IUnknown]) -> _T_IUnknown
        pass


def CoGetClassObject(clsid, clsctx=None, pServerInfo=None, interface=None):
    # type: (GUID, Optional[int], Optional[COSERVERINFO], Optional[Type[IUnknown]]) -> IUnknown
    if clsctx is None:
        clsctx = CLSCTX_SERVER
    if interface is None:
        import comtypes.server

        interface = comtypes.server.IClassFactory
    p = POINTER(interface)()
    _CoGetClassObject(clsid, clsctx, pServerInfo, interface._iid_, byref(p))
    return p  # type: ignore


@overload
def GetActiveObject(clsid: GUID, interface: None = None) -> IUnknown:
    ...


@overload
def GetActiveObject(clsid: GUID, interface: Type[_T_IUnknown]) -> _T_IUnknown:
    ...


def GetActiveObject(
    clsid: GUID, interface: Optional[Type[IUnknown]] = None
) -> IUnknown:
    """Retrieves a pointer to a running object"""
    p = POINTER(IUnknown)()
    oledll.oleaut32.GetActiveObject(byref(clsid), None, byref(p))
    if interface is not None:
        p = p.QueryInterface(interface)  # type: ignore
    return p  # type: ignore


class MULTI_QI(Structure):
    _fields_ = [("pIID", POINTER(GUID)), ("pItf", POINTER(c_void_p)), ("hr", HRESULT)]
    if TYPE_CHECKING:
        pIID: GUID
        pItf: _Pointer[c_void_p]
        hr: HRESULT


class _COAUTHIDENTITY(Structure):
    _fields_ = [
        ("User", POINTER(c_ushort)),
        ("UserLength", c_ulong),
        ("Domain", POINTER(c_ushort)),
        ("DomainLength", c_ulong),
        ("Password", POINTER(c_ushort)),
        ("PasswordLength", c_ulong),
        ("Flags", c_ulong),
    ]


COAUTHIDENTITY = _COAUTHIDENTITY


class _COAUTHINFO(Structure):
    _fields_ = [
        ("dwAuthnSvc", c_ulong),
        ("dwAuthzSvc", c_ulong),
        ("pwszServerPrincName", c_wchar_p),
        ("dwAuthnLevel", c_ulong),
        ("dwImpersonationLevel", c_ulong),
        ("pAuthIdentityData", POINTER(_COAUTHIDENTITY)),
        ("dwCapabilities", c_ulong),
    ]


COAUTHINFO = _COAUTHINFO


class _COSERVERINFO(Structure):
    _fields_ = [
        ("dwReserved1", c_ulong),
        ("pwszName", c_wchar_p),
        ("pAuthInfo", POINTER(_COAUTHINFO)),
        ("dwReserved2", c_ulong),
    ]
    if TYPE_CHECKING:
        dwReserved1: int
        pwszName: Optional[str]
        pAuthInfo: _COAUTHINFO
        dwReserved2: int


COSERVERINFO = _COSERVERINFO
_CoGetClassObject = _ole32.CoGetClassObject
_CoGetClassObject.argtypes = [
    POINTER(GUID),
    DWORD,
    POINTER(COSERVERINFO),
    POINTER(GUID),
    POINTER(c_void_p),
]


class tagBIND_OPTS(Structure):
    _fields_ = [
        ("cbStruct", c_ulong),
        ("grfFlags", c_ulong),
        ("grfMode", c_ulong),
        ("dwTickCountDeadline", c_ulong),
    ]


# XXX Add __init__ which sets cbStruct?
BIND_OPTS = tagBIND_OPTS


class tagBIND_OPTS2(Structure):
    _fields_ = [
        ("cbStruct", c_ulong),
        ("grfFlags", c_ulong),
        ("grfMode", c_ulong),
        ("dwTickCountDeadline", c_ulong),
        ("dwTrackFlags", c_ulong),
        ("dwClassContext", c_ulong),
        ("locale", c_ulong),
        ("pServerInfo", POINTER(_COSERVERINFO)),
    ]


# XXX Add __init__ which sets cbStruct?
BINDOPTS2 = tagBIND_OPTS2

# Structures for security setups
#########################################
class _SEC_WINNT_AUTH_IDENTITY(Structure):
    _fields_ = [
        ("User", POINTER(c_ushort)),
        ("UserLength", c_ulong),
        ("Domain", POINTER(c_ushort)),
        ("DomainLength", c_ulong),
        ("Password", POINTER(c_ushort)),
        ("PasswordLength", c_ulong),
        ("Flags", c_ulong),
    ]


SEC_WINNT_AUTH_IDENTITY = _SEC_WINNT_AUTH_IDENTITY


class _SOLE_AUTHENTICATION_INFO(Structure):
    _fields_ = [
        ("dwAuthnSvc", c_ulong),
        ("dwAuthzSvc", c_ulong),
        ("pAuthInfo", POINTER(_SEC_WINNT_AUTH_IDENTITY)),
    ]


SOLE_AUTHENTICATION_INFO = _SOLE_AUTHENTICATION_INFO


class _SOLE_AUTHENTICATION_LIST(Structure):
    _fields_ = [
        ("cAuthInfo", c_ulong),
        ("pAuthInfo", POINTER(_SOLE_AUTHENTICATION_INFO)),
    ]


SOLE_AUTHENTICATION_LIST = _SOLE_AUTHENTICATION_LIST


@overload
def CoCreateInstanceEx(
    clsid: GUID,
    interface: None = None,
    clsctx: Optional[int] = None,
    machine: Optional[str] = None,
    pServerInfo: Optional[COSERVERINFO] = None,
) -> IUnknown:
    ...


@overload
def CoCreateInstanceEx(
    clsid: GUID,
    interface: Type[_T_IUnknown],
    clsctx: Optional[int] = None,
    machine: Optional[str] = None,
    pServerInfo: Optional[COSERVERINFO] = None,
) -> _T_IUnknown:
    ...


def CoCreateInstanceEx(
    clsid: GUID,
    interface: Optional[Type[IUnknown]] = None,
    clsctx: Optional[int] = None,
    machine: Optional[str] = None,
    pServerInfo: Optional[COSERVERINFO] = None,
) -> IUnknown:
    """The basic windows api to create a COM class object and return a
    pointer to an interface, possibly on another machine.

    Passing both "machine" and "pServerInfo" results in a ValueError.

    """
    if clsctx is None:
        clsctx = CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER

    if pServerInfo is not None:
        if machine is not None:
            msg = "Can not specify both machine name and server info"
            raise ValueError(msg)
    elif machine is not None:
        serverinfo = COSERVERINFO()
        serverinfo.pwszName = machine
        pServerInfo = byref(serverinfo)  # type: ignore

    if interface is None:
        interface = IUnknown
    multiqi = MULTI_QI()
    multiqi.pIID = pointer(interface._iid_)  # type: ignore
    _ole32.CoCreateInstanceEx(
        byref(clsid), None, clsctx, pServerInfo, 1, byref(multiqi)
    )
    return cast(multiqi.pItf, POINTER(interface))  # type: ignore


################################################################
from comtypes._comobject import COMObject

# What's a coclass?
# a POINTER to a coclass is allowed as parameter in a function declaration:
# http://msdn.microsoft.com/library/en-us/midl/midl/oleautomation.asp

from comtypes._meta import _coclass_meta


class CoClass(COMObject, metaclass=_coclass_meta):
    pass


################################################################


# fmt: off
__known_symbols__ = [
    "BIND_OPTS", "tagBIND_OPTS", "BINDOPTS2", "tagBIND_OPTS2", "BSTR",
    "_check_version", "CLSCTX", "tagCLSCTX", "CLSCTX_ALL",
    "CLSCTX_DISABLE_AAA", "CLSCTX_ENABLE_AAA", "CLSCTX_ENABLE_CODE_DOWNLOAD",
    "CLSCTX_FROM_DEFAULT_CONTEXT", "CLSCTX_INPROC", "CLSCTX_INPROC_HANDLER",
    "CLSCTX_INPROC_HANDLER16", "CLSCTX_INPROC_SERVER",
    "CLSCTX_INPROC_SERVER16", "CLSCTX_LOCAL_SERVER", "CLSCTX_NO_CODE_DOWNLOAD",
    "CLSCTX_NO_CUSTOM_MARSHAL", "CLSCTX_NO_FAILURE_LOG",
    "CLSCTX_REMOTE_SERVER", "CLSCTX_RESERVED1", "CLSCTX_RESERVED2",
    "CLSCTX_RESERVED3", "CLSCTX_RESERVED4", "CLSCTX_RESERVED5",
    "CLSCTX_SERVER", "_COAUTHIDENTITY", "COAUTHIDENTITY", "_COAUTHINFO",
    "COAUTHINFO", "CoClass", "CoCreateInstance", "CoCreateInstanceEx",
    "_CoGetClassObject", "CoGetClassObject", "CoGetObject",
    "COINIT_APARTMENTTHREADED", "COINIT_DISABLE_OLE1DDE",
    "COINIT_MULTITHREADED", "COINIT_SPEED_OVER_MEMORY", "CoInitialize",
    "CoInitializeEx", "COMError", "COMMETHOD", "COMObject", "_COSERVERINFO",
    "COSERVERINFO", "CoUninitialize", "dispid", "DISPMETHOD", "DISPPROPERTY",
    "DWORD", "EOAC_NONE", "GetActiveObject", "_GUID", "GUID", "helpstring",
    "IID", "IPersist", "IServiceProvider", "IUnknown", "MULTI_QI",
    "ReturnHRESULT", "RPC_C_AUTHN_LEVEL_CONNECT", "RPC_C_AUTHN_WINNT",
    "RPC_C_AUTHZ_NONE", "RPC_C_IMP_LEVEL_IMPERSONATE",
    "_SEC_WINNT_AUTH_IDENTITY", "SEC_WINNT_AUTH_IDENTITY",
    "SEC_WINNT_AUTH_IDENTITY_UNICODE", "_SOLE_AUTHENTICATION_INFO",
    "SOLE_AUTHENTICATION_INFO", "_SOLE_AUTHENTICATION_LIST",
    "SOLE_AUTHENTICATION_LIST", "STDMETHOD", "wireHWND",
]
# fmt: on
