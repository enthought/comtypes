# comtypes version numbers follow semver (http://semver.org/) and PEP 440
__version__ = "1.4.2"

import atexit
from ctypes import *
from ctypes import _Pointer, _SimpleCData

try:
    from _ctypes import COMError
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
import os
import sys
import types

# fmt: off
from typing import (
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
    from comtypes import hints as hints  # type: ignore
else:
    _CData = _SimpleCData.__mro__[:-1][-1]

from comtypes.GUID import GUID
from comtypes import patcher
from comtypes._npsupport import interop as npsupport
from comtypes._memberspec import (
    ComMemberGenerator,
    _ComMemberSpec,
    DispMemberGenerator,
    _DispMemberSpec,
    _encode_idl,
    _resolve_argspec,
)


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


def _check_version(actual, tlib_cached_mtime=None):
    from comtypes.tools.codegenerator import version as required

    if actual != required:
        raise ImportError("Wrong version")
    if not hasattr(sys, "frozen"):
        g = sys._getframe(1).f_globals
        tlb_path = g.get("typelib_path")
        try:
            tlib_curr_mtime = os.stat(tlb_path).st_mtime
        except (OSError, TypeError):
            return
        if not tlib_cached_mtime or abs(tlib_curr_mtime - tlib_cached_mtime) >= 1:
            raise ImportError("Typelib different than module")


pythonapi.PyInstanceMethod_New.argtypes = [py_object]
pythonapi.PyInstanceMethod_New.restype = py_object
PyInstanceMethod_Type = type(pythonapi.PyInstanceMethod_New(id))


def instancemethod(func, inst, cls):
    mth = PyInstanceMethod_Type(func)
    if inst is None:
        return mth
    return mth.__get__(inst)


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


def _shutdown(
    func=_ole32_nohresult.CoUninitialize,
    _debug=logger.debug,
    _exc_clear=getattr(sys, "exc_clear", lambda: None),
):
    # Make sure no COM pointers stay in exception frames.
    _exc_clear()
    # Sometimes, CoUninitialize, running at Python shutdown,
    # raises an exception.  We suppress this when __debug__ is
    # False.
    _debug("Calling CoUninitialize()")
    if __debug__:
        func()
    else:
        try:
            func()
        except WindowsError:
            pass
    # Set the flag which means that calling obj.Release() is no longer
    # needed.
    if _cominterface_meta is not None:
        _cominterface_meta._com_shutting_down = True
    _debug("CoUninitialize() done.")


atexit.register(_shutdown)

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
# The metaclasses...


class _cominterface_meta(type):
    """Metaclass for COM interfaces.  Automatically creates high level
    methods from COMMETHOD lists.
    """

    _case_insensitive_: bool
    _iid_: GUID
    _methods_: List[_ComMemberSpec]
    _disp_methods_: List[_DispMemberSpec]

    # This flag is set to True by the atexit handler which calls
    # CoUninitialize.
    _com_shutting_down = False

    # Creates also a POINTER type for the newly created class.
    def __new__(cls, name, bases, namespace):
        methods = namespace.pop("_methods_", None)
        dispmethods = namespace.pop("_disp_methods_", None)
        new_cls = type.__new__(cls, name, bases, namespace)

        if methods is not None:
            new_cls._methods_ = methods
        if dispmethods is not None:
            new_cls._disp_methods_ = dispmethods

        # If we sublass a COM interface, for example:
        #
        # class IDispatch(IUnknown):
        #     ....
        #
        # then we need to make sure that POINTER(IDispatch) is a
        # subclass of POINTER(IUnknown) because of the way ctypes
        # typechecks work.
        if bases == (object,):
            _ptr_bases = (new_cls, _compointer_base)
        else:
            _ptr_bases = (new_cls, POINTER(bases[0]))

        # The interface 'new_cls' is used as a mixin.
        p = type(_compointer_base)(
            "POINTER(%s)" % new_cls.__name__,
            _ptr_bases,
            {"__com_interface__": new_cls, "_needs_com_addref_": None},
        )

        from ctypes import _pointer_type_cache

        _pointer_type_cache[new_cls] = p

        if new_cls._case_insensitive_:

            @patcher.Patch(p)
            class CaseInsensitive(object):
                # case insensitive attributes for COM methods and properties
                def __getattr__(self, name):
                    """Implement case insensitive access to methods and properties"""
                    try:
                        fixed_name = self.__map_case__[name.lower()]
                    except KeyError:
                        raise AttributeError(name)
                    if fixed_name != name:  # prevent unbounded recursion
                        return getattr(self, fixed_name)
                    raise AttributeError(name)

                # __setattr__ is pretty heavy-weight, because it is called for
                # EVERY attribute assignment.  Settings a non-com attribute
                # through this function takes 8.6 usec, while without this
                # function it takes 0.7 sec - 12 times slower.
                #
                # How much faster would this be if implemented in C?
                def __setattr__(self, name, value):
                    """Implement case insensitive access to methods and properties"""
                    object.__setattr__(
                        self, self.__map_case__.get(name.lower(), name), value
                    )

        @patcher.Patch(POINTER(p))
        class ReferenceFix(object):
            def __setitem__(self, index, value):
                # We override the __setitem__ method of the
                # POINTER(POINTER(interface)) type, so that the COM
                # reference count is managed correctly.
                #
                # This is so that we can implement COM methods that have to
                # return COM pointers more easily and consistent.  Instead of
                # using CopyComPointer in the method implementation, we can
                # simply do:
                #
                # def GetTypeInfo(self, this, ..., pptinfo):
                #     if not pptinfo: return E_POINTER
                #     pptinfo[0] = a_com_interface_pointer
                #     return S_OK
                if index != 0:
                    # CopyComPointer, which is in _ctypes, does only
                    # handle an index of 0.  This code does what
                    # CopyComPointer should do if index != 0.
                    if bool(value):
                        value.AddRef()
                    super(POINTER(p), self).__setitem__(index, value)
                    return
                from _ctypes import CopyComPointer

                CopyComPointer(value, self)

        return new_cls

    def __setattr__(self, name, value):
        if name == "_methods_":
            # XXX I'm no longer sure why the code generator generates
            # "_methods_ = []" in the interface definition, and later
            # overrides this by "Interface._methods_ = [...]
            # assert self.__dict__.get("_methods_", None) is None
            self._make_methods(value)
            self._make_specials()
        elif name == "_disp_methods_":
            assert self.__dict__.get("_disp_methods_", None) is None
            self._make_dispmethods(value)
            self._make_specials()
        type.__setattr__(self, name, value)

    def _make_specials(self):
        # This call installs methods that forward the Python protocols
        # to COM protocols.

        def has_name(name):
            # Determine whether a property or method named 'name'
            # exists
            if self._case_insensitive_:
                return name.lower() in self.__map_case__
            return hasattr(self, name)

        # XXX These special methods should be generated by the code generator.
        if has_name("Count"):

            @patcher.Patch(self)
            class _(object):
                def __len__(self):
                    "Return the the 'self.Count' property."
                    return self.Count

        if has_name("Item"):

            @patcher.Patch(self)
            class _(object):
                # 'Item' is the 'default' value.  Make it available by
                # calling the instance (Not sure this makes sense, but
                # win32com does this also).
                def __call__(self, *args, **kw):
                    "Return 'self.Item(*args, **kw)'"
                    return self.Item(*args, **kw)

                # does this make sense? It seems that all standard typelibs I've
                # seen so far that support .Item also support ._NewEnum
                @patcher.no_replace
                def __getitem__(self, index):
                    "Return 'self.Item(index)'"
                    # Handle tuples and all-slice
                    if isinstance(index, tuple):
                        args = index
                    elif index == _all_slice:
                        args = ()
                    else:
                        args = (index,)

                    try:
                        result = self.Item(*args)
                    except COMError as err:
                        (hresult, text, details) = err.args
                        if hresult == -2147352565:  # DISP_E_BADINDEX
                            raise IndexError("invalid index")
                        else:
                            raise

                    # Note that result may be NULL COM pointer. There is no way
                    # to interpret this properly, so it is returned as-is.

                    # Hm, should we call __ctypes_from_outparam__ on the
                    # result?
                    return result

                @patcher.no_replace
                def __setitem__(self, index, value):
                    "Attempt 'self.Item[index] = value'"
                    try:
                        self.Item[index] = value
                    except COMError as err:
                        (hresult, text, details) = err.args
                        if hresult == -2147352565:  # DISP_E_BADINDEX
                            raise IndexError("invalid index")
                        else:
                            raise
                    except TypeError:
                        msg = "%r object does not support item assignment"
                        raise TypeError(msg % type(self))

        if has_name("_NewEnum"):

            @patcher.Patch(self)
            class _(object):
                def __iter__(self):
                    "Return an iterator over the _NewEnum collection."
                    # This method returns a pointer to _some_ _NewEnum interface.
                    # It relies on the fact that the code generator creates next()
                    # methods for them automatically.
                    #
                    # Better would maybe to return an object that
                    # implements the Python iterator protocol, and
                    # forwards the calls to the COM interface.
                    enum = self._NewEnum
                    if isinstance(enum, types.MethodType):
                        # _NewEnum should be a propget property, with dispid -4.
                        #
                        # Sometimes, however, it is a method.
                        enum = enum()
                    if hasattr(enum, "Next"):
                        return enum
                    # _NewEnum returns an IUnknown pointer, QueryInterface() it to
                    # IEnumVARIANT
                    from comtypes.automation import IEnumVARIANT

                    return enum.QueryInterface(IEnumVARIANT)

    def _make_case_insensitive(self):
        # The __map_case__ dictionary maps lower case names to the
        # names in the original spelling to enable case insensitive
        # method and attribute access.
        try:
            self.__dict__["__map_case__"]
        except KeyError:
            d = {}
            d.update(getattr(self, "__map_case__", {}))
            self.__map_case__ = d

    def _make_dispmethods(self, methods: List[_DispMemberSpec]) -> None:
        if self._case_insensitive_:
            self._make_case_insensitive()
        # create dispinterface methods and properties on the interface 'self'
        member_gen = DispMemberGenerator(self.__name__)
        for m in methods:
            member_gen.add(m)
        for name, func_or_prop, is_prop in member_gen.items():
            setattr(self, name, func_or_prop)
            # COM is case insensitive.
            # For a method, this is the real name.  For a property,
            # this is the name WITHOUT the _set_ or _get_ prefix.
            if self._case_insensitive_:
                self.__map_case__[name.lower()] = name
                if is_prop:
                    self.__map_case__[name[5:].lower()] = name[5:]
        for name, accessor in member_gen.properties():
            setattr(self, name, accessor)
            # COM is case insensitive
            if self._case_insensitive_:
                self.__map_case__[name.lower()] = name

    def __get_baseinterface_methodcount(self):
        "Return the number of com methods in the base interfaces"
        try:
            result = 0
            for itf in self.mro()[1:-1]:
                result += len(itf.__dict__["_methods_"])
            return result
        except KeyError as err:
            (name,) = err.args
            if name == "_methods_":
                raise TypeError("baseinterface '%s' has no _methods_" % itf.__name__)
            raise

    def _make_methods(self, methods: List[_ComMemberSpec]) -> None:
        if self._case_insensitive_:
            self._make_case_insensitive()
        # register com interface. we insist on an _iid_ in THIS class!
        try:
            iid = self.__dict__["_iid_"]
        except KeyError:
            raise AttributeError("this class must define an _iid_")
        else:
            com_interface_registry[str(iid)] = self
        # create members
        vtbl_offset = self.__get_baseinterface_methodcount()
        member_gen = ComMemberGenerator(self.__name__, vtbl_offset, self._iid_)
        # create private low level, and public high level methods
        for m in methods:
            member_gen.add(m)
        for name, func, raw_func, is_prop in member_gen.methods():
            raw_mth = instancemethod(raw_func, None, self)
            setattr(self, "_%s__com_%s" % (self.__name__, name), raw_mth)
            mth = instancemethod(func, None, self)
            if not is_prop:
                # We install the method in the class, except when it's a property.
                # And we make sure we don't overwrite a property that's already present.
                mthname = name if not hasattr(self, name) else ("_%s" % name)
                setattr(self, mthname, mth)
            # For a method, this is the real name.
            # For a property, this is the name WITHOUT the _set_ or _get_ prefix.
            if self._case_insensitive_:
                self.__map_case__[name.lower()] = name
                if is_prop:
                    self.__map_case__[name[5:].lower()] = name[5:]
        # create public properties / attribute accessors
        for name, accessor in member_gen.properties():
            # Again, we should not overwrite class attributes that are already present.
            propname = name if not hasattr(self, name) else ("_%s" % name)
            setattr(self, propname, accessor)
            # COM is case insensitive
            if self._case_insensitive_:
                self.__map_case__[name.lower()] = name


################################################################


class _compointer_meta(type(c_void_p), _cominterface_meta):
    "metaclass for COM interface pointer classes"
    # no functionality, but needed to avoid a metaclass conflict


class _compointer_base(c_void_p, metaclass=_compointer_meta):
    "base class for COM interface pointer classes"

    def __del__(self, _debug=logger.debug):
        "Release the COM refcount we own."
        if self:
            # comtypes calls CoUninitialize() when the atexit handlers
            # runs.  CoUninitialize() cleans up the COM objects that
            # are still alive. Python COM pointers may still be
            # present but we can no longer call Release() on them -
            # this may give a protection fault.  So we need the
            # _com_shutting_down flag.
            #
            if not type(self)._com_shutting_down:
                _debug("Release %s", self)
                self.Release()

    def __cmp__(self, other):
        """Compare pointers to COM interfaces."""
        # COM identity rule
        #
        # XXX To compare COM interface pointers, should we
        # automatically QueryInterface for IUnknown on both items, and
        # compare the pointer values?
        if not isinstance(other, _compointer_base):
            return 1

        # get the value property of the c_void_p baseclass, this is the pointer value
        return cmp(
            super(_compointer_base, self).value, super(_compointer_base, other).value
        )

    def __eq__(self, other):
        if not isinstance(other, _compointer_base):
            return False
        # get the value property of the c_void_p baseclass, this is the pointer value
        return (
            super(_compointer_base, self).value == super(_compointer_base, other).value
        )

    def __hash__(self):
        """Return the hash value of the pointer."""
        # hash the pointer values
        return hash(super(_compointer_base, self).value)

    # redefine the .value property; return the object itself.
    def __get_value(self):
        return self

    value = property(__get_value, doc="""Return self.""")

    def __repr__(self):
        ptr = super(_compointer_base, self).value
        return "<%s ptr=0x%x at %x>" % (self.__class__.__name__, ptr or 0, id(self))

    # This fixes the problem when there are multiple python interface types
    # wrapping the same COM interface.  This could happen because some interfaces
    # are contained in multiple typelibs.
    #
    # It also allows to pass a CoClass instance to an api
    # expecting a COM interface.
    @classmethod
    def from_param(cls, value):
        """Convert 'value' into a COM pointer to the interface.

        This method accepts a COM pointer, or a CoClass instance
        which is QueryInterface()d."""
        if value is None:
            return None
        # CLF: 2013-01-18
        # A default value of 0, meaning null, can pass through to here.
        if value == 0:
            return None
        if isinstance(value, cls):
            return value
        # multiple python interface types for the same COM interface.
        # Do we need more checks here?
        if cls._iid_ == getattr(value, "_iid_", None):
            return value
        # Accept an CoClass instance which exposes the interface required.
        try:
            table = value._com_pointers_
        except AttributeError:
            pass
        else:
            try:
                # a kind of QueryInterface
                return table[cls._iid_]
            except KeyError:
                raise TypeError("Interface %s not supported" % cls._iid_)
        return value.QueryInterface(cls.__com_interface__)


################################################################


class BSTR(_SimpleCData):
    "The windows BSTR data type"
    _type_ = "X"
    _needsfree = False

    def __repr__(self):
        return "%s(%r)" % (self.__class__.__name__, self.value)

    def __ctypes_from_outparam__(self):
        self._needsfree = True
        return self.value

    def __del__(self, _free=windll.oleaut32.SysFreeString):
        # Free the string if self owns the memory
        # or if instructed by __ctypes_from_outparam__.
        if self._b_base_ is None or self._needsfree:
            _free(self)

    @classmethod
    def from_param(cls, value):
        """Convert into a foreign function call parameter."""
        if isinstance(value, cls):
            return value
        # Although the builtin SimpleCData.from_param call does the
        # right thing, it doesn't ensure that SysFreeString is called
        # on destruction.
        return cls(value)


################################################################
# IDL stuff


class helpstring(str):
    "Specifies the helpstring for a COM method or property."


class defaultvalue(object):
    "Specifies the default value for parameters marked optional."

    def __init__(self, value):
        self.value = value


class dispid(int):
    "Specifies the DISPID of a method or property."


# XXX STDMETHOD, COMMETHOD, DISPMETHOD, and DISPPROPERTY should return
# instances with more methods or properties, and should not behave as an unpackable.


def STDMETHOD(restype, name, argtypes=()) -> _ComMemberSpec:
    "Specifies a COM method slot without idlflags"
    return _ComMemberSpec(restype, name, argtypes, None, (), None)


def DISPMETHOD(idlflags, restype, name, *argspec) -> _DispMemberSpec:
    "Specifies a method of a dispinterface"
    return _DispMemberSpec("DISPMETHOD", name, tuple(idlflags), restype, argspec)


def DISPPROPERTY(idlflags, proptype, name) -> _DispMemberSpec:
    "Specifies a property of a dispinterface"
    return _DispMemberSpec("DISPPROPERTY", name, tuple(idlflags), proptype, ())


# tuple(idlflags) is for the method itself: (dispid, 'readonly')

# sample generated code:
#     DISPPROPERTY([5, 'readonly'], OLE_YSIZE_HIMETRIC, 'Height'),
#     DISPMETHOD(
#         [6], None, 'Render', ([], c_int, 'hdc'), ([], c_int, 'x'), ([], c_int, 'y')
#     )


def COMMETHOD(idlflags, restype, methodname, *argspec) -> _ComMemberSpec:
    """Specifies a COM method slot with idlflags.

    XXX should explain the sematics of the arguments.
    """
    # collect all helpstring instances
    # We should suppress docstrings when Python is started with -OO
    # join them together(does this make sense?) and replace by None if empty.
    helptext = "".join(t for t in idlflags if isinstance(t, helpstring)) or None
    paramflags, argtypes = _resolve_argspec(argspec)
    if "propget" in idlflags:
        name = "_get_%s" % methodname
    elif "propput" in idlflags:
        name = "_set_%s" % methodname
    elif "propputref" in idlflags:
        name = "_setref_%s" % methodname
    else:
        name = methodname
    return _ComMemberSpec(
        restype, name, argtypes, paramflags, tuple(idlflags), helptext
    )


################################################################
# IUnknown, the root of all evil...

_T_IUnknown = TypeVar("_T_IUnknown", bound="IUnknown")

if TYPE_CHECKING:

    class _IUnknown_Base(c_void_p, metaclass=_cominterface_meta):
        """This is workaround to avoid false-positive of static type checking.

        `IUnknown` behaves as a ctypes type, and `POINTER` can take it.
        This behavior is defined by some metaclasses in runtime.

        In runtime, this symbol in the namespace is just alias for
        `builtins.object`.
        """

        __com_QueryInterface: Callable[[Any, Any], int]
        __com_AddRef: Callable[[], int]
        __com_Release: Callable[[], int]

else:
    _IUnknown_Base = object


class IUnknown(_IUnknown_Base, metaclass=_cominterface_meta):
    """The most basic COM interface.

    Each subclasses of IUnknown must define these class attributes:

    _iid_ - a GUID instance defining the identifier of this interface

    _methods_ - a list of methods for this interface.

    The _methods_ list must in VTable order.  Methods are specified
    with STDMETHOD or COMMETHOD calls.
    """

    _case_insensitive_: ClassVar[bool] = False
    _iid_: ClassVar[GUID] = GUID("{00000000-0000-0000-C000-000000000046}")
    _methods_: ClassVar[List[_ComMemberSpec]] = [
        STDMETHOD(HRESULT, "QueryInterface", [POINTER(GUID), POINTER(c_void_p)]),
        STDMETHOD(c_ulong, "AddRef"),
        STDMETHOD(c_ulong, "Release"),
    ]

    # NOTE: Why not `QueryInterface(T) -> _Pointer[T]`?
    # Any static type checkers is not able to provide members of `T` from `_Pointer[T]`,
    # regardless of the pointer is able to access members of contents in runtime.
    # And if `isinstance(p, POINTER(T))` is `True`, then `isinstance(p, T)` is also `True`.
    # So returning `T` is not a lie, and good way to know what members the class has.
    def QueryInterface(
        self, interface: Type[_T_IUnknown], iid: Optional[GUID] = None
    ) -> _T_IUnknown:
        """QueryInterface(interface) -> instance"""
        p = POINTER(interface)()
        if iid is None:
            iid = interface._iid_
        self.__com_QueryInterface(byref(iid), byref(p))
        clsid = self.__dict__.get("__clsid")
        if clsid is not None:
            p.__dict__["__clsid"] = clsid
        return p  # type: ignore

    # these are only so that they get a docstring.
    # XXX There should be other ways to install a docstring.
    def AddRef(self) -> int:
        """Increase the internal refcount by one and return it."""
        return self.__com_AddRef()

    def Release(self) -> int:
        """Decrease the internal refcount by one and return it."""
        return self.__com_Release()


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
