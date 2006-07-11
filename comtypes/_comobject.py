from ctypes import *
from comtypes.hresult import *

import os
import logging
logger = logging.getLogger(__name__)
_debug = logger.debug
_warning = logger.warning

################################################################
# COM object implementation
from _ctypes import CopyComPointer

def prepare_comobject(inst):
    # When a CoClass instance is created, COM pointers to all
    # interfaces are created.  Also, the CoClass must be kept alive as
    # until the COM reference count drops to zero, even if no Python
    # code keeps a reference to the object.
    #
    # The _com_pointers_ instance variable maps string interface iids
    # to C compatible COM pointers.
    inst._com_pointers_ = {}
    # COM refcount starts at zero.
    inst._refcnt = c_long(0)
    for itf in inst._com_interfaces_[::-1]:
        make_interface_pointer(inst, itf)

from comtypes.errorinfo import ReportException

def catch_errors(obj, mth, interface):
    iid = interface._iid_
    clsid = getattr(obj, "_reg_clsid_", None)
    def func(*args, **kw):
        try:
            return mth(*args, **kw)
        except Exception:
            _warning("%s", interface, exc_info=True)
            return ReportException(E_FAIL, iid, clsid=clsid)
    return func

def _do_implement(interface_name, method_name):
    def mth(*args):
        _debug("unimplemented method %s_%s called", interface_name, method_name)
        return E_NOTIMPL
    return mth

def make_interface_pointer(inst, itf,
                           _debug=_debug):
    methods = [] # method implementations
    fields = [] # (name, prototype) for virtual function table
    iids = [] # interface identifiers.
    # iterate over interface inheritance in reverse order to build the
    # virtual function table, and leave out the 'object' base class.
    for interface in itf.__mro__[-2::-1]:
        iids.append(interface._iid_)
        for m in interface._methods_:
            restype, mthname, argtypes, paramflags, idlflags, helptext = m
            proto = WINFUNCTYPE(restype, c_void_p, *argtypes)
            fields.append((mthname, proto))
            # try the simple name, like 'QueryInterface'
            mth = getattr(inst, mthname, None)
            if mth is None:
                # qualified name, like 'IUnknown_QueryInterface'
                mth = getattr(inst, "%s_%s" % (interface.__name__, mthname), None)
            if mth is None:
                mth = _do_implement(interface.__name__, mthname)
                # XXX Should we try harder for case-insensitive lookup of methods?
                _debug("%r: %s.%s not implemented", inst, interface.__name__, mthname)
            else:
                mth = catch_errors(inst, mth, interface)
            methods.append(proto(mth))
    class Vtbl(Structure):
        _fields_ = fields
    Vtbl.__name__ = "Vtbl_%s" % itf.__name__
    vtbl = Vtbl(*methods)
    for iid in iids:
        inst._com_pointers_[iid] = pointer(pointer(vtbl))

################################################################

if os.name == "ce":
    _InterlockedIncrement = windll.coredll.InterlockedIncrement
    _InterlockedDecrement = windll.coredll.InterlockedDecrement
else:
    _InterlockedIncrement = windll.kernel32.InterlockedIncrement
    _InterlockedDecrement = windll.kernel32.InterlockedDecrement

class COMObject(object):
    _instances_ = {}
    _factory = None

    def __new__(cls, *args, **kw):
        self = super(COMObject, cls).__new__(cls, *args, **kw)
        if isinstance(self, c_void_p):
            # We build the VTables only for direct instances of
            # CoClass, not for POINTERs to CoClass.
            return self
        if hasattr(self, "_com_interfaces_"):
            prepare_comobject(self)
        return self

    #########################################################
    # IUnknown methods implementations
    def IUnknown_AddRef(self, this,
                        __InterlockedIncrement=_InterlockedIncrement,
                        _debug=_debug):
        result = __InterlockedIncrement(byref(self._refcnt))
        if result == 1:
            # keep reference to the object in a class variable.
            COMObject._instances_[self] = None
            _debug("%d active COM objects: Added   %r", len(COMObject._instances_), self)
        _debug("%r.AddRef() -> %s", self, result)
        return result

    def IUnknown_Release(self, this,
                         __InterlockedDecrement=_InterlockedDecrement,
                         _byref=byref,
                        _debug=_debug):
        # If this is called at COM shutdown, byref() and
        # _InterlockedDecrement() must still be available, although
        # module level variables may have been deleted already - so we
        # supply them as default arguments.
        result = __InterlockedDecrement(_byref(self._refcnt))
        _debug("%r.Release() -> %s", self, result)
        if result == 0:
            # For whatever reasons, at cleanup it may be that
            # COMObject is already cleaned (set to None)
            try:
                del COMObject._instances_[self]
            except AttributeError:
                _debug("? active COM objects: Removed %r", self)
            else:
                _debug("%d active COM objects: Removed %r", len(COMObject._instances_), self)
            if self._factory is not None:
                self._factory.LockServer(None, 0)
        return result

    def IUnknown_QueryInterface(self, this, riid, ppvObj,
                        _debug=_debug):
        # XXX This is probably too slow.
        # riid[0].hashcode() alone takes 33 us!
        iid = riid[0]
        ptr = self._com_pointers_.get(iid, None)
        if ptr is not None:
            # CopyComPointer(src, dst) calls AddRef!
            _debug("%r.QueryInterface(%s) -> S_OK", self, iid)
            return CopyComPointer(ptr, ppvObj)
        _debug("%r.QueryInterface(%s) -> E_NOINTERFACE", self, iid)
        return E_NOINTERFACE

    ################################################################
    # ISupportErrorInfo method implementation
    def ISupportErrorInfo_InterfaceSupportsErrorInfo(self, this, riid):
        if riid[0] in self._com_pointers_:
            return S_OK
        return S_FALSE

__all__ = ["COMObject"]
