from ctypes import *
from comtypes.hresult import *

import os
import new
import logging
logger = logging.getLogger(__name__)
_debug = logger.debug
_warning = logger.warning
_error = logger.error

################################################################
# COM object implementation
from _ctypes import CopyComPointer

from comtypes import COMError, ReturnHRESULT
from comtypes.errorinfo import ISupportErrorInfo, ReportException, ReportError
from comtypes.typeinfo import IProvideClassInfo, IProvideClassInfo2
from comtypes import IPersist

class E_NotImplemented(Exception):
    """COM method is not implemented"""

def HRESULT_FROM_WIN32(errcode):
    "Convert a Windows error code into a HRESULT value."
    if errcode is None:
        return 0x80000000
    if errcode & 0x80000000:
        return errcode
    return (errcode & 0xFFFF) | 0x80070000

def winerror(exc):
    """Return the windows error code from a WindowsError or COMError
    instance."""
    try:
        code = exc[0]
        if isinstance(code, (int, long)):
            return code
    except IndexError:
        pass
    # Sometimes, a WindowsError instance has no error code.  An access
    # violation raised by ctypes has only text, for example.  In this
    # cases we return a generic error code.
    return E_FAIL

def _do_implement(interface_name, method_name):
    def _not_implemented(*args):
        """Return E_NOTIMPL because the method is not implemented."""
        _debug("unimplemented method %s_%s called", interface_name, method_name)
        return E_NOTIMPL
    return _not_implemented

def catch_errors(obj, mth, interface, mthname):
    clsid = getattr(obj, "_reg_clsid_", None)
    def func(*args, **kw):
        try:
            return mth(*args, **kw)
        except ReturnHRESULT, (hresult, text):
            return ReportError(text, iid=interface._iid_, clsid=clsid, hresult=hresult)
        except (COMError, WindowsError), details:
            _error("Exception in %s.%s implementation:", interface.__name__, mthname, exc_info=True)
            return HRESULT_FROM_WIN32(winerror(details))
        except E_NotImplemented:
            _warning("Unimplemented method %s.%s called", interface.__name__, mthname)
            return E_NOTIMPL
        except:
            _error("Exception in %s.%s implementation:", interface.__name__, mthname, exc_info=True)
            return ReportException(E_FAIL, interface._iid_, clsid=clsid)
    return func

################################################################

def hack(inst, mth, paramflags, interface, mthname):
    if paramflags is None:
        return catch_errors(inst, mth, interface, mthname)
    code = mth.func_code
    if code.co_varnames[1:2] == ("this",):
        return catch_errors(inst, mth, interface, mthname)
    dirflags = [f[0] for f in paramflags]
    # An argument is [IN] if it is not [OUT] !
    # This handles the case where no direction is defined in the IDL file.
    # number of input arguments:
    args_in = len([f for f in dirflags if (f & 2) == 0])
    # number of output arguments:
    args_out = len([f for f in dirflags if f & 2])
    if args_in != code.co_argcount - 1:
        return catch_errors(inst, mth, interface, mthname)
    # This code assumes that input args are always first, and output
    # args are always last.  Have to check with the IDL docs if this
    # is always correct.

    clsid = getattr(inst, "_reg_clsid_", None)
    def wrapper(this, *args):
        outargs = args[len(args)-args_out:]
        for a in outargs:
            if not a:
                return E_POINTER
        try:
            result = mth(*args[:args_in])
            if args_out == 1:
                outargs[0][0] = result
            elif args_out != 0:
                if len(result) != args_out:
                    raise ValueError("Method should have returned a %s-tuple" % args_out)
                for i, value in enumerate(result):
                    outargs[i][0] = value
        except ReturnHRESULT, (hresult, text):
            return ReportError(text, iid=interface._iid_, clsid=clsid, hresult=hresult)
        except COMError, (hr, text, details):
            _error("Exception in %s.%s implementation:", interface.__name__, mthname, exc_info=True)
            try:
                descr, source, helpfile, helpcontext, progid = details
            except (ValueError, TypeError):
                msg = str(details)
            else:
                msg = "%s: %s" % (source, descr)
            hr = HRESULT_FROM_WIN32(hr)
            return ReportError(msg, iid=interface._iid_, clsid=clsid, hresult=hr)
        except WindowsError, details:
            _error("Exception in %s.%s implementation:", interface.__name__, mthname, exc_info=True)
            hr = HRESULT_FROM_WIN32(winerror(details))
            return ReportException(hr, interface._iid_, clsid=clsid)
        except E_NotImplemented:
            _warning("Unimplemented method %s.%s called", interface.__name__, mthname)
            return E_NOTIMPL
        except:
            _error("Exception in %s.%s implementation:", interface.__name__, mthname, exc_info=True)
            return ReportException(E_FAIL, interface._iid_, clsid=clsid)
        return S_OK

    return wrapper

class _MethodFinder(object):
    def __init__(self, inst):
        self.inst = inst
        # map lower case names to names with correct spelling.
        self.names = dict([(n.lower(), n) for n in dir(inst)])

    def get_impl(self, instance, interface, mthname, paramflags, idlflags):
        mth = self.find_impl(interface, mthname, paramflags, idlflags)
        if mth is None:
            return _do_implement(interface.__name__, mthname)
        return hack(self.inst, mth, paramflags, interface, mthname)

    def find_impl(self, interface, mthname, paramflags, idlflags):
        fq_name = "%s_%s" % (interface.__name__, mthname)
        if interface._case_insensitive_:
            mthname = self.names.get(mthname.lower(), mthname)
            fq_name = self.names.get(fq_name.lower(), fq_name)

        try:
            # qualified name, like 'IUnknown_QueryInterface'
            return getattr(self.inst, fq_name)
        except AttributeError:
            pass
        try:
            # simple name, like 'QueryInterface'
            return getattr(self.inst, mthname)
        except AttributeError:
            pass
        propname = mthname[5:]
        if interface._case_insensitive_:
            propname = self.names.get(propname.lower(), propname)
        # propput and propget is done with 'normal' attribute access,
        # but only for COM properties that do not take additional
        # arguments:

        if "propget" in idlflags and len(paramflags) == 1:
            return self.getter(propname)
        if "propput" in idlflags and len(paramflags) == 1:
            return self.setter(propname)
        _debug("%r: %s.%s not implemented", self.inst, interface.__name__, mthname)
        return None

    def setter(self, propname):
        #
        def set(self, value):
            try:
                # XXX this may not be correct is the object implements
                # _get_PropName but not _set_PropName
                setattr(self, propname, value)
            except AttributeError:
                raise E_NotImplemented()
        return new.instancemethod(set, self.inst, type(self.inst))

    def getter(self, propname):
        #
        def get(self):
            try:
                return getattr(self, propname)
            except AttributeError:
                raise E_NotImplemented()
        return new.instancemethod(get, self.inst, type(self.inst))

def _create_vtbl_type(fields, itf):
    try:
        return _vtbl_types[fields]
    except KeyError:
        class Vtbl(Structure):
            _fields_ = fields
        Vtbl.__name__ = "Vtbl_%s" % itf.__name__
        _vtbl_types[fields] = Vtbl
        return Vtbl

# Ugh. Another type cache to avoid leaking types.
_vtbl_types = {}

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
            self.__prepare_comobject()
        return self

    def __prepare_comobject(self):
        # When a CoClass instance is created, COM pointers to all
        # interfaces are created.  Also, the CoClass must be kept alive as
        # until the COM reference count drops to zero, even if no Python
        # code keeps a reference to the object.
        #
        # The _com_pointers_ instance variable maps string interface iids
        # to C compatible COM pointers.
        self._com_pointers_ = {}
        # COM refcount starts at zero.
        self._refcnt = c_long(0)

        # Some interfaces have a default implementation in COMObject:
        # - ISupportErrorInfo
        # - IPersist (if the subclass has a _reg_clsid_ attribute)
        # - IProvideClassInfo (if the subclass has a _reg_clsid_ attribute)
        # - IProvideClassInfo2 (if the subclass has a _outgoing_interfaces_ attribute)
        #
        # Add these if they are not listed in _com_interfaces_.
        interfaces = tuple(self._com_interfaces_)
        if ISupportErrorInfo not in interfaces:
            interfaces += (ISupportErrorInfo,)
        if hasattr(self, "_reg_typelib_"):
            from comtypes.typeinfo import LoadRegTypeLib
            self._COMObject__typelib = LoadRegTypeLib(*self._reg_typelib_)
            if hasattr(self, "_reg_clsid_"):
                if IProvideClassInfo not in interfaces:
                    interfaces += (IProvideClassInfo,)
                if hasattr(self, "_outgoing_interfaces_") and \
                   IProvideClassInfo2 not in interfaces:
                    interfaces += (IProvideClassInfo2,)
        if hasattr(self, "_reg_clsid_"):
                if IPersist not in interfaces:
                    interfaces += (IPersist,)
        for itf in interfaces[::-1]:
            self.__make_interface_pointer(itf)

    def __make_interface_pointer(self, itf):
        methods = [] # method implementations
        fields = [] # (name, prototype) for virtual function table
        iids = [] # interface identifiers.
        # iterate over interface inheritance in reverse order to build the
        # virtual function table, and leave out the 'object' base class.
        finder = _MethodFinder(self)
        for interface in itf.__mro__[-2::-1]:
            iids.append(interface._iid_)
            for m in interface._methods_:
                restype, mthname, argtypes, paramflags, idlflags, helptext = m
                proto = WINFUNCTYPE(restype, c_void_p, *argtypes)
                fields.append((mthname, proto))
                mth = finder.get_impl(self, interface, mthname, paramflags, idlflags)
                methods.append(proto(mth))
        Vtbl = _create_vtbl_type(tuple(fields), itf)
        vtbl = Vtbl(*methods)
        for iid in iids:
            self._com_pointers_[iid] = pointer(pointer(vtbl))

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

    def QueryInterface(self, interface):
        "Query the object for an interface pointer"
        # This method is NOT the implementation of
        # IUnknown::QueryInterface, instead it is supposed to be
        # called on an COMObject by user code.  It allows to get COM
        # interface pointers from COMObject instances.
        ptr = self._com_pointers_.get(interface._iid_, None)
        if ptr is None:
            raise COMError(E_NOINTERFACE, FormatError(E_NOINTERFACE),
                           (None, None, 0, None, None))
        # CopyComPointer(src, dst) calls AddRef!
        result = POINTER(interface)()
        CopyComPointer(ptr, byref(result))
        return result

    ################################################################
    # ISupportErrorInfo::InterfaceSupportsErrorInfo implementation
    def ISupportErrorInfo_InterfaceSupportsErrorInfo(self, this, riid):
        if riid[0] in self._com_pointers_:
            return S_OK
        return S_FALSE

    ################################################################
    # IProvideClassInfo::GetClassInfo implementation
    def IProvideClassInfo_GetClassInfo(self):
        try:
            self.__typelib
        except AttributeError:
            raise WindowsError(E_NOTIMPL)
        return self.__typelib.GetTypeInfoOfGuid(self._reg_clsid_)

    ################################################################
    # IProvideClassInfo2::GetGUID implementation

    def IProvideClassInfo2_GetGUID(self, dwGuidKind):
        # GUIDKIND_DEFAULT_SOURCE_DISP_IID = 1
        if dwGuidKind != 1:
            raise WindowsError(E_INVALIDARG)
        return self._outgoing_interfaces_[0]._iid_

    ################################################################
    # IDispatch methods
    @property
    def __typeinfo(self):
        # XXX Looks like this better be a static property, set by the
        # code that sets __typelib also...
        iid = self._com_interfaces_[0]._iid_
        return self.__typelib.GetTypeInfoOfGuid(iid)

    def IDispatch_GetTypeInfoCount(self):
        try:
            self.__typelib
        except AttributeError:
            return 0
        else:
            return 1

    def IDispatch_GetTypeInfo(self, itinfo, lcid):
        if itinfo != 0:
            raise WindowsError(DISP_E_BADINDEX)
        try:
            self.__typelib
        except AttributeError:
            raise WindowsError(E_NOTIMPL)
        else:
            return self.__typeinfo

    def IDispatch_GetIDsOfNames(self, this, riid, rgszNames, cNames, lcid, rgDispId):
        # Use windll to let DispGetIDsOfNames return a HRESULT instead
        # of raising an error:
        try:
            self.__typeinfo
        except AttributeError:
            return E_NOTIMPL
        return windll.oleaut32.DispGetIDsOfNames(self.__typeinfo,
                                                 rgszNames, cNames, rgDispId)

    def IDispatch_Invoke(self, this, dispIdMember, riid, lcid, wFlags,
                         pDispParams, pVarResult, pExcepInfo, puArgErr):
        try:
            self.__typeinfo
        except AttributeError:
            # Hm, we pretend to implement IDispatch, but have no
            # typeinfo, and so cannot fulfill the contract.  Should we
            # better return E_NOTIMPL or DISP_E_MEMBERNOTFOUND?  Some
            # clients call IDispatch_Invoke with 'known' DISPID_...'
            # values, without going through GetIDsOfNames first.
            return DISP_E_MEMBERNOTFOUND
        impl = self._com_pointers_[self._com_interfaces_[0]._iid_]
        # Use windll to let DispInvoke return a HRESULT instead
        # of raising an error:
        return windll.oleaut32.DispInvoke(impl,
                                          self.__typeinfo,
                                          dispIdMember, wFlags, pDispParams,
                                          pVarResult, pExcepInfo, puArgErr)

    ################################################################
    # IPersist interface
    def IPersist_GetClassID(self):
        return self._reg_clsid_

__all__ = ["COMObject"]
