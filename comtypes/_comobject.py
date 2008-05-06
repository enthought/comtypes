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

# so we don't have to import comtypes.automation
DISPATCH_METHOD = 1
DISPATCH_PROPERTYGET = 2
DISPATCH_PROPERTYPUT = 4
DISPATCH_PROPERTYPUTREF = 8

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
    # An argument is an input arg either if flags are NOT set in the
    # idl file, or if the flags contain 'in'. In other words, the
    # direction flag is either exactly '0' or has the '1' bit set:
    args_in = len([f for f in dirflags if (f == 0) or (f & 1)])
    # number of output arguments:
    args_out = len([f for f in dirflags if f & 2])
    ## XXX Remove this:
##    if args_in != code.co_argcount - 1:
##        return catch_errors(inst, mth, interface, mthname)
    # This code assumes that input args are always first, and output
    # args are always last.  Have to check with the IDL docs if this
    # is always correct.

    clsid = getattr(inst, "_reg_clsid_", None)
    def wrapper(this, *args):
        outargs = args[len(args)-args_out:]
        # Method implementations could check for and return E_POINTER
        # themselves.  Or an error will be raised when
        # 'outargs[i][0] = value' is executed.
##        for a in outargs:
##            if not a:
##                return E_POINTER
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

    def get_impl(self, interface, mthname, paramflags, idlflags):
        mth = self.find_impl(interface, mthname, paramflags, idlflags)
        if mth is None:
            return _do_implement(interface.__name__, mthname)
        return hack(self.inst, mth, paramflags, interface, mthname)

    def find_method(self, fq_name, mthname):
        # Try to find a method, first with the fully qualified name
        # ('IUnknown_QueryInterface'), if that fails try the simple
        # name ('QueryInterface')
        try:
            return getattr(self.inst, fq_name)
        except AttributeError:
            pass
        return getattr(self.inst, mthname)

    def find_impl(self, interface, mthname, paramflags, idlflags):
        fq_name = "%s_%s" % (interface.__name__, mthname)
        if interface._case_insensitive_:
            # simple name, like 'QueryInterface'
            mthname = self.names.get(mthname.lower(), mthname)
            # qualified name, like 'IUnknown_QueryInterface'
            fq_name = self.names.get(fq_name.lower(), fq_name)

        try:
            return self.find_method(fq_name, mthname)
        except AttributeError:
            pass
        propname = mthname[5:] # strip the '_get_' or '_set' prefix
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
        self = super(COMObject, cls).__new__(cls)
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
        finder = self._get_method_finder_(itf)
        for interface in itf.__mro__[-2::-1]:
            iids.append(interface._iid_)
            for m in interface._methods_:
                restype, mthname, argtypes, paramflags, idlflags, helptext = m
                proto = WINFUNCTYPE(restype, c_void_p, *argtypes)
                fields.append((mthname, proto))
                mth = finder.get_impl(interface, mthname, paramflags, idlflags)
                methods.append(proto(mth))
        Vtbl = _create_vtbl_type(tuple(fields), itf)
        vtbl = Vtbl(*methods)
        for iid in iids:
            self._com_pointers_[iid] = pointer(pointer(vtbl))
        if hasattr(itf, "_disp_methods_"):
            self._dispimpl_ = {}
        for m in getattr(itf, "_disp_methods_", ()):
            what, mthname, idlflags, restype, argspec = m
            #################
            # What we have:
            #
            # restypes is a ctypes type or None
            # argspec is seq. of (['in'], paramtype, paramname) tuples (or lists?)
            #################
            # What we need:
            #
            # idlflags must contain 'propget', 'propset' and so on:
            # Must be constructed by converting disptype
            #
            # paramflags must be a sequence
            # of (F_IN|F_OUT|F_RETVAL, paramname[, default-value]) tuples
            # comtypes has this function which may help:
            #    def _encode_idl(names):
            #        # sum up "in", "out", ... values found in _PARAMFLAGS, ignoring all others.
            #        return sum([_PARAMFLAGS.get(n, 0) for n in names])
            #################
            dispid = idlflags[0] # XXX can the dispid be at a different index?  Check codegenerator.
            if what == "DISPMETHOD":
                if 'propget' in idlflags:
                    invkind = 2 # DISPATCH_PROPERTYGET
                    mthname = "_get_" + mthname
                elif 'propput' in idlflags:
                    invkind = 4 # DISPATCH_PROPERTYPUT
                    mthname = "_set_" + mthname
                elif 'propputref' in idlflags:
                    invkind = 8 # DISPATCH_PROPERTYPUTREF
                    mthname = "_setref_" + mthname
                else:
                    invkind = 1 # DISPATCH_METHOD
                    if restype:
                        argspec = argspec + ((['out'], restype, ""),)
            elif what == "DISPPROPERTY":
                import sys
                # has get and (set, if not "readonly" in idlflags)
                print >> sys.stderr, "Not yet Implemented", mthname
##                import pdb; pdb.set_trace()
                invkind = 0
                pass

            from comtypes import _encode_idl
            paramflags = [((_encode_idl(x[0]), x[1]) + tuple(x[3:])) for x in argspec]

##            import sys
##            print >> sys.stderr, "GET_IMPL", interface.__name__, mthname
##            print >> sys.stderr, "\tparamflags:", paramflags
##            print >> sys.stderr, "\tidlflags:", idlflags
##            print >> sys.stderr, finder.get_impl(interface, mthname, paramflags, idlflags)
##            print >> sys.stderr
##            print >> sys.stderr, "%s_%s" % (interface.__name__, mthname)
            # We should build a _dispmap_ (or whatever) now, that maps
            # invkind and dispid to implementations that the finder finds;
            # and eventually calls them in IDispatch_Invoke.
            impl = finder.get_impl(interface, mthname, paramflags, idlflags)
            self._dispimpl_[(dispid, invkind)] = impl

    def _get_method_finder_(self, itf):
        # This method can be overridden to customize how methods are
        # found.
        return _MethodFinder(self)

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

    def IDispatch_GetTypeInfo(self, this, itinfo, lcid, ptinfo):
        if itinfo != 0:
            return DISP_E_BADINDEX
        try:
            ptinfo[0] = self.__typeinfo
            return S_OK
        except AttributeError:
            return E_NOTIMPL

    def IDispatch_GetIDsOfNames(self, this, riid, rgszNames, cNames, lcid, rgDispId):
        # This call uses windll instead of oledll so that a failed
        # call to DispGetIDsOfNames will return a HRESULT instead of
        # raising an error.
        try:
            tinfo = self.__typeinfo
        except AttributeError:
            return E_NOTIMPL
        return windll.oleaut32.DispGetIDsOfNames(tinfo,
                                                 rgszNames, cNames, rgDispId)

    def IDispatch_Invoke(self, this, dispIdMember, riid, lcid, wFlags,
                         pDispParams, pVarResult, pExcepInfo, puArgErr):
        try:
            self._dispimpl_
        except AttributeError:
            try:
                tinfo = self.__typeinfo
            except AttributeError:
                # Hm, we pretend to implement IDispatch, but have no
                # typeinfo, and so cannot fulfill the contract.  Should we
                # better return E_NOTIMPL or DISP_E_MEMBERNOTFOUND?  Some
                # clients call IDispatch_Invoke with 'known' DISPID_...'
                # values, without going through GetIDsOfNames first.
                return DISP_E_MEMBERNOTFOUND
            # This call uses windll instead of oledll so that a failed
            # call to DispInvoke will return a HRESULT instead of raising
            # an error.
            interface = self._com_interfaces_[0]
            ptr = self._com_pointers_[interface._iid_]
            return windll.oleaut32.DispInvoke(ptr,
                                              tinfo,
                                              dispIdMember, wFlags, pDispParams,
                                              pVarResult, pExcepInfo, puArgErr)

        try:
            # XXX Hm, wFlags should be considered a SET of flags...
            mth = self._dispimpl_[(dispIdMember, wFlags)]
        except KeyError:
            return DISP_E_MEMBERNOTFOUND

        params = pDispParams[0]
        args = [params.rgvarg[i].value for i in range(params.cArgs)[::-1]]
        if pVarResult:
            args += [pVarResult]
        return mth(this, *args)

##        # No typeinfo, or a non-dual dispinterface.  We have to
##        # implement Invoke completely ourself.
##        impl = self._find_impl(dispIdMember, wFlags, bool(pVarResult))
##        # _find_impl could return an integer error code which we return. 
##        if isinstance(impl, (int, long)):
##            return impl

##        # _find_impl returned a callable; prepare the arguments and
##        # call it.
##        params = pDispParams[0]
##        args = [params.rgvarg[i].value for i in range(params.cArgs)[::-1]]
##        if pVarResult:
##            args += [pVarResult]
##        return impl(this, *args)

##    def _find_impl(self, dispid, wFlags, expects_result,
##                   finder=None):
##        # This method tries to find an implementation for dispid and
##        # wFlags.  If not found, an integer HRESULT error code is
##        # returned; otherwise a function/method that Invoke must call.
##        try:
##            return self._dispimpl_[(dispid, wFlags)]
##        except KeyError:
##            pass

##        interface = self._com_interfaces_[0]

##        methods = interface._disp_methods_
##        # XXX This uses a linear search
##        descr = [m for m in methods
##                 if m[2][0] == dispid]
##        if not descr:
##            self._dispimpl_[(dispid, wFlags)] = DISP_E_MEMBERNOTFOUND
##            return DISP_E_MEMBERNOTFOUND
##        disptype, name, idlflags, restype, argspec = descr[0]

##        if disptype == "DISPMETHOD":
##            if (wFlags & DISPATCH_METHOD) == 0:
##                self._dispimpl_[(dispid, wFlags)] = DISP_E_MEMBERNOTFOUND
##                return DISP_E_MEMBERNOTFOUND

##        elif disptype == "DISPPROPERTY":

##            if wFlags & DISPATCH_PROPERTYGET:
##                name = "_get_" + name
##            elif wFlags & DISPATCH_PROPERTYPUT:
##                name = "_set_" + name
##            elif wFlags & DISPATCH_PROPERTYPUTREF:
##                name = "_setref_" + name
##            else:
##                self._dispimpl_[(dispid, wFlags)] = DISP_E_MEMBERNOTFOUND
##                return DISP_E_MEMBERNOTFOUND

##        else:
##            # this should not happen at all: it is a bug in comtypes
##            return E_FAIL

##        from comtypes import _encode_idl
##        paramflags = [(_encode_idl(m[0]),) + m[1:] for m in argspec]
##        if expects_result:
##            paramflags += [[2]]
            
##        if finder is None:
##            finder = _MethodFinder(self)
##        impl = finder.get_impl(interface, name, paramflags, [])
##        self._dispimpl_[(dispid, wFlags)] = impl
##        return impl

    ################################################################
    # IPersist interface
    def IPersist_GetClassID(self):
        return self._reg_clsid_

__all__ = ["COMObject"]
