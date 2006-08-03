import ctypes
import comtypes
import comtypes.automation
import comtypes.typeinfo

################################################################

# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

# XXX This section is unneeded and wrong - no need for a function prototype to
# have paramflags as instance variable.  Should be removed as soon as
# comtypes.custom is rewritten.
from _ctypes import CFuncPtr as _CFuncPtr, FUNCFLAG_STDCALL as _FUNCFLAG_STDCALL
from ctypes import _win_functype_cache

# For backward compatibility, the signature of WINFUNCTYPE cannot be
# changed, so we have to add this - which is basically the same, but
# allows to specify parameter flags from the win32 PARAMFLAGS
# enumeration.  Maybe later we have to add optional default parameter
# values and parameter names as well.
def COMMETHODTYPE(restype, argtypes, paramflags):
    flags = paramflags
    try:
        return _win_functype_cache[(restype, argtypes, flags)]
    except KeyError:
        class WinFunctionType(_CFuncPtr):
            _argtypes_ = argtypes
            _restype_ = restype
            _flags_ = _FUNCFLAG_STDCALL
            parmflags = flags
        _win_functype_cache[(restype, argtypes, flags)] = WinFunctionType
        return WinFunctionType

# -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*


def get_type(ti, tdesc):
    # Return a ctypes type for a typedesc
    if tdesc.vt == 26:
        return ctypes.POINTER(get_type(ti, tdesc._.lptdesc[0]))
    if tdesc.vt == 29:
        ta = ti.GetRefTypeInfo(tdesc._.hreftype).GetTypeAttr()
        typ = {0: ctypes.c_int, # TKIND_ENUM
               # XXX ? TKIND_RECORD, TKIND_MODULE, TKIND_COCLASS, TKIND_ALIAS, TKIND_UNION
               3: comtypes.IUnknown, # TKIND_INTERFACE
               4: comtypes.automation.IDispatch # TKIND_DISPATCH
               }[ta.typekind]
        return typ
    return comtypes.automation.VT2CTYPE[tdesc.vt]

def param_info(ti, ed):
    # Return a tuple (flags, type) or (flags, type, default)
    # from an ELEMDESC instance.
    flags = ed._.paramdesc.wParamFlags
    typ = get_type(ti, ed.tdesc)
    if flags & comtypes.typeinfo.PARAMFLAG_FHASDEFAULT:
        var = ed._.paramdesc.pparamdescex[0].varDefaultValue
        return flags, typ, var.value
    return flags, typ

def method_proto(name, ti, fd):
    # return a function prototype with parmflags: an instance of COMMETHODTYPE
    assert fd.funckind == comtypes.typeinfo.FUNC_PUREVIRTUAL, fd.funckind # FUNC_PUREVIRTUAL
    assert fd.callconv == comtypes.typeinfo.CC_STDCALL, fd.callconv
##    names = ti.GetNames(fd.memid, fd.cParams + 1)
    restype = param_info(ti, fd.elemdescFunc)[1] # result type of com method
    argtypes = [] # argument types of com method
    parmflags = []
    for i in range(fd.cParams):
        # flags, typ[, default]
        flags, typ = param_info(ti, fd.lprgelemdescParam[i])[:2]
        argtypes.append(typ)
        parmflags.append(flags)
    proto = COMMETHODTYPE(restype, tuple(argtypes), tuple(parmflags))
    return proto

def get_custom_interface(comobj, typeinfo=None):
    # return comobj, typeinfo where typeinfo describes a custom
    # interface, and comobj implements it
    if typeinfo is None:
        # If the caller didn't specify typeinfo, and comobj doesn't
        # provide any, we're stumped - and the caller should handle
        # whatever exception this raises.
        typeinfo = comobj.GetTypeInfo()
    ta = typeinfo.GetTypeAttr()
    if ta.typekind == comtypes.typeinfo.TKIND_INTERFACE:
        # correct typeinfo, still need to QI for this interface
        iid = ta.guid
    elif ta.typekind == comtypes.typeinfo.TKIND_DISPATCH:
        # try to get the dual interface portion from a dispatch interface
        href = typeinfo.GetRefTypeOfImplType(-1)
        typeinfo = typeinfo.GetRefTypeInfo(href)
        ta = typeinfo.GetTypeAttr()
        if ta.typekind != comtypes.typeinfo.TKIND_INTERFACE:
            # it didn't work
            raise TypeError, "could not get custom interface"
        iid = ta.guid
    else:
        raise TypeError, "could not get custom interface"
    # XXX We should make sure that our custom interface derives from IDispatch,
    # if not, we need to specify IUnknown in the following call.
    comobj = comobj.QueryInterface(comtypes.automation.IDispatch, iid)
    return comobj, typeinfo

################

PDisp = ctypes.POINTER(comtypes.automation.IDispatch)
PUnk = ctypes.POINTER(comtypes.IUnknown)

def _wrap(obj):
    if isinstance(obj, PDisp):
        return _Dynamic(obj)
##    if isinstance(obj, PUnk):
##        return _Dynamic(obj)
    return obj

def _unwrap(obj):
    return getattr(obj, "_comobj", obj)

class _ComInvoker(object):
    def __init__(self, comobj, name, caller=None, getter=None, setter=None):
        self.comobj = comobj
        self.name = name
        self.caller = caller
        self.getter = getter
        self.setter = setter

    def __repr__(self):
        return "<_ComInvoker %s (call=%s, get=%s, set=%s)>" % \
               (self.name, self.caller, self.getter, self.setter)

    def _do_invoke(self, comcall, args, kw):
        args = [_unwrap(a) for a in args]
        retvals = []
        for index, x in enumerate(comcall.parmflags):
            if x & 2: # PARAMFLAG_FOUT
                typ = comcall.argtypes[index]._type_
                inst = typ()
                args = args[:index] + [ctypes.byref(inst)] + args[index:]
                retvals.append(inst)
        result = comcall(self.comobj, *args, **kw)
        if retvals:
            result = tuple([_wrap(x.value) for x in retvals])
            if len(result) == 1:
                result = result[0]
        return result

    def __call__(self, *args, **kw):
        if not self.caller:
            raise TypeError, "uncallable object"
        return self._do_invoke(self.caller, args, kw)

    def __getitem__(self, index):
        if not self.getter:
            raise TypeError, "unindexable object"
        if isinstance(index, tuple):
            return self._do_invoke(self.getter, index, {})
        return self._do_invoke(self.getter, (index,), {})

    def __setitem__(self, index, value):
        if not self.setter:
            raise TypeError, "object doesn't support item assignment"
        if isinstance(index, tuple):
            self._do_invoke(self.setter, index + (value,), {})
        else:
            self._do_invoke(self.setter, (index, value), {})

    def set(self, value):
        self._do_invoke(self.setter, (value,), {})

    def get(self):
        if not self.getter:
            return self
        if 1 in self.getter.parmflags:
            return self
        return self._do_invoke(self.getter, (), {})

from comtypes.automation import INVOKE_FUNC as _FUNC, INVOKE_PROPERTYGET as _PROPGET, \
     INVOKE_PROPERTYPUT as _PROPPUT

################################################################

class _Dynamic(object):
    _typecomp = _typeinfo = _comobj = _commap = None # pychecker
    
    def __init__(self, comobj, typeinfo=None):
        comobj, typeinfo = get_custom_interface(comobj, typeinfo)
        self.__dict__["_typeinfo"] = typeinfo
        self.__dict__["_typecomp"] = typeinfo.GetTypeComp()
        self.__dict__["_comobj"] = comobj
        self.__dict__["_commap"] = {}
        
    def __getattr__(self, name):
        invoker = self.__get_invoker(name, _FUNC | _PROPGET | _PROPPUT)
        return invoker.get()

    def __setattr__(self, name, value):
        prop = self.__get_invoker(name, _PROPPUT)
        prop.set(value)

    def __get_invoker(self, name, what):
        try:
            return self._commap[(name, what)]
        except KeyError:
            prop = self.__create_invoker(name, what)
            self._commap[(name, what)] = prop
            return prop

    def __create_invoker(self, name, what):
        comcalls = {}
        for invkind in (_FUNC, _PROPGET, _PROPPUT):
            if what and invkind:
                try:
                    kind, desc = self._typecomp.Bind(name, invkind)
                except WindowsError:
                    continue
                assert desc.invkind == invkind
                if kind != "function":
                    raise "NYI" # can this happen?
                proto = method_proto(name, self._typeinfo, desc)
                comcalls[invkind] = proto(desc.oVft/4, name)
        return _ComInvoker(self._comobj, name,
                           caller=comcalls.get(_FUNC),
                           getter=comcalls.get(_PROPGET),
                           setter=comcalls.get(_PROPPUT))
                
        

################################################################

def ActiveXObject(progid,
                  interface=comtypes.automation.IDispatch,
                  clsctx=comtypes.CLSCTX_ALL):
    # XXX Should we also allow GUID instances for the first parameter?
    # Or strings containing guids?
    clsid = comtypes.GUID.from_progid(progid)
    p = comtypes.CoCreateInstance(clsid, interface=interface, clsctx=clsctx)
    return _Dynamic(p)

if __name__ == "__main__":
    f = ActiveXObject("InternetExplorer.Application")
    for n in "Name Visible Silent Offline ReadyState AddressBar Resizable".split():
        print n, getattr(f, n)
    f.Visible = True
##    print f.ClientToWindow(1, 2)
    import time
##    time.sleep(0.5)
    f.Quit()
