import ctypes

import comtypes
from comtypes import TYPE_CHECKING

if TYPE_CHECKING:
    from comtypes import _CData
    from typing import (
        Any, Callable, Dict, Iterator, List, Optional, Tuple, Type, Union as _UnionT
    )
    PositionalParamFlagType = Tuple[int, Optional[str]]
    OptionalParamFlagType = Tuple[int, Optional[str], Any]
    ParamFlagType = _UnionT[PositionalParamFlagType, OptionalParamFlagType]
    PositionalArgSpecElmType = Tuple[List[str], Type[_CData], str]
    OptionalArgSpecElmType = Tuple[List[str], Type[_CData], str, Any]
    ArgSpecElmType = _UnionT[PositionalArgSpecElmType, OptionalArgSpecElmType]


_PARAMFLAGS = {
    "in": 1,
    "out": 2,
    "lcid": 4,
    "retval": 8,
    "optional": 16,
    }


def _encode_idl(names):
    # sum up all values found in _PARAMFLAGS, ignoring all others.
    return sum([_PARAMFLAGS.get(n, 0) for n in names])


_NOTHING = object()


def _unpack_argspec(idl, typ, name=None, defval=_NOTHING):
    # type: (List[str], Type[_CData], Optional[str], Any) -> Tuple[List[str], Type[_CData], Optional[str], Any]
    return idl, typ, name, defval


def _resolve_argspec(items):
    # type: (Tuple[ArgSpecElmType, ...]) -> Tuple[Tuple[ParamFlagType, ...], Tuple[Type[_CData], ...]]
    """Unpacks and converts from argspec to paramflags and argtypes.

    - paramflags is a sequence of `(pflags: int, argname: str, | None[, defval: Any])`.
    - argtypes is a sequence of `type[_CData]`.
    """
    from comtypes.automation import VARIANT

    paramflags = []
    argtypes = []
    for item in items:
        idl, typ, argname, defval = _unpack_argspec(*item)  # type: ignore
        pflags = _encode_idl(idl)
        if "optional" in idl:
            if defval is _NOTHING:
                if typ is VARIANT:
                    defval = VARIANT.missing
                elif typ is ctypes.POINTER(VARIANT):
                    defval = ctypes.pointer(VARIANT.missing)
                else:
                    # msg = ("'optional' only allowed for VARIANT and VARIANT*, not for %s" % typ.__name__)
                    # warnings.warn(msg, IDLWarning, stacklevel=2)
                    defval = typ()
        if defval is _NOTHING:
            paramflags.append((pflags, argname))
        else:
            paramflags.append((pflags, argname, defval))
        argtypes.append(typ)
    return tuple(paramflags), tuple(argtypes)


class _MemberSpec(object):
    """Specifier of a slot of method or property."""
    __slots__ = ("name", "idlflags", "restype")
    def __init__(self, name, idlflags, restype):
        self.name = name  # type: str
        self.idlflags = idlflags  # type: Tuple[_UnionT[str, int], ...]
        self.restype = restype  # type: Optional[Type[_CData]]

    def is_prop(self):
        # type: () -> bool
        propflags = ("propget", "propput", "propputref")
        return any(f in propflags for f in self.idlflags)


class _ComMemberSpec(_MemberSpec):
    """Specifier for a slot of COM method or property."""
    __slots__ = ("argtypes", "paramflags", "doc")

    def __init__(self, restype, name, argtypes, paramflags, idlflags, doc):
        self.argtypes = argtypes  # type: Tuple[Type[_CData], ...]
        self.paramflags = paramflags  # type: Optional[Tuple[ParamFlagType, ...]]
        self.doc = doc  # type: Optional[str]
        super(_ComMemberSpec, self).__init__(name, idlflags, restype)

    def __iter__(self):
        # for backward compatibility:
        # A function that returns this object used to return a `tuple`.
        # So it is implemented as unpackable as well.
        for item in (self.restype, self.name, self.argtypes, self.paramflags, self.idlflags, self.doc):
            yield item


class _DispMemberSpec(_MemberSpec):
    """Specifier for a slot of dispinterface method or property."""
    __slots__ = ("what", "argspec")

    def __init__(self, what, name, idlflags, restype, argspec):
        self.what = what  # type: str
        self.argspec = argspec  # type: Tuple[ArgSpecElmType, ...]
        super(_DispMemberSpec, self).__init__(name, idlflags, restype)

    @property
    def memid(self):
        # type: () -> int
        try:
            return [x for x in self.idlflags if isinstance(x, int)][0]
        except IndexError:
            raise TypeError("no dispid found in idlflags")

    def __iter__(self):
        # for backward compatibility:
        # A function that returns this object used to return a `tuple`.
        # So it is implemented as unpackable as well.
        for item in (self.what, self.name, self.idlflags, self.restype, self.argspec):
            yield item


def _fix_inout_args(func, argtypes, paramflags):
    # type: (Callable[..., Any], Tuple[Type[_CData], ...], Tuple[ParamFlagType, ...]) -> Callable[..., Any]
    """This function provides a workaround for a bug in `ctypes`.

    [in, out] parameters must be converted with the argtype's `from_param`
    method BEFORE they are passed to the `_ctypes.build_callargs` function
    in `Modules/_ctypes/_ctypes.c`.
    """
    # For details see below.
    #
    # TODO: The workaround should be disabled when a ctypes
    # version is used where the bug is fixed.
    SIMPLETYPE = type(ctypes.c_int)
    BYREFTYPE = type(ctypes.byref(ctypes.c_int()))
    def call_with_inout(self, *args, **kw):
        args = list(args)
        # Indexed by order in the output
        outargs = {}
        outnum = 0
        for i, info in enumerate(paramflags):
            direction = info[0]
            if direction & 3 == 3:
                # This is an [in, out] parameter.
                #
                # Determine name and required type of the parameter.
                name = info[1]
                # [in, out] parameters are passed as pointers,
                # this is the pointed-to type:
                atyp = argtypes[i]._type_

                # Get the actual parameter, either as positional or
                # keyword arg.
                try:
                    try:
                        v = args[i]
                    except IndexError:
                        v = kw[name]
                except KeyError:
                    # no parameter was passed, make an empty one
                    # of the required type
                    v = atyp()
                else:
                    # parameter was passed, call .from_param() to
                    # convert it to a ctypes type.
                    if getattr(v, "_type_", None) is atyp:
                        # Array of or pointer to type 'atyp' was
                        # passed, pointer to 'atyp' expected.
                        pass
                    elif type(atyp) is SIMPLETYPE:
                        # The from_param method of simple types
                        # (c_int, c_double, ...) returns a byref()
                        # object which we cannot use since later
                        # it will be wrapped in a pointer.  Simply
                        # call the constructor with the argument
                        # in that case.
                        v = atyp(v)
                    else:
                        v = atyp.from_param(v)
                        assert not isinstance(v, BYREFTYPE)
                outargs[outnum] = v
                outnum += 1
                if len(args) > i:
                    args[i] = v
                else:
                    kw[name] = v
            elif direction & 2 == 2:
                outnum += 1
        rescode = func(self, *args, **kw)
        # If there is only a single output value, then do not expect it to
        # be iterable.
        if outnum == 1:  # rescode is not iterable
            if len(outargs) == 1:
                rescode = rescode.__ctypes_from_outparam__()
            return rescode
        rescode = list(rescode)
        for outnum, o in outargs.items():
            rescode[outnum] = o.__ctypes_from_outparam__()
        return rescode
    return call_with_inout


class PropertyMapping(object):
    def __init__(self):
        self._data = {}  # type: Dict[Tuple[str, Optional[str], int], List[Optional[Callable[..., Any]]]]

    def add_propget(self, name, doc, nargs, func):
        # type: (str, Optional[str], int, Callable[..., Any]) -> None
        self._data.setdefault((name, doc, nargs), [None, None, None])[0] = func

    def add_propput(self, name, doc, nargs, func):
        # type: (str, Optional[str], int, Callable[..., Any]) -> None
        self._data.setdefault((name, doc, nargs), [None, None, None])[1] = func

    def add_propputref(self, name, doc, nargs, func):
        # type: (str, Optional[str], int, Callable[..., Any]) -> None
        self._data.setdefault((name, doc, nargs), [None, None, None])[2] = func

    def __iter__(self):
        # type: () -> Iterator[Tuple[str, Optional[str], int, Optional[Callable[..., Any]], Optional[Callable[..., Any]]]]
        for (name, doc, nargs), (fget, propput, propputref) in self._data.items():
            if propput is not None and propputref is not None:
                # Create a setter method that examines the argument type
                # and calls 'propputref' if it is an Object (in the VB
                # sense), or call 'propput' otherwise.
                put, putref = propput, propputref
                def put_or_putref(self, *args):
                    if comtypes._is_object(args[-1]):
                        return putref(self, *args)
                    return put(self, *args)
                fset = put_or_putref
            elif propputref is not None:
                fset = propputref
            else:
                fset = propput
            yield (name, doc, nargs, fget, fset)


class PropertyGenerator(object):
    def __init__(self, cls_name):
        # type: (str) -> None
        self._mapping = PropertyMapping()
        self._cls_name = cls_name

    def add(self, m, func):
        # type: (_MemberSpec, Callable[..., Any]) -> None
        """Adds member spec and func to mapping."""
        if "propget" in m.idlflags:
            name, doc, nargs = self.to_propget_keys(m)
            self._mapping.add_propget(name, doc, nargs, func)
        elif "propput" in m.idlflags:
            name, doc, nargs = self.to_propput_keys(m)
            self._mapping.add_propput(name, doc, nargs, func)
        elif "propputref" in m.idlflags:
            name, doc, nargs = self.to_propputref_keys(m)
            self._mapping.add_propputref(name, doc, nargs, func)
        else:
            raise TypeError("no propflag found in idlflags")

    # The following code assumes that the docstrings for
    # propget and propput are identical.
    def __iter__(self):
        # type: () -> Iterator[Tuple[str, _UnionT[property, named_property]]]
        for name, doc, nargs, fget, fset in self._mapping:
            if nargs == 0:
                prop = property(fget, fset, None, doc)
            else:
                # Hm, must be a descriptor where the __get__ method
                # returns a bound object having __getitem__ and
                # __setitem__ methods.
                prop = named_property("%s.%s" % (self._cls_name, name), fget, fset, doc)
            yield (name, prop)

    def to_propget_keys(self, m):
        # type: (_MemberSpec) -> Tuple[str, Optional[str], int]
        raise NotImplementedError

    def to_propput_keys(self, m):
        # type: (_MemberSpec) -> Tuple[str, Optional[str], int]
        raise NotImplementedError

    def to_propputref_keys(self, m):
        # type: (_MemberSpec) -> Tuple[str, Optional[str], int]
        raise NotImplementedError


class ComPropertyGenerator(PropertyGenerator):
    # XXX Hm.  What, when paramflags is None?
    # Or does have '0' values?
    # Seems we loose then, at least for properties...
    def to_propget_keys(self, m):
        # type: (_ComMemberSpec) -> Tuple[str, Optional[str], int]
        assert m.name.startswith("_get_")
        assert m.paramflags is not None
        nargs = len([f for f in m.paramflags if f[0] & 7 in (0, 1)])
        # XXX or should we do this?
        # nargs = len([f for f in paramflags if (f[0] & 1) or (f[0] == 0)])
        return m.name[len("_get_"):], m.doc, nargs

    def to_propput_keys(self, m):
        # type: (_ComMemberSpec) -> Tuple[str, Optional[str], int]
        assert m.name.startswith("_set_")
        assert m.paramflags is not None
        nargs = len([f for f in m.paramflags if f[0] & 7 in (0, 1)]) - 1
        return m.name[len("_set_"):], m.doc, nargs

    def to_propputref_keys(self, m):
        # type: (_ComMemberSpec) -> Tuple[str, Optional[str], int]
        assert m.name.startswith("_setref_")
        assert m.paramflags is not None
        nargs = len([f for f in m.paramflags if f[0] & 7 in (0, 1)]) - 1
        return m.name[len("_setref_"):], m.doc, nargs


class DispPropertyGenerator(PropertyGenerator):
    def to_propget_keys(self, m):
        # type: (_DispMemberSpec) -> Tuple[str, Optional[str], int]
        return m.name, None, len(m.argspec)

    def to_propput_keys(self, m):
        # type: (_DispMemberSpec) -> Tuple[str, Optional[str], int]
        return m.name, None, len(m.argspec) - 1

    def to_propputref_keys(self, m):
        # type: (_DispMemberSpec) -> Tuple[str, Optional[str], int]
        return m.name, None, len(m.argspec) - 1


class ComMemberGenerator(object):
    def __init__(self, cls_name, vtbl_offset, iid):
        # type: (str, int, comtypes.GUID) -> None
        self._vtbl_offset = vtbl_offset
        self._iid = iid
        self._props = ComPropertyGenerator(cls_name)
        # sequence of (name: str, func: Callable, raw_func: Callable, is_prop: bool)
        self._mths = []  # type: List[Tuple[str, Callable[..., Any], Callable[..., Any], bool]]
        self._member_index = 0

    def add(self, m):
        # type: (_ComMemberSpec) -> None
        proto = ctypes.WINFUNCTYPE(m.restype, *m.argtypes)
        # a low level unbound method calling the com method.
        # attach it with a private name (__com_AddRef, for example),
        # so that custom method implementations can call it.
        vidx = self._member_index + self._vtbl_offset
        # If the method returns a HRESULT, we pass the interface iid,
        # so that we can request error info for the interface.
        iid = self._iid if m.restype == ctypes.HRESULT else None
        raw_func = proto(vidx, m.name, None, iid)  # low level
        func = self._fix_args(m, proto(vidx, m.name, m.paramflags, iid))  # high level
        func.__doc__ = m.doc
        func.__name__ = m.name  # for pyhelp
        is_prop = m.is_prop()
        if is_prop:
            self._props.add(m, func)
        self._mths.append((m.name, func, raw_func, is_prop))
        self._member_index += 1

    def _fix_args(self, m, func):
        # type: (_ComMemberSpec, Callable[..., Any]) -> Callable[..., Any]
        """This is a workaround. See `_fix_inout_args` docstring and comments."""
        if m.paramflags:
            dirflags = [(p[0]&3) for p in m.paramflags]
            if 3 in dirflags:
                return _fix_inout_args(func, m.argtypes, m.paramflags)
        return func

    def methods(self):
        return iter(self._mths)

    def properties(self):
        return iter(self._props)


class DispMemberGenerator(object):
    def __init__(self, cls_name):
        # type: (str) -> None
        self._props = DispPropertyGenerator(cls_name)
        # sequence of (name: str, func_or_prop: Callable | property, is_prop: bool)
        self._items = []  # type: List[Tuple[str, _UnionT[Callable[..., Any], property], bool]]

    def add(self, m):
        # type: (_DispMemberSpec) -> None
        if m.what == "DISPPROPERTY":  # DISPPROPERTY
            assert not m.argspec # XXX does not yet work for properties with parameters
            is_prop = True
            accessor = self._make_disp_property(m)
            self._items.append((m.name, accessor, is_prop))
        else:  # DISPMETHOD
            func = self._make_disp_method(m)
            func.__name__ = m.name
            is_prop = m.is_prop()
            if is_prop:
                self._props.add(m, func)
            else:
                self._items.append((m.name, func, is_prop))

    def _make_disp_property(self, m):
        # type: (_DispMemberSpec) -> property
        # XXX doc string missing in property
        memid = m.memid
        def fget(obj):
            return obj.Invoke(memid, _invkind=2) # DISPATCH_PROPERTYGET
        if "readonly" in m.idlflags:
            return property(fget)
        def fset(obj, value):
            # Detect whether to use DISPATCH_PROPERTYPUT or
            # DISPATCH_PROPERTYPUTREF
            invkind = 8 if comtypes._is_object(value) else 4
            return obj.Invoke(memid, value, _invkind=invkind)
        return property(fget, fset)

    # Should the funcs/mths we create have restype and/or argtypes attributes?
    def _make_disp_method(self, m):
        # type: (_DispMemberSpec) -> Callable[..., Any]
        memid = m.memid
        if "propget" in m.idlflags:
            def getfunc(obj, *args, **kw):
                return obj.Invoke(memid, _invkind=2, *args, **kw) # DISPATCH_PROPERTYGET
            return getfunc
        elif "propput" in m.idlflags:
            def putfunc(obj, *args, **kw):
                return obj.Invoke(memid, _invkind=4, *args, **kw) # DISPATCH_PROPERTYPUT
            return putfunc
        elif "propputref" in m.idlflags:
            def putreffunc(obj, *args, **kw):
                return obj.Invoke(memid, _invkind=8, *args, **kw) # DISPATCH_PROPERTYPUTREF
            return putreffunc
        # a first attempt to make use of the restype.  Still, support for
        # named arguments and default argument values should be added.
        if hasattr(m.restype, "__com_interface__"):
            interface = m.restype.__com_interface__  # type: ignore
            def comitffunc(obj, *args, **kw):
                result = obj.Invoke(memid, _invkind=1, *args, **kw)
                if result is None:
                    return
                return result.QueryInterface(interface)
            return comitffunc
        def func(obj, *args, **kw):
            return obj.Invoke(memid, _invkind=1, *args, **kw) # DISPATCH_METHOD
        return func

    def items(self):
        return iter(self._items)

    def properties(self):
        return iter(self._props)


################################################################
# helper classes for COM propget / propput
# Should they be implemented in C for speed?

class bound_named_property(object):
    def __init__(self, name, fget, fset, instance):
        self.name = name
        self.instance = instance
        self.fget = fget
        self.fset = fset

    def __getitem__(self, index):
        if self.fget is None:
            raise TypeError("unsubscriptable object")
        if isinstance(index, tuple):
            return self.fget(self.instance, *index)
        elif index == comtypes._all_slice:
            return self.fget(self.instance)
        else:
            return self.fget(self.instance, index)

    def __call__(self, *args):
        if self.fget is None:
            raise TypeError("object is not callable")
        return self.fget(self.instance, *args)

    def __setitem__(self, index, value):
        if self.fset is None:
            raise TypeError("object does not support item assignment")
        if isinstance(index, tuple):
            self.fset(self.instance, *(index + (value,)))
        elif index == comtypes._all_slice:
            self.fset(self.instance, value)
        else:
            self.fset(self.instance, index, value)

    def __repr__(self):
        return "<bound_named_property %r at %x>" % (self.name, id(self))

    def __iter__(self):
        """ Explicitly disallow iteration. """
        msg = "%r is not iterable" % self.name
        raise TypeError(msg)


class named_property(object):
    def __init__(self, name, fget=None, fset=None, doc=None):
        self.name = name
        self.fget = fget
        self.fset = fset
        self.__doc__ = doc

    def __get__(self, instance, owner=None):
        if instance is None:
            return self
        return bound_named_property(self.name, self.fget, self.fset, instance)

    # Make this a data descriptor
    def __set__(self, instance):
        raise AttributeError("Unsettable attribute")

    def __repr__(self):
        return "<named_property %r at %x>" % (self.name, id(self))
