# More type descriptions from parsed COM typelibaries, extending those
# in typedesc_base

import ctypes
from comtypes import TYPE_CHECKING
from comtypes.typeinfo import ITypeLib, TLIBATTR
from comtypes.tools.typedesc_base import *

if TYPE_CHECKING:
    from typing import Any, List, Optional, Tuple, Union as _UnionT


class TypeLib(object):
    def __init__(self, name, guid, major, minor, doc=None):
        # type: (str, str, int, int, Optional[str]) -> None
        self.name = name
        self.guid = guid
        self.major = major
        self.minor = minor
        self.doc = doc

    def __repr__(self):
        return "<TypeLib(%s: %s, %s, %s)>" % (self.name, self.guid, self.major, self.minor)

class Constant(object):
    def __init__(self, name, typ, value):
        # type: (str, _UnionT[Typedef, FundamentalType], Any) -> None
        self.name = name
        self.typ = typ
        self.value = value

class External(object):
    def __init__(self, tlib, name, size, align, docs=None):
        # type: (ITypeLib, str, int, int, Optional[Tuple[str, str]]) -> None
        # the type library containing the symbol
        self.tlib = tlib
        # name of symbol
        self.symbol_name = name
        self.size = size
        self.align = align
        # type lib description
        self.docs = docs

    def get_head(self):
        # type: () -> External
        # codegen might call this
        return self

class SAFEARRAYType(object):
    def __init__(self, typ):
        # type: (Any) -> None
        self.typ = typ
        self.align = self.size = ctypes.sizeof(ctypes.c_void_p) * 8

class ComMethod(object):
    # custom COM method, parsed from typelib
    def __init__(self, invkind, memid, name, returns, idlflags, doc):
        # type: (int, int, str, Any, List[str], Optional[str]) -> None
        self.invkind = invkind
        self.name = name
        self.returns = returns
        self.idlflags = idlflags
        self.memid = memid
        self.doc = doc
        self.arguments = []  # type: List[Tuple[Any, str, List[str], Optional[Any]]]

    def add_argument(self, typ, name, idlflags, default):
        # type: (Any, str, List[str], Optional[Any]) -> None
        self.arguments.append((typ, name, idlflags, default))

class DispMethod(object):
    # dispatchable COM method, parsed from typelib
    def __init__(self, dispid, invkind, name, returns, idlflags, doc):
        # type: (int, int, str, Any, List[str], Optional[str]) -> None
        self.dispid = dispid
        self.invkind = invkind
        self.name = name
        self.returns = returns
        self.idlflags = idlflags
        self.doc = doc
        self.arguments = []  # type: List[Tuple[Any, str, List[str], Optional[Any]]]

    def add_argument(self, typ, name, idlflags, default):
        # type: (Any, str, List[str], Optional[Any]) -> None
        self.arguments.append((typ, name, idlflags, default))

class DispProperty(object):
    # dispatchable COM property, parsed from typelib
    def __init__(self, dispid, name, typ, idlflags, doc):
        # type: (int, str, Any, List[str], Optional[Any]) -> None
        self.dispid = dispid
        self.name = name
        self.typ = typ
        self.idlflags = idlflags
        self.doc = doc

class DispInterfaceHead(object):
    def __init__(self, itf):
        # type: (DispInterface) -> None
        self.itf = itf

class DispInterfaceBody(object):
    def __init__(self, itf):
        # type: (DispInterface) -> None
        self.itf = itf

class DispInterface(object):
    def __init__(self, name, members, base, iid, idlflags):
        # type: (str, List[_UnionT[DispMethod, DispProperty]], Any, str, List[str]) -> None
        self.name = name
        self.members = members
        self.base = base
        self.iid = iid
        self.idlflags = idlflags
        self.itf_head = DispInterfaceHead(self)
        self.itf_body = DispInterfaceBody(self)

    def get_body(self):
        # type: () -> DispInterfaceBody
        return self.itf_body

    def get_head(self):
        # type: () -> DispInterfaceHead
        return self.itf_head

class ComInterfaceHead(object):
    def __init__(self, itf):
        # type: (ComInterface) -> None
        self.itf = itf

class ComInterfaceBody(object):
    def __init__(self, itf):
        # type: (ComInterface) -> None
        self.itf = itf

class ComInterface(object):
    def __init__(self, name, members, base, iid, idlflags):
        # type: (str, List[ComMethod], Any, str, List[str]) -> None
        self.name = name
        self.members = members
        self.base = base
        self.iid = iid
        self.idlflags = idlflags
        self.itf_head = ComInterfaceHead(self)
        self.itf_body = ComInterfaceBody(self)

    def get_body(self):
        # type: () -> ComInterfaceBody
        return self.itf_body

    def get_head(self):
        # type: () -> ComInterfaceHead
        return self.itf_head

class CoClass(object):
    def __init__(self, name, clsid, idlflags, tlibattr):
        # type: (str, str, List[str], TLIBATTR) -> None
        self.name = name
        self.clsid = clsid
        self.idlflags = idlflags
        self.tlibattr = tlibattr
        self.interfaces = []  # type: List[Tuple[Any, int]]

    def add_interface(self, itf, idlflags):
        # type: (Any, int) -> None
        self.interfaces.append((itf, idlflags))
