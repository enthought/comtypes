# More type descriptions from parsed COM typelibaries, extending those
# in typedesc_base

import ctypes
from typing import Any, List, Optional, Tuple, Union as _UnionT

from comtypes.typeinfo import ITypeLib, TLIBATTR
from comtypes.tools.typedesc_base import *


class TypeLib(object):
    def __init__(
        self, name: str, guid: str, major: int, minor: int, doc: Optional[str] = None
    ) -> None:
        self.name = name
        self.guid = guid
        self.major = major
        self.minor = minor
        self.doc = doc

    def __repr__(self):
        return "<TypeLib(%s: %s, %s, %s)>" % (
            self.name,
            self.guid,
            self.major,
            self.minor,
        )


class Constant(object):
    def __init__(
        self, name: str, typ: _UnionT[Typedef, FundamentalType], value: Any
    ) -> None:
        self.name = name
        self.typ = typ
        self.value = value


class External(object):
    def __init__(
        self,
        tlib: ITypeLib,
        name: str,
        size: int,
        align: int,
        docs: Optional[Tuple[str, str]] = None,
    ) -> None:
        # the type library containing the symbol
        self.tlib = tlib
        # name of symbol
        self.symbol_name = name
        self.size = size
        self.align = align
        # type lib description
        self.docs = docs

    def get_head(self) -> "External":
        # codegen might call this
        return self


class SAFEARRAYType(object):
    def __init__(self, typ: Any) -> None:
        self.typ = typ
        self.align = self.size = ctypes.sizeof(ctypes.c_void_p) * 8


class ComMethod(object):
    # custom COM method, parsed from typelib
    def __init__(
        self,
        invkind: int,
        memid: int,
        name: str,
        returns: Any,
        idlflags: List[str],
        doc: Optional[str],
    ) -> None:
        self.invkind = invkind
        self.name = name
        self.returns = returns
        self.idlflags = idlflags
        self.memid = memid
        self.doc = doc
        self.arguments: List[Tuple[Any, str, List[str], Optional[Any]]] = []

    def add_argument(
        self, typ: Any, name: str, idlflags: List[str], default: Optional[Any]
    ) -> None:
        self.arguments.append((typ, name, idlflags, default))


class DispMethod(object):
    # dispatchable COM method, parsed from typelib
    def __init__(
        self,
        dispid: int,
        invkind: int,
        name: str,
        returns: Any,
        idlflags: List[str],
        doc: Optional[str],
    ) -> None:
        self.dispid = dispid
        self.invkind = invkind
        self.name = name
        self.returns = returns
        self.idlflags = idlflags
        self.doc = doc
        self.arguments: List[Tuple[Any, str, List[str], Optional[Any]]] = []

    def add_argument(
        self, typ: Any, name: str, idlflags: List[str], default: Optional[Any]
    ) -> None:
        self.arguments.append((typ, name, idlflags, default))


class DispProperty(object):
    # dispatchable COM property, parsed from typelib
    def __init__(
        self, dispid: int, name: str, typ: Any, idlflags: List[str], doc: Optional[Any]
    ) -> None:
        self.dispid = dispid
        self.name = name
        self.typ = typ
        self.idlflags = idlflags
        self.doc = doc


class DispInterfaceHead(object):
    def __init__(self, itf: "DispInterface") -> None:
        self.itf = itf


class DispInterfaceBody(object):
    def __init__(self, itf: "DispInterface") -> None:
        self.itf = itf


class DispInterface(object):
    def __init__(
        self,
        name: str,
        members: List[_UnionT[DispMethod, DispProperty]],
        base: Any,
        iid: str,
        idlflags: List[str],
    ) -> None:
        self.name = name
        self.members = members
        self.base = base
        self.iid = iid
        self.idlflags = idlflags
        self.itf_head = DispInterfaceHead(self)
        self.itf_body = DispInterfaceBody(self)

    def get_body(self) -> DispInterfaceBody:
        return self.itf_body

    def get_head(self) -> DispInterfaceHead:
        return self.itf_head


class ComInterfaceHead(object):
    def __init__(self, itf: "ComInterface") -> None:
        self.itf = itf


class ComInterfaceBody(object):
    def __init__(self, itf: "ComInterface") -> None:
        self.itf = itf


class ComInterface(object):
    def __init__(
        self,
        name: str,
        members: List[ComMethod],
        base: Any,
        iid: str,
        idlflags: List[str],
    ) -> None:
        self.name = name
        self.members = members
        self.base = base
        self.iid = iid
        self.idlflags = idlflags
        self.itf_head = ComInterfaceHead(self)
        self.itf_body = ComInterfaceBody(self)

    def get_body(self) -> ComInterfaceBody:
        return self.itf_body

    def get_head(self) -> ComInterfaceHead:
        return self.itf_head


class CoClass(object):
    def __init__(
        self, name: str, clsid: str, idlflags: List[str], tlibattr: TLIBATTR
    ) -> None:
        self.name = name
        self.clsid = clsid
        self.idlflags = idlflags
        self.tlibattr = tlibattr
        self.interfaces: List[Tuple[Any, int]] = []

    def add_interface(self, itf: Any, idlflags: int) -> None:
        self.interfaces.append((itf, idlflags))
