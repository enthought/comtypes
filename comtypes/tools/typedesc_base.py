# typedesc.py - classes representing C type descriptions

import comtypes
from comtypes import TYPE_CHECKING

if TYPE_CHECKING:
    from typing import Any, List, Optional, Tuple, Union as _UnionT, SupportsInt


class Argument(object):
    "a Parameter in the argument list of a callable (Function, Method, ...)"
    def __init__(self, atype, name):
        self.atype = atype
        self.name = name

class _HasArgs(object):

    def __init__(self):
        self.arguments = []

    def add_argument(self, arg):
        assert isinstance(arg, Argument)
        self.arguments.append(arg)

    def iterArgTypes(self):
        for a in self.arguments:
            yield a.atype

    def iterArgNames(self):
        for a in self.arguments:
            yield a.name

    def fixup_argtypes(self, typemap):
        for a in self.arguments:
            a.atype = typemap[a.atype]


################

class Alias(object):
    # a C preprocessor alias, like #define A B
    def __init__(self, name, alias, typ=None):
        self.name = name
        self.alias = alias
        self.typ = typ

class Macro(object):
    # a C preprocessor definition with arguments
    def __init__(self, name, args, body):
        # all arguments are strings, args is the literal argument list
        # *with* the parens around it:
        # Example: Macro("CD_INDRIVE", "(status)", "((int)status > 0)")
        self.name = name
        self.args = args
        self.body = body

class File(object):
    def __init__(self, name):
        self.name = name

class Function(_HasArgs):
    location = None
    def __init__(self, name, returns, attributes, extern):
        _HasArgs.__init__(self)
        self.name = name
        self.returns = returns
        self.attributes = attributes # dllimport, __stdcall__, __cdecl__
        self.extern = extern

class Constructor(_HasArgs):
    location = None
    def __init__(self, name):
        _HasArgs.__init__(self)
        self.name = name

class OperatorFunction(_HasArgs):
    location = None
    def __init__(self, name, returns):
        _HasArgs.__init__(self)
        self.name = name
        self.returns = returns

class FunctionType(_HasArgs):
    location = None
    def __init__(self, returns, attributes):
        _HasArgs.__init__(self)
        self.returns = returns
        self.attributes = attributes

class Method(_HasArgs):
    location = None
    def __init__(self, name, returns):
        _HasArgs.__init__(self)
        self.name = name
        self.returns = returns

class FundamentalType(object):
    location = None
    def __init__(self, name, size, align):
        self.name = name
        if name != "void":
            self.size = int(size)
            self.align = int(align)

class PointerType(object):
    location = None
    def __init__(self, typ, size, align):
        self.typ = typ
        self.size = int(size)
        self.align = int(align)

class Typedef(object):
    location = None
    def __init__(self, name, typ):
        self.name = name
        self.typ = typ

class ArrayType(object):
    location = None
    def __init__(self, typ, min, max):
        # type: (Any, int, int) -> None
        self.typ = typ
        self.min = min
        self.max = max

class StructureHead(object):
    location = None
    def __init__(self, struct):
        # type: (_Struct_Union_Base) -> None
        self.struct = struct

class StructureBody(object):
    location = None
    def __init__(self, struct):
        # type: (_Struct_Union_Base) -> None
        self.struct = struct

class _Struct_Union_Base(object):
    if TYPE_CHECKING:
        name = comtypes.hints.AnnoField()  # type: str
        align = comtypes.hints.AnnoField()  # type: int
        members = comtypes.hints.AnnoField()  # type: List[_UnionT[Field, Method, Constructor]]
        bases = comtypes.hints.AnnoField()  # type: List[_Struct_Union_Base]
        artificial = comtypes.hints.AnnoField()  # type: Optional[Any]
        size = comtypes.hints.AnnoField()  # type: Optional[int]
        _recordinfo_ = comtypes.hints.AnnoField()  # type: Tuple[str, int, int, int, str]

    location = None
    def __init__(self):
        self.struct_body = StructureBody(self)
        self.struct_head = StructureHead(self)

    def get_body(self):
        # type: () -> StructureBody
        return self.struct_body

    def get_head(self):
        # type: () -> StructureHead
        return self.struct_head

class Structure(_Struct_Union_Base):
    def __init__(self, name, align, members, bases, size, artificial=None):
        # type: (str, SupportsInt, List[Field], List[Any], Optional[SupportsInt], Optional[Any]) -> None
        self.name = name
        self.align = int(align)
        self.members = members
        self.bases = bases
        self.artificial = artificial
        if size is not None:
            self.size = int(size)
        else:
            self.size = None
        super(Structure, self).__init__()

class Union(_Struct_Union_Base):
    def __init__(self, name, align, members, bases, size, artificial=None):
        # type: (str, SupportsInt, List[Field], List[Any], Optional[SupportsInt], Optional[Any]) -> None
        self.name = name
        self.align = int(align)
        self.members = members
        self.bases = bases
        self.artificial = artificial
        if size is not None:
            self.size = int(size)
        else:
            self.size = None
        super(Union, self).__init__()

class Field(object):
    def __init__(self, name, typ, bits, offset):
        # type: (str, Any, Optional[Any], SupportsInt) -> None
        self.name = name
        self.typ = typ
        self.bits = bits
        self.offset = int(offset)

class CvQualifiedType(object):
    def __init__(self, typ, const, volatile):
        self.typ = typ
        self.const = const
        self.volatile = volatile

class Enumeration(object):
    location = None
    def __init__(self, name, size, align):
        # type: (str, SupportsInt, SupportsInt) -> None
        self.name = name
        self.size = int(size)
        self.align = int(align)
        self.values = []  # type: List[EnumValue]

    def add_value(self, v):
        # type: (EnumValue) -> None
        self.values.append(v)

class EnumValue(object):
    def __init__(self, name, value, enumeration):
        # type: (str, int, Enumeration) -> None
        self.name = name
        self.value = value
        self.enumeration = enumeration

class Variable(object):
    location = None
    def __init__(self, name, typ, init=None):
        self.name = name
        self.typ = typ
        self.init = init

################################################################
