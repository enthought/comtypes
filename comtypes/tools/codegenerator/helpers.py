import keyword
import logging
import textwrap
from typing import Any
from typing import Dict, List, Set, Tuple
from typing import Iterator
from typing import Union as _UnionT

import comtypes
from comtypes.tools import typedesc


version = comtypes.__version__

logger = logging.getLogger(__name__)

__warn_on_munge__ = __debug__


class lcid(object):
    def __repr__(self):
        return "_lcid"


lcid = lcid()


class dispid(object):
    def __init__(self, memid):
        self.memid = memid

    def __repr__(self):
        return "dispid(%s)" % self.memid


class helpstring(object):
    def __init__(self, text):
        self.text = text

    def __repr__(self):
        return "helpstring(%r)" % self.text


# XXX Should this be in ctypes itself?
ctypes_names = {
    "unsigned char": "c_ubyte",
    "signed char": "c_byte",
    "char": "c_char",
    "wchar_t": "c_wchar",
    "short unsigned int": "c_ushort",
    "short int": "c_short",
    "long unsigned int": "c_ulong",
    "long int": "c_long",
    "long signed int": "c_long",
    "unsigned int": "c_uint",
    "int": "c_int",
    "long long unsigned int": "c_ulonglong",
    "long long int": "c_longlong",
    "double": "c_double",
    "float": "c_float",
    # Hm...
    "void": "None",
}


def get_real_type(tp):
    if type(tp) is typedesc.Typedef:
        return get_real_type(tp.typ)
    elif isinstance(tp, typedesc.CvQualifiedType):
        return get_real_type(tp.typ)
    return tp


ASSUME_STRINGS = True


def _calc_packing(struct, fields, pack, isStruct):
    # Try a certain packing, raise PackingError if field offsets,
    # total size ot total alignment is wrong.
    if struct.size is None:  # incomplete struct
        return -1
    if struct.name in dont_assert_size:
        return None
    if struct.bases:
        size = struct.bases[0].size
        total_align = struct.bases[0].align
    else:
        size = 0
        total_align = 8  # in bits
    for i, f in enumerate(fields):
        if f.bits:  # this code cannot handle bit field sizes.
            # print "##XXX FIXME"
            return -2  # XXX FIXME
        s, a = storage(f.typ)
        if pack is not None:
            a = min(pack, a)
        if size % a:
            size += a - size % a
        if isStruct:
            if size != f.offset:
                raise PackingError("field %s offset (%s/%s)" % (f.name, size, f.offset))
            size += s
        else:
            size = max(size, s)
        total_align = max(total_align, a)
    if total_align != struct.align:
        raise PackingError("total alignment (%s/%s)" % (total_align, struct.align))
    a = total_align
    if pack is not None:
        a = min(pack, a)
    if size % a:
        size += a - size % a
    if size != struct.size:
        raise PackingError("total size (%s/%s)" % (size, struct.size))


def calc_packing(struct, fields):
    # try several packings, starting with unspecified packing
    isStruct = isinstance(struct, typedesc.Structure)
    for pack in [None, 16 * 8, 8 * 8, 4 * 8, 2 * 8, 1 * 8]:
        try:
            _calc_packing(struct, fields, pack, isStruct)
        except PackingError as details:
            continue
        else:
            if pack is None:
                return None

            return int(pack / 8)

    raise PackingError("PACKING FAILED: %s" % details)


class PackingError(Exception):
    pass


# XXX These should be filtered out in gccxmlparser.
dont_assert_size = set(
    [
        "__si_class_type_info_pseudo",
        "__class_type_info_pseudo",
    ]
)


def storage(t):
    # return the size and alignment of a type
    if isinstance(t, typedesc.Typedef):
        return storage(t.typ)
    elif isinstance(t, typedesc.ArrayType):
        s, a = storage(t.typ)
        return s * (int(t.max) - int(t.min) + 1), a
    return int(t.size), int(t.align)


################################################################


def name_wrapper_module(tlib):
    """Determine the name of a typelib wrapper module"""
    libattr = tlib.GetLibAttr()
    modname = "_%s_%s_%s_%s" % (
        str(libattr.guid)[1:-1].replace("-", "_"),
        libattr.lcid,
        libattr.wMajorVerNum,
        libattr.wMinorVerNum,
    )
    return "comtypes.gen.%s" % modname


def name_friendly_module(tlib):
    """Determine the friendly-name of a typelib module.
    If cannot get friendly-name from typelib, returns `None`.
    """
    try:
        modulename = tlib.GetDocumentation(-1)[0]
    except comtypes.COMError:
        return
    return "comtypes.gen.%s" % modulename


################################################################

_DefValType = _UnionT["lcid", Any, None]
_IdlFlagType = _UnionT[str, dispid, helpstring]


def _to_arg_definition(
    type_name: str,
    arg_name: str,
    idlflags: List[str],
    default: _DefValType,
) -> str:
    if default is not None:
        code = f"        ({idlflags!r}, {type_name}, '{arg_name}', {default!r})"
        if len(code) > 80:
            code = (
                "        (\n"
                f"            {idlflags!r},\n"
                f"            {type_name},\n"
                f"            '{arg_name}',\n"
                f"            {default!r}\n"
                "        )"
            )
    else:
        code = f"        ({idlflags!r}, {type_name}, '{arg_name}')"
        if len(code) > 80:
            code = (
                "        (\n"
                f"            {idlflags!r},\n"
                f"            {type_name},\n"
                f"            '{arg_name}',\n"
                "        )"
            )
    return code


class ComMethodGenerator(object):
    def __init__(self, m: typedesc.ComMethod, isdual: bool) -> None:
        self._m = m
        self._isdual = isdual
        self.data: List[str] = []
        self._to_type_name = TypeNamer()

    def generate(self) -> str:
        if not self._m.arguments:
            self._make_noargs()
        else:
            self._make_withargs()
        return "\n".join(self.data)

    def _get_common_elms(self) -> Tuple[List[_IdlFlagType], str, str]:
        idlflags: List[_IdlFlagType] = []
        if self._isdual:
            idlflags.append(dispid(self._m.memid))
            idlflags.extend(self._m.idlflags)
        else:  # We don't include the dispid for non-dispatch COM interfaces
            idlflags.extend(self._m.idlflags)
        if __debug__ and self._m.doc:
            idlflags.insert(1, helpstring(self._m.doc))
        type_name = self._to_type_name(self._m.returns)
        return (idlflags, type_name, self._m.name)

    def _make_noargs(self) -> None:
        flags, type_name, member_name = self._get_common_elms()
        code = f"    COMMETHOD({flags!r}, {type_name}, '{member_name}'),"
        if len(code) > 80:
            code = (
                "    COMMETHOD(\n"
                f"        {flags!r},\n"
                f"        {type_name},\n"
                f"        '{member_name}',\n"
                "    ),"
            )
        self.data.append(code)

    def _make_withargs(self) -> None:
        flags, type_name, member_name = self._get_common_elms()
        code = (
            "    COMMETHOD(\n"
            f"        {flags!r},\n"
            f"        {type_name},\n"
            f"        '{member_name}',"
        )
        self.data.append(code)
        arglist = [_to_arg_definition(*i) for i in self._iter_args()]
        self.data.append(",\n".join(arglist))
        self.data.append("    ),")

    def _iter_args(self) -> Iterator[Tuple[str, str, List[str], _DefValType]]:
        for typ, arg_name, _f, _defval in self._m.arguments:
            ###########################################################
            # IDL files that contain 'open arrays' or 'conformant
            # varying arrays' method parameters are strange.
            # These arrays have both a 'size_is()' and
            # 'length_is()' attribute, like this example from
            # dia2.idl (in the DIA SDK):
            #
            # interface IDiaSymbol: IUnknown {
            # ...
            #     HRESULT get_dataBytes(
            #         [in] DWORD cbData,
            #         [out] DWORD *pcbData,
            #         [out, size_is(cbData),
            #          length_is(*pcbData)] BYTE data[]
            #     );
            #
            # The really strange thing is that the decompiled type
            # library then contains this declaration, which declares
            # the interface itself as [out] method parameter:
            #
            # interface IDiaSymbol: IUnknown {
            # ...
            #     HRESULT _stdcall get_dataBytes(
            #         [in] unsigned long cbData,
            #         [out] unsigned long* pcbData,
            #         [out] IDiaSymbol data);
            #
            # Of course, comtypes does not accept a COM interface
            # as method parameter; so replace the parameter type
            # with the comtypes spelling of 'unsigned char *', and
            # mark the parameter as [in, out], so the IDL
            # equivalent would be like this:
            #
            # interface IDiaSymbol: IUnknown {
            # ...
            #     HRESULT _stdcall get_dataBytes(
            #         [in] unsigned long cbData,
            #         [out] unsigned long* pcbData,
            #         [in, out] BYTE data[]);
            ###########################################################
            idlflags = list(_f)  # shallow copy to avoid side effects
            if isinstance(typ, typedesc.ComInterface):
                type_name = "OPENARRAY"
                if "in" not in idlflags:
                    idlflags.append("in")
            else:
                type_name = self._to_type_name(typ)
            if "lcid" in idlflags:  # and 'in' in idlflags:
                default = lcid
            else:
                default = _defval
            yield (type_name, arg_name, idlflags, default)


class DispMethodGenerator(object):
    def __init__(self, m: typedesc.DispMethod) -> None:
        self._m = m
        self.data: List[str] = []
        self._to_type_name = TypeNamer()

    def generate(self) -> str:
        if not self._m.arguments:
            self._make_noargs()
        else:
            self._make_withargs()
        return "\n".join(self.data)

    def _get_common_elms(self) -> Tuple[List[_IdlFlagType], str, str]:
        idlflags: List[_IdlFlagType] = []
        idlflags.append(dispid(self._m.dispid))
        idlflags.extend(self._m.idlflags)
        if __debug__ and self._m.doc:
            idlflags.insert(1, helpstring(self._m.doc))
        type_name = self._to_type_name(self._m.returns)
        return (idlflags, type_name, self._m.name)

    def _make_noargs(self) -> None:
        flags, type_name, member_name = self._get_common_elms()
        code = f"    DISPMETHOD({flags!r}, {type_name}, '{member_name}'),"
        if len(code) > 80:
            code = (
                "    DISPMETHOD(\n"
                f"        {flags!r},\n"
                f"        {type_name},\n"
                f"        '{member_name}',\n"
                "    ),"
            )
        self.data.append(code)

    def _make_withargs(self) -> None:
        flags, type_name, member_name = self._get_common_elms()
        code = (
            "    DISPMETHOD(\n"
            f"        {flags!r},\n"
            f"        {type_name},\n"
            f"        '{member_name}',"
        )
        self.data.append(code)
        arglist = [_to_arg_definition(*i) for i in self._iter_args()]
        self.data.append(",\n".join(arglist))
        self.data.append("    ),")

    def _iter_args(self) -> Iterator[Tuple[str, str, List[str], _DefValType]]:
        for typ, arg_name, idlflags, default in self._m.arguments:
            type_name = self._to_type_name(typ)
            yield (type_name, arg_name, idlflags, default)


class DispPropertyGenerator(object):
    def __init__(self, m: typedesc.DispProperty) -> None:
        self._m = m
        self._to_type_name = TypeNamer()

    def generate(self) -> str:
        flags, type_name, member_name = self._get_common_elms()
        code = f"    DISPPROPERTY({flags!r}, {type_name}, '{member_name}'),"
        if len(code) > 80:
            code = (
                "    DISPPROPERTY(\n"
                f"        {flags!r},\n"
                f"        {type_name},\n"
                f"        '{member_name}'\n"
                "    ),"
            )
        return code

    def _get_common_elms(self) -> Tuple[List[_IdlFlagType], str, str]:
        idlflags: List[_IdlFlagType] = []
        idlflags.append(dispid(self._m.dispid))
        idlflags.extend(self._m.idlflags)
        if __debug__ and self._m.doc:
            idlflags.insert(1, helpstring(self._m.doc))
        type_name = self._to_type_name(self._m.typ)
        return (idlflags, type_name, self._m.name)


class TypeNamer(object):
    def __call__(self, t: Any) -> str:
        # Return a string, containing an expression which can be used
        # to refer to the type. Assumes the 'from ctypes import *'
        # namespace is available.
        if isinstance(t, typedesc.SAFEARRAYType):
            return "_midlSAFEARRAY(%s)" % self(t.typ)
        # if isinstance(t, typedesc.CoClass):
        #     return "%s._com_interfaces_[0]" % t.name
        if isinstance(t, typedesc.Typedef):
            return t.name
        if isinstance(t, typedesc.PointerType):
            _t, pcnt = self._inspect_PointerType(t)
            return "%s%s%s" % ("POINTER(" * pcnt, self(_t), ")" * pcnt)
        elif isinstance(t, typedesc.ArrayType):
            return "%s * %s" % (self(t.typ), int(t.max) + 1)
        elif isinstance(t, typedesc.FunctionType):
            args = [self(x) for x in [t.returns] + list(t.iterArgTypes())]
            if "__stdcall__" in t.attributes:
                return "WINFUNCTYPE(%s)" % ", ".join(args)
            else:
                return "CFUNCTYPE(%s)" % ", ".join(args)
        elif isinstance(t, typedesc.CvQualifiedType):
            # const and volatile are ignored
            return "%s" % self(t.typ)
        elif isinstance(t, typedesc.FundamentalType):
            return ctypes_names[t.name]
        elif isinstance(t, typedesc.Structure):
            return t.name
        elif isinstance(t, typedesc.Enumeration):
            if t.name:
                return t.name
            return "c_int"  # enums are integers
        elif isinstance(t, typedesc.EnumValue):
            if keyword.iskeyword(t.name):
                return t.name + "_"
            return t.name
        elif isinstance(t, typedesc.External):
            # t.symbol_name - symbol to generate
            # t.tlib - the ITypeLib pointer to the typelibrary containing the symbols definition
            modname = name_wrapper_module(t.tlib)
            return "%s.%s" % (modname, t.symbol_name)
        return t.name

    def _inspect_PointerType(
        self, t: typedesc.PointerType, count: int = 0
    ) -> Tuple[Any, int]:
        if ASSUME_STRINGS:
            x = get_real_type(t.typ)
            if isinstance(x, typedesc.FundamentalType):
                if x.name == "char":
                    return typedesc.Typedef("STRING", x), count
                elif x.name == "wchar_t":
                    return typedesc.Typedef("WSTRING", x), count
        if isinstance(t.typ, typedesc.FunctionType):
            return t.typ, count
        if isinstance(t.typ, typedesc.FundamentalType):
            if t.typ.name == "void":
                return typedesc.Typedef("c_void_p", t.typ), count
        if isinstance(t.typ, typedesc.PointerType):
            return self._inspect_PointerType(t.typ, count + 1)
        return t.typ, count + 1


class ImportedNamespaces(object):
    def __init__(self):
        self.data = {}

    def add(self, name1, name2=None, symbols=None):
        """Adds a namespace will be imported.

        Examples:
            >>> imports = ImportedNamespaces()
            >>> imports.add('datetime')
            >>> imports.add('ctypes', '*')
            >>> imports.add('decimal', 'Decimal')
            >>> imports.add('GUID', symbols={'GUID': 'comtypes'})
            >>> for name in ('COMMETHOD', 'DISPMETHOD', 'IUnknown', 'dispid',
            ...              'CoClass', 'BSTR', 'DISPPROPERTY'):
            ...     imports.add('comtypes', name)
            >>> imports.add('ctypes.wintypes')
            >>> print(imports.getvalue())
            from ctypes import *
            import datetime
            from decimal import Decimal
            from comtypes import (
                BSTR, CoClass, COMMETHOD, dispid, DISPMETHOD, DISPPROPERTY, GUID,
                IUnknown
            )
            import ctypes.wintypes
            >>> assert imports.get_symbols() == {
            ...     'Decimal', 'GUID', 'COMMETHOD', 'DISPMETHOD', 'IUnknown',
            ...     'dispid', 'CoClass', 'BSTR', 'DISPPROPERTY'
            ... }
        """
        if name2 is None:
            import_ = name1
            if not symbols:
                self.data[import_] = None
                return
            from_ = symbols[import_]
        else:
            from_, import_ = name1, name2
        self.data[import_] = from_

    def __contains__(self, item):
        """Returns item has already added.

        Examples:
            >>> imports = ImportedNamespaces()
            >>> imports.add('datetime')
            >>> imports.add('ctypes', '*')
            >>> 'datetime' in imports
            True
            >>> ('ctypes', '*') in imports
            True
            >>> 'os' in imports
            False
            >>> 'ctypes' in imports
            False
            >>> ('ctypes', 'c_int') in imports
            False
        """
        if isinstance(item, tuple):
            from_, import_ = item
        else:
            from_, import_ = None, item
        if import_ in self.data:
            return self.data[import_] == from_
        return False

    def get_symbols(self) -> Set[str]:
        names = set()
        for key, val in self.data.items():
            if val is None or key == "*":
                continue
            names.add(key)
        return names

    def _make_line(self, from_, imports):
        import_ = ", ".join(imports)
        code = "from %s import %s" % (from_, import_)
        if len(code) <= 80:
            return code
        wrapper = textwrap.TextWrapper(
            subsequent_indent="    ", initial_indent="    ", break_long_words=False
        )
        import_ = "\n".join(wrapper.wrap(import_))
        code = "from %s import (\n%s\n)" % (from_, import_)
        return code

    def getvalue(self):
        ns = {}
        lines = []
        for key, val in self.data.items():
            if val is None:
                ns[key] = val
            elif key == "*":
                lines.append("from %s import *" % val)
            else:
                ns.setdefault(val, set()).add(key)
        for key, val in ns.items():
            if val is None:
                lines.append("import %s" % key)
            else:
                names = sorted(val, key=lambda s: s.lower())
                lines.append(self._make_line(key, names))
        return "\n".join(lines)


class DeclaredNamespaces(object):
    def __init__(self):
        self.data = {}

    def add(self, alias, definition, comment=None):
        """Adds a namespace will be declared.

        Examples:
            >>> declarations = DeclaredNamespaces()
            >>> declarations.add('STRING', 'c_char_p')
            >>> declarations.add('_lcid', '0', 'change this if required')
            >>> print(declarations.getvalue())
            STRING = c_char_p
            _lcid = 0  # change this if required
            >>> assert declarations.get_symbols() == {
            ...     'STRING', '_lcid'
            ... }
        """
        self.data[(alias, definition)] = comment

    def get_symbols(self) -> Set[str]:
        names = set()
        for alias, _ in self.data.keys():
            names.add(alias)
        return names

    def getvalue(self):
        lines = []
        for (alias, definition), comment in self.data.items():
            code = "%s = %s" % (alias, definition)
            if comment:
                code = code + "  # %s" % comment
            lines.append(code)
        return "\n".join(lines)


class EnumerationNamespaces(object):
    def __init__(self):
        self.data: Dict[str, List[Tuple[str, int]]] = {}

    def add(self, enum_name: str, member_name: str, value: int) -> None:
        """Adds a namespace will be enumeration and its member.

        Examples:
            <BLANKLINE> is necessary for doctest
            >>> enums = EnumerationNamespaces()
            >>> assert not enums
            >>> enums.add('Foo', 'ham', 1)
            >>> assert enums
            >>> enums.add('Foo', 'spam', 2)
            >>> enums.add('Bar', 'bacon', 3)
            >>> assert 'Foo' in enums
            >>> assert 'Baz' not in enums
            >>> print(enums.to_intflags())
            class Foo(IntFlag):
                ham = 1
                spam = 2
            <BLANKLINE>
            <BLANKLINE>
            class Bar(IntFlag):
                bacon = 3
            >>> print(enums.to_constants())
            # values for enumeration 'Foo'
            ham = 1
            spam = 2
            Foo = c_int  # enum
            <BLANKLINE>
            # values for enumeration 'Bar'
            bacon = 3
            Bar = c_int  # enum
        """
        self.data.setdefault(enum_name, []).append((member_name, value))

    def __contains__(self, item: str) -> bool:
        return item in self.data

    def __bool__(self) -> bool:
        return bool(self.data)

    def get_symbols(self) -> Set[str]:
        return set(self.data)

    def to_constants(self) -> str:
        blocks = []
        for enum_name, enum_members in self.data.items():
            lines = []
            lines.append(f"# values for enumeration '{enum_name}'")
            for n, v in enum_members:
                lines.append(f"{n} = {v}")
            lines.append(f"{enum_name} = c_int  # enum")
            blocks.append("\n".join(lines))
        return "\n\n".join(blocks)

    def to_intflags(self) -> str:
        blocks = []
        for enum_name, enum_members in self.data.items():
            lines = []
            lines.append(f"class {enum_name}(IntFlag):")
            for member_name, value in enum_members:
                lines.append(f"    {member_name} = {value}")
            blocks.append("\n".join(lines))
        return "\n\n\n".join(blocks)
