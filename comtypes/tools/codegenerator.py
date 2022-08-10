# Code generator to generate code for everything contained in COM type
# libraries.
from __future__ import print_function
import os
import sys
if sys.version_info >= (3, 0):
    import io
else:
    import cStringIO as io
import keyword
import ctypes
import textwrap

from comtypes.tools import typedesc
import comtypes
import comtypes.client
import comtypes.client._generate

version = comtypes.__version__

import logging
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
    if struct.size is None: # incomplete struct
        return -1
    if struct.name in dont_assert_size:
        return None
    if struct.bases:
        size = struct.bases[0].size
        total_align = struct.bases[0].align
    else:
        size = 0
        total_align = 8 # in bits
    for i, f in enumerate(fields):
        if f.bits: # this code cannot handle bit field sizes.
            # print "##XXX FIXME"
            return -2 # XXX FIXME
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
    for pack in [None, 16*8, 8*8, 4*8, 2*8, 1*8]:
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

class Generator(object):

    def __init__(self, ofi, known_symbols=None):
        self.output = ofi
        self.stream = io.StringIO()
        self.imports = ImportedNamespaces()
        self.declarations = DeclaredNamespaces()
        self.known_symbols = known_symbols or {}

        self.done = set() # type descriptions that have been generated
        self.names = set() # names that have been generated
        self.last_item_class = False

    def generate(self, item):
        if item in self.done:
            return
        if isinstance(item, typedesc.StructureHead):
            name = getattr(item.struct, "name", None)
        else:
            name = getattr(item, "name", None)
        if name in self.known_symbols:
            self.imports.add(name, symbols=self.known_symbols)

            self.done.add(item)
            if isinstance(item, typedesc.Structure):
                self.done.add(item.get_head())
                self.done.add(item.get_body())
            return
        mth = getattr(self, type(item).__name__)
        # to avoid infinite recursion, we have to mark it as done
        # before actually generating the code.
        self.done.add(item)
        mth(item)

    def generate_all(self, items):
        for item in items:
            self.generate(item)

    def _make_relative_path(self, path1, path2):
        """path1 and path2 are pathnames.
        Return path1 as a relative path to path2, if possible.
        """
        path1 = os.path.abspath(path1)
        path2 = os.path.abspath(path2)
        common = os.path.commonprefix([os.path.normcase(path1),
                                       os.path.normcase(path2)])
        if not os.path.isdir(common):
            return path1
        if not common.endswith("\\"):
            return path1
        if not os.path.isdir(path2):
            path2 = os.path.dirname(path2)
        # strip the common prefix
        path1 = path1[len(common):]
        path2 = path2[len(common):]

        parts2 = path2.split("\\")
        return "..\\" * len(parts2) + path1

    def _generate_typelib_path(self, filename):
        # NOTE: the logic in this function appears completely different from that
        # of the handling of tlib (given as a string) in GetModule. There, relative
        # references are resolved wrt to the directory of the calling module. Here,
        # resolution is with respect to current working directory -- later to be
        # relativized to comtypes.gen.
        if filename is not None:
            if os.path.isabs(filename):
                # absolute path
                self.declarations.add("typelib_path", repr(filename))
            elif not os.path.dirname(filename) and not os.path.isfile(filename):
                # no directory given, and not in current directory.
                self.declarations.add("typelib_path", repr(filename))
            else:
                # relative path; make relative to comtypes.gen.
                path = self._make_relative_path(filename, comtypes.gen.__path__[0])
                self.imports.add('os')
                definition = "os.path.normpath(\n" \
                    "    os.path.abspath(os.path.join(os.path.dirname(__file__),\n" \
                    "                                 %r)))" % path
                self.declarations.add("typelib_path", definition)
                p = os.path.normpath(os.path.abspath(os.path.join(comtypes.gen.__path__[0],
                                                                  path)))
                assert os.path.isfile(p)

    def generate_code(self, items, filename):

        tlib_mtime = None

        if filename is not None:
            # get full path to DLL first (os.stat can't work with relative DLL paths properly)
            loaded_typelib = comtypes.typeinfo.LoadTypeLib(filename)
            full_filename = comtypes.tools.tlbparser.get_tlib_filename(
                loaded_typelib)

            while full_filename and not os.path.exists(full_filename):
                full_filename = os.path.split(full_filename)[0]

            if full_filename and os.path.isfile(full_filename):
                # get DLL timestamp at the moment of wrapper generation

                tlib_mtime = os.stat(full_filename).st_mtime

                if not full_filename.endswith(filename):
                    filename = full_filename

        self.filename = filename
        self.declarations.add("_lcid", "0", "change this if required")
        self._generate_typelib_path(filename)

        items = set(items)
        loops = 0
        while items:
            loops += 1
            self.more = set()
            self.generate_all(items)

            items |= self.more
            items -= self.done

        self.imports.add("ctypes", "*")  # HACK: wildcard import is so ugly.
        if tlib_mtime is not None:
            logger.debug("filename: \"%s\": tlib_mtime: %s", filename, tlib_mtime)
            self.imports.add('comtypes', '_check_version')

        if filename is not None:
            # Hm, what is the CORRECT encoding?
            print("# -*- coding: mbcs -*-", file=self.output)
            print(file=self.output)
        print(self.imports.getvalue(), file=self.output)
        print(file=self.output)
        print(self.declarations.getvalue(), file=self.output)
        print(file=self.output)
        print(self.stream.getvalue(), file=self.output)
        dunder_all = "__all__ = [%s]" % ", ".join(repr(str(n)) for n in self.names)
        if len(dunder_all) > 80:
            wrapper = textwrap.TextWrapper(subsequent_indent="    ",
                                           initial_indent="    ",
                                           break_long_words=False)
            names = ", ".join(repr(str(n)) for n in self.names)
            dunder_all = "__all__ = [\n%s\n]" % "\n".join(wrapper.wrap(names))
        print(dunder_all, file=self.output)
        print(file=self.output)
        if tlib_mtime is not None:
            print("_check_version(%r, %f)" % (version, tlib_mtime),
                    file=self.output)

        return loops

    def type_name(self, t):
        # Return a string, containing an expression which can be used
        # to refer to the type. Assumes the 'from ctypes import *'
        # namespace is available.
        if isinstance(t, typedesc.SAFEARRAYType):
            return "_midlSAFEARRAY(%s)" % self.type_name(t.typ)
        # if isinstance(t, typedesc.CoClass):
        #     return "%s._com_interfaces_[0]" % t.name
        if isinstance(t, typedesc.Typedef):
            return t.name
        if isinstance(t, typedesc.PointerType):
            if ASSUME_STRINGS:
                x = get_real_type(t.typ)
                if isinstance(x, typedesc.FundamentalType):
                    if x.name == "char":
                        return "STRING"
                    elif x.name == "wchar_t":
                        return "WSTRING"

            result = "POINTER(%s)" % self.type_name(t.typ)
            # XXX Better to inspect t.typ!
            if result.startswith("POINTER(WINFUNCTYPE"):
                return result[len("POINTER("):-1]
            if result.startswith("POINTER(CFUNCTYPE"):
                return result[len("POINTER("):-1]
            elif result == "POINTER(None)":
                return "c_void_p"
            return result
        elif isinstance(t, typedesc.ArrayType):
            return "%s * %s" % (self.type_name(t.typ), int(t.max)+1)
        elif isinstance(t, typedesc.FunctionType):
            args = [self.type_name(x) for x in [t.returns] + list(t.iterArgTypes())]
            if "__stdcall__" in t.attributes:
                return "WINFUNCTYPE(%s)" % ", ".join(args)
            else:
                return "CFUNCTYPE(%s)" % ", ".join(args)
        elif isinstance(t, typedesc.CvQualifiedType):
            # const and volatile are ignored
            return "%s" % self.type_name(t.typ)
        elif isinstance(t, typedesc.FundamentalType):
            return ctypes_names[t.name]
        elif isinstance(t, typedesc.Structure):
            return t.name
        elif isinstance(t, typedesc.Enumeration):
            if t.name:
                return t.name
            return "c_int" # enums are integers
        elif isinstance(t, typedesc.EnumValue):
            if keyword.iskeyword(t.name):
                return t.name + "_"
            return t.name
        elif isinstance(t, typedesc.External):
            # t.symbol_name - symbol to generate
            # t.tlib - the ITypeLib pointer to the typelibrary containing the symbols definition
            modname = comtypes.client._generate._name_module(t.tlib)
            return "%s.%s" % (modname, t.symbol_name)
        return t.name

    def need_VARIANT_imports(self, value):
        text = repr(value)
        if "Decimal(" in text:
            self.imports.add("decimal", "Decimal")
        if "datetime.datetime(" in text:
            self.imports.add("datetime")

    def need_GUID(self):
        if "GUID" in self.known_symbols:
            self.imports.add("GUID", symbols=self.known_symbols)

    def _to_docstring(self, orig, depth=1):
        # type: (str, int) -> str
        # increasing `depth` by one increases indentation by one
        indent = "    " * depth
        # some chars are replaced to avoid causing a `SyntaxError`
        repled = orig.replace("\\", r"\\").replace("\"", r"'")
        return '%s"""%s"""' % (indent, repled)

    _arraytypes = 0
    def ArrayType(self, tp):
        self._arraytypes += 1
        self.generate(get_real_type(tp.typ))
        self.generate(tp.typ)

    _enumvalues = 0
    def EnumValue(self, tp):
        self.last_item_class = False
        value = int(tp.value)
        if keyword.iskeyword(tp.name):
            # XXX use logging!
            if __warn_on_munge__:
                print("# Fixing keyword as EnumValue for %s" % tp.name)
        tp_name = self.type_name(tp)
        print("%s = %d" % (tp_name, value), file=self.stream)
        self.names.add(tp_name)
        self._enumvalues += 1

    _enumtypes = 0
    def Enumeration(self, tp):
        self._enumtypes += 1
        self.last_item_class = False
        if tp.name:
            print("# values for enumeration '%s'" % tp.name, file=self.stream)
        else:
            print("# values for unnamed enumeration", file=self.stream)
        # Some enumerations have the same name for the enum type
        # and an enum value.  Excel's XlDisplayShapes is such an example.
        # Since we don't have separate namespaces for the type and the values,
        # we generate the TYPE last, overwriting the value. XXX
        for item in tp.values:
            self.generate(item)
        if tp.name:
            print("%s = c_int  # enum" % tp.name, file=self.stream)
            self.names.add(tp.name)

    _typedefs = 0
    def Typedef(self, tp):
        self._typedefs += 1
        if type(tp.typ) in (typedesc.Structure, typedesc.Union):
            self.generate(tp.typ.get_head())
            self.more.add(tp.typ)
        else:
            self.generate(tp.typ)
        definition = self.type_name(tp.typ)
        if tp.name != definition:
            if definition in self.known_symbols:
                self.declarations.add(tp.name, definition)
            else:
                print("%s = %s" % (tp.name, definition), file=self.stream)
                self.last_item_class = False
        self.names.add(tp.name)

    def FundamentalType(self, item):
        pass # we should check if this is known somewhere

    def StructureHead(self, head):
        for struct in head.struct.bases:
            self.generate(struct.get_head())
            self.more.add(struct)
        if head.struct.location:
            self.last_item_class = False
            print("# %s %s" % head.struct.location, file=self.stream)
        basenames = [self.type_name(b) for b in head.struct.bases]
        if basenames:
            self.need_GUID()

            if not self.last_item_class:
                print(file=self.stream)
                print(file=self.stream)

            self.last_item_class = True

            method_names = [m.name for m in head.struct.members if type(m) is typedesc.Method]
            print("class %s(%s):" % (head.struct.name, ", ".join(basenames)), file=self.stream)
            print("    _iid_ = GUID('{}') # please look up iid and fill in!", file=self.stream)
            if "Enum" in method_names:
                print("    def __iter__(self):", file=self.stream)
                print("        return self.Enum()", file=self.stream)
            elif method_names == "Next Skip Reset Clone".split():
                print("    def __iter__(self):", file=self.stream)
                print("        return self", file=self.stream)
                print(file=self.stream)
                print("    def next(self):", file=self.stream)
                print("         arr, fetched = self.Next(1)", file=self.stream)
                print("         if fetched == 0:", file=self.stream)
                print("             raise StopIteration", file=self.stream)
                print("         return arr[0]", file=self.stream)

            print(file=self.stream)
            print(file=self.stream)

        else:
            methods = [m for m in head.struct.members if type(m) is typedesc.Method]

            if methods:
                # Hm. We cannot generate code for IUnknown...
                if not self.last_item_class:
                    print(file=self.stream)

                self.last_item_class = True
                print("assert 0, 'cannot generate code for IUnknown'", file=self.stream)
                print(file=self.stream)
                print(file=self.stream)
                print("class %s(_com_interface):" % head.struct.name, file=self.stream)
                print("    pass", file=self.stream)
                print(file=self.stream)
                print(file=self.stream)
            elif type(head.struct) == typedesc.Structure:
                if not self.last_item_class:
                    print(file=self.stream)
                    print(file=self.stream)

                self.last_item_class = True

                print("class %s(Structure):" % head.struct.name, file=self.stream)
                if hasattr(head.struct, "_recordinfo_"):
                    print("    _recordinfo_ = %r" % (head.struct._recordinfo_,), file=self.stream)
                else:
                    print("    pass", file=self.stream)
                print(file=self.stream)
                print(file=self.stream)
            elif type(head.struct) == typedesc.Union:
                if not self.last_item_class:
                    print(file=self.stream)
                    print(file=self.stream)

                self.last_item_class = True

                print("class %s(Union):" % head.struct.name, file=self.stream)
                print("    pass", file=self.stream)
                print(file=self.stream)
                print(file=self.stream)
        self.names.add(head.struct.name)

    _structures = 0
    def Structure(self, struct):
        self._structures += 1
        self.generate(struct.get_head())
        self.generate(struct.get_body())

    Union = Structure

    def StructureBody(self, body):
        fields = []
        methods = []
        for m in body.struct.members:
            if type(m) is typedesc.Field:
                fields.append(m)
                if type(m.typ) is typedesc.Typedef:
                    self.generate(get_real_type(m.typ))
                self.generate(m.typ)
            elif type(m) is typedesc.Method:
                methods.append(m)
                self.generate(m.returns)
                self.generate_all(m.iterArgTypes())
            elif type(m) is typedesc.Constructor:
                pass

        # we don't need _pack_ on Unions (I hope, at least), and not
        # on COM interfaces:
        if not methods:
            try:
                pack = calc_packing(body.struct, fields)
                if pack is not None:
                    self.last_item_class = False
                    print("%s._pack_ = %s" % (body.struct.name, pack), file=self.stream)
            except PackingError as details:
                # if packing fails, write a warning comment to the output.
                import warnings
                message = "Structure %s: %s" % (body.struct.name, details)
                warnings.warn(message, UserWarning)
                print("# WARNING: %s" % details, file=self.stream)
                self.last_item_class = False

        if fields:
            if body.struct.bases:
                assert len(body.struct.bases) == 1
                self.generate(body.struct.bases[0].get_body())
            # field definition normally span several lines.
            # Before we generate them, we need to 'import' everything they need.
            # So, call type_name for each field once,
            for f in fields:
                self.type_name(f.typ)

            if not self.last_item_class:
                print(file=self.stream)

            self.last_item_class = False

            print("%s._fields_ = [" % body.struct.name, file=self.stream)
            if body.struct.location:
                print("    # %s %s" % body.struct.location, file=self.stream)
            # unnamed fields will get autogenerated names "_", "_1". "_2", "_3", ...
            unnamed_index = 0
            for f in fields:
                if not f.name:
                    if unnamed_index:
                        fieldname = "_%d" % unnamed_index
                    else:
                        fieldname = "_"
                    unnamed_index += 1
                    print("    # Unnamed field renamed to '%s'" % fieldname, file=self.stream)
                else:
                    fieldname = f.name
                if f.bits is None:
                    print("    ('%s', %s)," % (fieldname, self.type_name(f.typ)), file=self.stream)
                else:
                    print("    ('%s', %s, %s)," % (fieldname, self.type_name(f.typ), f.bits), file=self.stream)
            print("]", file=self.stream)

            if body.struct.size is None:
                print(file=self.stream)
                msg = ("# The size provided by the typelib is incorrect.\n"
                       "# The size and alignment check for %s is skipped.")
                print(msg % body.struct.name, file=self.stream)
            elif body.struct.name not in dont_assert_size:
                print(file=self.stream)
                size = body.struct.size // 8
                print("assert sizeof(%s) == %s, sizeof(%s)" % \
                      (body.struct.name, size, body.struct.name), file=self.stream)
                align = body.struct.align // 8
                print("assert alignment(%s) == %s, alignment(%s)" % \
                      (body.struct.name, align, body.struct.name), file=self.stream)

        if methods:
            self.imports.add("comtypes", "COMMETHOD")
            # method definitions normally span several lines.
            # Before we generate them, we need to 'import' everything they need.
            # So, call type_name for each field once,
            for m in methods:
                self.type_name(m.returns)
                for a in m.iterArgTypes():
                    self.type_name(a)

            if not self.last_item_class:
                print(file=self.stream)

            self.last_item_class = False
            print("%s._methods_ = [" % body.struct.name, file=self.stream)
            if body.struct.location:
                print("# %s %s" % body.struct.location, file=self.stream)

            for m in methods:
                if m.location:
                    print("    # %s %s" % m.location, file=self.stream)
                print(
                    (
                        "    COMMETHOD(\n"
                        "        [], \n"
                        "        %s,\n"
                        "        '%s',\n"
                    ) % (self.type_name(m.returns), m.name),
                    file=self.stream
                )
                for a in m.iterArgTypes():
                    print("        ([], %s),\n" % self.type_name(a), file=self.stream)
                    print("    ),", file=self.stream)
            print("]", file=self.stream)

    ################################################################
    # top-level typedesc generators
    #
    def TypeLib(self, lib):
        # lib.name, lib.gui, lib.major, lib.minor, lib.doc

        # Hm, in user code we have to write:
        # class MyServer(COMObject, ...):
        #     _com_interfaces_ = [MyTypeLib.IInterface]
        #     _reg_typelib_ = MyTypeLib.Library._reg_typelib_
        #                               ^^^^^^^
        # Should the '_reg_typelib_' attribute be at top-level in the
        # generated code, instead as being an attribute of the
        # 'Library' symbol?
        if not self.last_item_class:
            print(file=self.stream)
            print(file=self.stream)

        self.last_item_class = True

        print("class Library(object):", file=self.stream)
        if lib.doc:
            print(self._to_docstring(lib.doc), file=self.stream)

        if lib.name:
            print("    name = %r" % lib.name, file=self.stream)

        print("    _reg_typelib_ = (%r, %r, %r)" % (lib.guid, lib.major, lib.minor), file=self.stream)
        print(file=self.stream)
        print(file=self.stream)

    def External(self, ext):
        modname = comtypes.client._generate._name_module(ext.tlib)
        if modname not in self.imports:
            comtypes.client.GetModule(ext.tlib)
            self.imports.add(modname)

    def Constant(self, tp):
        self.last_item_class = False
        print("%s = %r  # Constant %s" % (tp.name,
                                         tp.value,
                                         self.type_name(tp.typ)), file=self.stream)
        self.names.add(tp.name)

    def SAFEARRAYType(self, sa):
        self.generate(sa.typ)
        self.imports.add("comtypes.automation", "_midlSAFEARRAY")

    _pointertypes = 0
    def PointerType(self, tp):
        self._pointertypes += 1
        if type(tp.typ) is typedesc.ComInterface:
            # this defines the class
            self.generate(tp.typ.get_head())
            # this defines the _methods_
            self.more.add(tp.typ)
        elif type(tp.typ) is typedesc.PointerType:
            self.generate(tp.typ)
        elif type(tp.typ) in (typedesc.Union, typedesc.Structure):
            self.generate(tp.typ.get_head())
            self.more.add(tp.typ)
        elif type(tp.typ) is typedesc.Typedef:
            self.generate(tp.typ)
        else:
            self.generate(tp.typ)
        if not ASSUME_STRINGS:
            return
        real_type = get_real_type(tp.typ)
        if isinstance(real_type, typedesc.FundamentalType):
            if real_type.name == "char":
                self.declarations.add("STRING", "c_char_p")
            elif real_type.name == "wchar_t":
                self.declarations.add("WSTRING", "c_wchar_p")

    def CoClass(self, coclass):
        self.need_GUID()
        self.imports.add("comtypes", "CoClass")
        if not self.last_item_class:
            print(file=self.stream)
            print(file=self.stream)

        self.last_item_class = True

        print("class %s(CoClass):" % coclass.name, file=self.stream)
        doc = getattr(coclass, "doc", None)
        if doc:
            print(self._to_docstring(doc), file=self.stream)
        print("    _reg_clsid_ = GUID(%r)" % coclass.clsid, file=self.stream)
        print("    _idlflags_ = %s" % coclass.idlflags, file=self.stream)
        if self.filename is not None:
            print("    _typelib_path_ = typelib_path", file=self.stream)
        # X print >> self.stream, "POINTER(%s).__ctypes_from_outparam__ = wrap" % coclass.name

        libid = coclass.tlibattr.guid
        wMajor, wMinor = coclass.tlibattr.wMajorVerNum, coclass.tlibattr.wMinorVerNum
        print("    _reg_typelib_ = (%r, %s, %s)" % (str(libid), wMajor, wMinor), file=self.stream)
        print(file=self.stream)
        print(file=self.stream)

        for itf, idlflags in coclass.interfaces:
            self.generate(itf.get_head())
        implemented = []
        sources = []
        for item in coclass.interfaces:
            # item is (interface class, impltypeflags)
            if item[1] & 2: # IMPLTYPEFLAG_FSOURCE
                # source interface
                where = sources
            else:
                # sink interface
                where = implemented
            if item[1] & 1: # IMPLTYPEFLAG_FDEAULT
                # The default interface should be the first item on the list
                where.insert(0, self.type_name(item[0]))
            else:
                where.append(self.type_name(item[0]))

        if implemented:
            self.last_item_class = False
            print("%s._com_interfaces_ = [%s]" % (coclass.name, ", ".join(implemented)), file=self.stream)
        if sources:
            self.last_item_class = False
            print("%s._outgoing_interfaces_ = [%s]" % (coclass.name, ", ".join(sources)), file=self.stream)

        self.names.add(coclass.name)

    def ComInterface(self, itf):
        self.generate(itf.get_head())
        self.generate(itf.get_body())
        self.names.add(itf.name)

    def _is_enuminterface(self, itf):
        # Check if this is an IEnumXXX interface
        if not itf.name.startswith("IEnum"):
            return False
        member_names = [mth.name for mth in itf.members]
        for name in ("Next", "Skip", "Reset", "Clone"):
            if name not in member_names:
                return False
        return True

    def ComInterfaceHead(self, head):
        if head.itf.name in self.known_symbols:
            return
        base = head.itf.base
        if head.itf.base is None:
            # we don't beed to generate IUnknown
            return
        self.generate(base.get_head())
        self.more.add(base)
        basename = self.type_name(head.itf.base)

        self.need_GUID()

        if not self.last_item_class:
            print(file=self.stream)
            print(file=self.stream)

        self.last_item_class = True

        print("class %s(%s):" % (head.itf.name, basename), file=self.stream)
        doc = getattr(head.itf, "doc", None)
        if doc:
            print(self._to_docstring(doc), file=self.stream)

        print("    _case_insensitive_ = True", file=self.stream)
        print("    _iid_ = GUID(%r)" % head.itf.iid, file=self.stream)
        print("    _idlflags_ = %s" % head.itf.idlflags, file=self.stream)

        if self._is_enuminterface(head.itf):
            print(file=self.stream)
            print("    def __iter__(self):", file=self.stream)
            print("        return self", file=self.stream)
            print(file=self.stream)

            print("    def next(self):", file=self.stream)
            print("        item, fetched = self.Next(1)", file=self.stream)
            print("        if fetched:", file=self.stream)
            print("            return item", file=self.stream)
            print("        raise StopIteration", file=self.stream)
            print(file=self.stream)

            print("    def __getitem__(self, index):", file=self.stream)
            print("        self.Reset()", file=self.stream)
            print("        self.Skip(index)", file=self.stream)
            print("        item, fetched = self.Next(1)", file=self.stream)
            print("        if fetched:", file=self.stream)
            print("            return item", file=self.stream)
            print("        raise IndexError(index)", file=self.stream)

        print(file=self.stream)
        print(file=self.stream)

    def ComInterfaceBody(self, body):
        # The base class must be fully generated, including the
        # _methods_ list.
        self.generate(body.itf.base)

        # make sure we can generate the body
        for m in body.itf.members:
            for a in m.arguments:
                self.generate(a[0])
            self.generate(m.returns)

        if not self.last_item_class:
            print(file=self.stream)

        self.last_item_class = False
        print("%s._methods_ = [" % body.itf.name, file=self.stream)
        for m in body.itf.members:
            if isinstance(m, typedesc.ComMethod):
                self.make_ComMethod(m, "dual" in body.itf.idlflags)
            else:
                raise TypeError("what's this?")

        print("]", file=self.stream)
        print(file=self.stream)
        print("################################################################", file=self.stream)
        print("# code template for %s implementation" % body.itf.name, file=self.stream)
        print("# class %s_Impl(object):" % body.itf.name, file=self.stream)

        methods = {}
        for m in body.itf.members:
            if isinstance(m, typedesc.ComMethod):
                # m.arguments is a sequence of tuples:
                # (argtype, argname, idlflags, docstring)
                # Some typelibs have unnamed method parameters!
                inargs = [a[1] or '<unnamed>' for a in m.arguments
                        if not 'out' in a[2]]
                outargs = [a[1] or '<unnamed>' for a in m.arguments
                           if 'out' in a[2]]
                if 'propget' in m.idlflags:
                    methods.setdefault(m.name, [0, inargs, outargs, m.doc])[0] |= 1
                elif 'propput' in m.idlflags:
                    methods.setdefault(m.name, [0, inargs[:-1], inargs[-1:], m.doc])[0] |= 2
                else:
                    methods[m.name] = [0, inargs, outargs, m.doc]

        for name, (typ, inargs, outargs, doc) in methods.items():
            if typ == 0: # method
                print("#     def %s(%s):" % (name, ", ".join(["self"] + inargs)), file=self.stream)
                print("#         %r" % (doc or "-no docstring-"), file=self.stream)
                print("#         #return %s" % (", ".join(outargs)), file=self.stream)
            elif typ == 1: # propget
                print("#     @property", file=self.stream)
                print("#     def %s(%s):" % (name, ", ".join(["self"] + inargs)), file=self.stream)
                print("#         %r" % (doc or "-no docstring-"), file=self.stream)
                print("#         #return %s" % (", ".join(outargs)), file=self.stream)
            elif typ == 2: # propput
                print("#     def _set(%s):" % ", ".join(["self"] + inargs + outargs), file=self.stream)
                print("#         %r" % (doc or "-no docstring-"), file=self.stream)
                print("#     %s = property(fset = _set, doc = _set.__doc__)" % name, file=self.stream)
            elif typ == 3: # propget + propput
                print("#     def _get(%s):" % ", ".join(["self"] + inargs), file=self.stream)
                print("#         %r" % (doc or "-no docstring-"), file=self.stream)
                print("#         #return %s" % (", ".join(outargs)), file=self.stream)
                print("#     def _set(%s):" % ", ".join(["self"] + inargs + outargs), file=self.stream)
                print("#         %r" % (doc or "-no docstring-"), file=self.stream)
                print("#     %s = property(_get, _set, doc = _set.__doc__)" % name, file=self.stream)
            else:
                raise RuntimeError("BUG")
            print("#", file=self.stream)

    def DispInterface(self, itf):
        self.generate(itf.get_head())
        self.generate(itf.get_body())
        self.names.add(itf.name)

    def DispInterfaceHead(self, head):
        self.generate(head.itf.base)
        basename = self.type_name(head.itf.base)

        self.need_GUID()
        if not self.last_item_class:
            print(file=self.stream)
            print(file=self.stream)

        self.last_item_class = True

        print("class %s(%s):" % (head.itf.name, basename), file=self.stream)
        print("    _case_insensitive_ = True", file=self.stream)
        doc = getattr(head.itf, "doc", None)
        if doc:
            print("    %r" % doc, file=self.stream)
        print("    _iid_ = GUID(%r)" % head.itf.iid, file=self.stream)
        print("    _idlflags_ = %s" % head.itf.idlflags, file=self.stream)
        print("    _methods_ = []", file=self.stream)
        print(file=self.stream)
        print(file=self.stream)

    def DispInterfaceBody(self, body):
        # make sure we can generate the body
        for m in body.itf.members:
            if isinstance(m, typedesc.DispMethod):
                for a in m.arguments:
                    self.generate(a[0])
                self.generate(m.returns)
            elif isinstance(m, typedesc.DispProperty):
                self.generate(m.typ)
            else:
                raise TypeError(m)

        if not self.last_item_class:
            print(file=self.stream)

        self.last_item_class = False

        print("%s._disp_methods_ = [" % body.itf.name, file=self.stream)
        for m in body.itf.members:
            if isinstance(m, typedesc.DispMethod):
                self.make_DispMethod(m)
            elif isinstance(m, typedesc.DispProperty):
                self.make_DispProperty(m)
            else:
                raise TypeError(m)
        print("]", file=self.stream)

    ################################################################
    # non-toplevel method generators
    #
    def make_ComMethod(self, m, isdual):
        self.imports.add("comtypes", "COMMETHOD")
        # typ, name, idlflags, default
        if isdual:
            self.imports.add("comtypes", "dispid")
            idlflags = [dispid(m.memid)] + m.idlflags
        else:
            # We don't include the dispid for non-dispatch COM interfaces
            idlflags = m.idlflags
        if __debug__ and m.doc:
            self.imports.add("comtypes", "helpstring")
            idlflags.insert(1, helpstring(m.doc))

        self.last_item_class = False
        if not m.arguments:
            code = "    COMMETHOD(%r, %s, '%s')," % (idlflags, self.type_name(m.returns), m.name)
            if len(code) > 80:
                code = (
                    "    COMMETHOD(\n"
                    "        %r,\n"
                    "        %s,\n"
                    "        '%s',\n"
                    "    ),"
                ) % (idlflags, self.type_name(m.returns), m.name)

            print(code, file=self.stream)
        else:
            code = (
                "    COMMETHOD(\n"
                "        %r,\n"
                "        %s,\n"
                "        '%s',"
            ) % (idlflags, self.type_name(m.returns), m.name)
            print(code, file=self.stream)
            arglist = []
            for typ, name, idlflags, default in m.arguments:
                type_name = self.type_name(typ)
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
                if isinstance(typ, typedesc.ComInterface):
                    self.declarations.add("OPENARRAY", "POINTER(c_ubyte)",
                        "hack, see comtypes/tools/codegenerator.py")
                    type_name = "OPENARRAY"
                    if 'in' not in idlflags:
                        idlflags.append('in')
                if 'lcid' in idlflags:# and 'in' in idlflags:
                    default = lcid
                if default is not None:
                    self.need_VARIANT_imports(default)

                    code = "        (%r, %s, '%s', %r)" % (idlflags, type_name, name, default)

                    if len(code) > 80:
                        code = (
                            "        (\n"
                            "            %r,\n"
                            "            %s,\n"
                            "            '%s',\n"
                            "            %r\n"
                            "        )"
                        ) % (idlflags, type_name, name, default)
                else:
                    code = "        (%r, %s, '%s')" % (idlflags, type_name, name)
                    if len(code) > 80:
                        code = (
                            "        (\n"
                            "            %r,\n"
                            "            %s,\n"
                            "            '%s',\n"
                            "        )"
                        ) % (idlflags, type_name, name)

                arglist.append(code)

            print(",\n".join(arglist), file=self.stream)
            print("    ),", file=self.stream)

    def make_DispMethod(self, m):
        self.imports.add("comtypes", "DISPMETHOD")
        self.imports.add("comtypes", "dispid")
        idlflags = [dispid(m.dispid)] + m.idlflags
        if __debug__ and m.doc:
            self.imports.add("comtypes", "helpstring")
            idlflags.insert(1, helpstring(m.doc))

        self.last_item_class = False

        # typ, name, idlflags, default
        if not m.arguments:
            code = "    DISPMETHOD(%r, %s, '%s')," % (idlflags, self.type_name(m.returns), m.name)
            if len(code) > 80:
                code = (
                    "    DISPMETHOD(\n"
                    "        %r,\n"
                    "        %s,\n"
                    "        '%s'\n"
                    "    ),"
                ) % (idlflags, self.type_name(m.returns), m.name)

            print(code, file=self.stream)
        else:
            code = (
                "    DISPMETHOD(\n"
                "        %r,\n"
                "        %s,\n"
                "        '%s',"
                ) % (idlflags, self.type_name(m.returns), m.name)

            print(code, file=self.stream)

            arglist = []
            for typ, name, idlflags, default in m.arguments:
                self.need_VARIANT_imports(default)
                if default is not None:
                    code = "        (%r, %s, '%s', %r)" % (idlflags, self.type_name(typ), name, default)
                    if len(code) > 80:
                        code = (
                            "        (\n"
                            "            %r,\n"
                            "            %s,\n"
                            "            '%s',\n"
                            "            %r\n"
                            "        )"
                        ) % (idlflags, self.type_name(typ), name, default)


                else:
                    code = "        (%r, %s, '%s')" % (idlflags, self.type_name(typ), name)

                    if len(code) > 80:
                        code = (
                            "        (\n"
                            "            %r,\n"
                            "            %s,\n"
                            "            '%s'\n"
                            "        )"
                        ) % (idlflags, self.type_name(typ), name)

                arglist.append(code)

            print(",\n".join(arglist), file=self.stream)
            print("    ),", file=self.stream)

    def make_DispProperty(self, prop):
        self.imports.add("comtypes", "DISPPROPERTY")
        self.imports.add("comtypes", "dispid")
        idlflags = [dispid(prop.dispid)] + prop.idlflags
        if __debug__ and prop.doc:
            self.imports.add("comtypes", "helpstring")
            idlflags.insert(1, helpstring(prop.doc))

        self.last_item_class = False
        code = "    DISPPROPERTY(%r, %s, '%s')," % (idlflags, self.type_name(prop.typ), prop.name)
        if len(code) > 80:
            code = (
                "    DISPPROPERTY(\n"
                "        %r,\n"
                "        %s,\n"
                "        '%s'\n"
                "    ),"
            ) % (idlflags, self.type_name(prop.typ), prop.name)

        print(code, file=self.stream)


class ImportedNamespaces(object):
    def __init__(self):
        if sys.version_info >= (3, 7):
            self.data = {}
        else:
            from collections import OrderedDict
            self.data = OrderedDict()

    def add(self, name1, name2=None, symbols=None):
        """Adds a namespace will be imported.

        Examples:
            >>> imports = ImportedNamespaces()
            >>> imports.add('datetime')
            >>> imports.add('ctypes', '*')
            >>> imports.add('decimal', 'Decimal')
            >>> imports.add('GUID', symbols={'GUID': 'comtypes'})
            >>> for name in ('COMMETHOD', 'IUnknown', 'dispid', 'CoClass',
            ...              'BSTR', 'DISPPROPERTY'):
            ...     imports.add('comtypes', name)
            >>> imports.add('ctypes.wintypes')
            >>> print(imports.getvalue())
            from ctypes import *
            import datetime
            from decimal import Decimal
            from comtypes import BSTR, CoClass, COMMETHOD, dispid, DISPPROPERTY, \\
                GUID, IUnknown
            import ctypes.wintypes
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

    def _make_line(self, import_, from_=None):
        if from_ is None:
            return "import %s" % import_
        code = "from %s import %s" % (from_, import_)
        if len(code) > 80:
            wrapper = textwrap.TextWrapper(subsequent_indent="    ",
                                           break_long_words=False)
            code = " \\\n".join(wrapper.wrap(code))
        return code

    def getvalue(self):
        ns = {}
        lines = []
        for key, val in self.data.items():
            if val is None:
                ns[key] = val
            elif key == "*":
                lines.append(self._make_line("*", val))
            else:
                ns.setdefault(val, set()).add(key)
        for key, val in ns.items():
            if val is None:
                lines.append(self._make_line(key))
            else:
                names = ", ".join(sorted(val, key=lambda s: s.lower()))
                lines.append(self._make_line(names, key))
        return "\n".join(lines)


class DeclaredNamespaces(object):
    def __init__(self):
        if sys.version_info >= (3, 7):
            self.data = {}
        else:
            from collections import OrderedDict
            self.data = OrderedDict()

    def add(self, alias, definition, comment=None):
        """Adds a namespace will be declared.

        Examples:
            >>> declarations = DeclaredNamespaces()
            >>> declarations.add('STRING', 'c_char_p')
            >>> declarations.add('_lcid', '0', 'change this if required')
            >>> print(declarations.getvalue())
            STRING = c_char_p
            _lcid = 0  # change this if required
        """
        self.data[(alias, definition)] = comment

    def getvalue(self):
        lines = []
        for (alias, definition), comment in self.data.items():
            code = "%s = %s" % (alias, definition)
            if comment:
                code = code + "  # %s" % comment
            lines.append(code)
        return "\n".join(lines)


# shortcut for development
if __name__ == "__main__":
    from . import tlbparser
    tlbparser.main()
