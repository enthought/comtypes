# Create ctypes wrapper code for abstract type descriptions.
# Type descriptions are collections of typedesc instances.

import typedesc, sys, types

ASSUME_STRINGS = True

try:
    set
except NameError:
    from sets import Set as set

try:
    sorted
except NameError:
    def sorted(seq, cmp):
        seq = list(seq)
        seq.sort(cmp)
        return seq

try:
    import cStringIO as StringIO
except ImportError:
    import StringIO


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

################

def storage(t):
    # return the size and alignment of a type
    if isinstance(t, typedesc.Typedef):
        return storage(t.typ)
    elif isinstance(t, typedesc.ArrayType):
        s, a = storage(t.typ)
        return s * (int(t.max) - int(t.min) + 1), a
    return int(t.size), int(t.align)

class PackingError(Exception):
    pass

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
##            print "##XXX FIXME"
            return -2 # XXX FIXME
        s, a = storage(f.typ)
        if pack is not None:
            a = min(pack, a)
        if size % a:
            size += a - size % a
        if isStruct:
            if size != f.offset:
                raise PackingError, "field %s offset (%s/%s)" % (f.name, size, f.offset)
            size += s
        else:
            size = max(size, s)
        total_align = max(total_align, a)
    if total_align != struct.align:
        raise PackingError, "total alignment (%s/%s)" % (total_align, struct.align)
    a = total_align
    if pack is not None:
        a = min(pack, a)
    if size % a:
        size += a - size % a
    if size != struct.size:
        raise PackingError, "total size (%s/%s)" % (size, struct.size)

def calc_packing(struct, fields):
    # try several packings, starting with unspecified packing
    isStruct = isinstance(struct, typedesc.Structure)
    for pack in [None, 16*8, 8*8, 4*8, 2*8, 1*8]:
        try:
            _calc_packing(struct, fields, pack, isStruct)
        except PackingError, details:
            continue
        else:
            if pack is None:
                return None
            return pack/8
    raise PackingError, "PACKING FAILED: %s" % details

def decode_value(init):
    # decode init value from gccxml
    if init[0] == "0":
        return int(init, 16) # hex integer
    elif init[0] == "'":
        return eval(init) # character
    elif init[0] == '"':
        return eval(init) # string
    return int(init) # integer

def get_real_type(tp):
    if type(tp) is typedesc.Typedef:
        return get_real_type(tp.typ)
    elif isinstance(tp, typedesc.CvQualifiedType):
        return get_real_type(tp.typ)
    return tp

# XXX These should be filtered out in gccxmlparser.
dont_assert_size = set(
    [
    "__si_class_type_info_pseudo",
    "__class_type_info_pseudo",
    ]
    )

################################################################

class Generator(object):
    def __init__(self, output,
                 use_decorators=False,
                 known_symbols=None,
                 searched_dlls=None):
        self.output = output
        self.stream = StringIO.StringIO()
        self.imports = StringIO.StringIO()
##        self.stream = self.imports = self.output
        self.use_decorators = use_decorators
        self.known_symbols = known_symbols or {}
        self.searched_dlls = searched_dlls or []

        self.done = set() # type descriptions that have been generated
        self.names = set() # names that have been generated

    def init_value(self, t, init):
        tn = self.type_name(t, False)
        if tn in ["c_ulonglong", "c_ulong", "c_uint", "c_ushort", "c_ubyte"]:
            return decode_value(init)
        elif tn in ["c_longlong", "c_long", "c_int", "c_short", "c_byte"]:
            return decode_value(init)
        elif tn in ["c_float", "c_double"]:
            return float(init)
        elif tn == "POINTER(c_char)":
            if init[0] == '"':
                value = eval(init)
            else:
                value = int(init, 16)
            return value
        elif tn == "POINTER(c_wchar)":
            if init[0] == '"':
                value = eval(init)
            else:
                value = int(init, 16)
            if isinstance(value, str):
                value = value[:-1] # gccxml outputs "D\000S\000\000" for L"DS"
                value = value.decode("utf-16") # XXX Is this correct?
            return value
        elif tn == "c_void_p":
            if init[0] == "0":
                value = int(init, 16)
            else:
                value = int(init) # hm..
            # Hm, ctypes represents them as SIGNED int
            return value
        elif tn == "c_char":
            return decode_value(init)
        elif tn == "c_wchar":
            value = decode_value(init)
            if isinstance(value, int):
                return unichr(value)
            return value
        elif tn.startswith("POINTER("):
            # Hm, POINTER(HBITMAP__) for example
            return decode_value(init)
        else:
            raise ValueError, "cannot decode %s(%r)" % (tn, init)

    def type_name(self, t, generate=True):
        # Return a string containing an expression that can be used to
        # refer to the type. Assumes the 'from ctypes import *'
        # namespace is available.
        if isinstance(t, typedesc.Typedef):
            return t.name
        if isinstance(t, typedesc.PointerType):
            if ASSUME_STRINGS:
                x = get_real_type(t.typ)
                if isinstance(x, typedesc.FundamentalType):
                    if x.name == "char":
                        self.need_STRING()
                        return "STRING"
                    elif x.name == "wchar_t":
                        self.need_WSTRING()
                        return "WSTRING"

            result = "POINTER(%s)" % self.type_name(t.typ, generate)
            # XXX Better to inspect t.typ!
            if result.startswith("POINTER(WINFUNCTYPE"):
                return result[len("POINTER("):-1]
            if result.startswith("POINTER(CFUNCTYPE"):
                return result[len("POINTER("):-1]
            elif result == "POINTER(None)":
                return "c_void_p"
            return result
        elif isinstance(t, typedesc.ArrayType):
            return "%s * %s" % (self.type_name(t.typ, generate), int(t.max)+1)
        elif isinstance(t, typedesc.FunctionType):
            args = [self.type_name(x, generate) for x in [t.returns] + list(t.iterArgTypes())]
            if "__stdcall__" in t.attributes:
                return "WINFUNCTYPE(%s)" % ", ".join(args)
            else:
                return "CFUNCTYPE(%s)" % ", ".join(args)
        elif isinstance(t, typedesc.CvQualifiedType):
            # const and volatile are ignored
            return "%s" % self.type_name(t.typ, generate)
        elif isinstance(t, typedesc.FundamentalType):
            return ctypes_names[t.name]
        elif isinstance(t, typedesc.Structure):
            return t.name
        elif isinstance(t, typedesc.Enumeration):
            if t.name:
                return t.name
            return "c_int" # enums are integers
        elif isinstance(t, typedesc.Typedef):
            return t.name
        return t.name

    ################################################################

    def Alias(self, alias):
        if alias.typ is not None: # we can resolve it
            self.generate(alias.typ)
            if alias.alias in self.names:
                print >> self.stream, "%s = %s # alias" % (alias.name, alias.alias)
                self.names.add(alias.name)
                return
        # we cannot resolve it
        print >> self.stream, "# %s = %s # alias" % (alias.name, alias.alias)
        print "# unresolved alias: %s = %s" % (alias.name, alias.alias)
            

    def Macro(self, macro):
        # We don't know if we can generate valid, error free Python
        # code All we can do is to try to compile the code.  If the
        # compile fails, we know it cannot work, so we generate
        # commented out code.  If it succeeds, it may fail at runtime.
        code = "def %s%s: return %s # macro" % (macro.name, macro.args, macro.body)
        try:
            compile(code, "<string>", "exec")
        except SyntaxError:
            print >> self.stream, "#", code
        else:
            print >> self.stream, code
            self.names.add(macro.name)

    def StructureHead(self, head):
        for struct in head.struct.bases:
            self.generate(struct.get_head())
            self.more.add(struct)
        if head.struct.location:
            print >> self.stream, "# %s %s" % head.struct.location
        basenames = [self.type_name(b) for b in head.struct.bases]
        if basenames:
            self.need_GUID()
            method_names = [m.name for m in head.struct.members if type(m) is typedesc.Method]
            print >> self.stream, "class %s(%s):" % (head.struct.name, ", ".join(basenames))
            print >> self.stream, "    _iid_ = GUID('{}') # please look up iid and fill in!"
            if "Enum" in method_names:
                print >> self.stream, "    def __iter__(self):"
                print >> self.stream, "        return self.Enum()"
            elif method_names == "Next Skip Reset Clone".split():
                print >> self.stream, "    def __iter__(self):"
                print >> self.stream, "        return self"
                print >> self.stream
                print >> self.stream, "    def next(self):"
                print >> self.stream, "         arr, fetched = self.Next(1)"
                print >> self.stream, "         if fetched == 0:"
                print >> self.stream, "             raise StopIteration"
                print >> self.stream, "         return arr[0]"
        else:
            methods = [m for m in head.struct.members if type(m) is typedesc.Method]
            if methods:
                # Hm. We cannot generate code for IUnknown...
                print >> self.stream, "assert 0, 'cannot generate code for IUnknown'"
                print >> self.stream, "class %s(_com_interface):" % head.struct.name
            elif type(head.struct) == typedesc.Structure:
                print >> self.stream, "class %s(Structure):" % head.struct.name
            elif type(head.struct) == typedesc.Union:
                print >> self.stream, "class %s(Union):" % head.struct.name
            print >> self.stream, "    pass"
        self.names.add(head.struct.name)

    _structures = 0
    def Structure(self, struct):
        self._structures += 1
        self.generate(struct.get_head())
        self.generate(struct.get_body())

    Union = Structure
        
    _typedefs = 0
    def Typedef(self, tp):
        self._typedefs += 1
        if type(tp.typ) in (typedesc.Structure, typedesc.Union):
            self.generate(tp.typ.get_head())
            self.more.add(tp.typ)
        else:
            self.generate(tp.typ)
        if self.type_name(tp.typ) in self.known_symbols:
            stream = self.imports
        else:
            stream = self.stream
        if tp.name != self.type_name(tp.typ):
            print >> stream, "%s = %s" % \
                  (tp.name, self.type_name(tp.typ))
        self.names.add(tp.name)

    _arraytypes = 0
    def ArrayType(self, tp):
        self._arraytypes += 1
        self.generate(get_real_type(tp.typ))
        self.generate(tp.typ)

    _functiontypes = 0
    def FunctionType(self, tp):
        self._functiontypes += 1
        self.generate(tp.returns)
        self.generate_all(tp.iterArgTypes())
        
    _pointertypes = 0
    def PointerType(self, tp):
        self._pointertypes += 1
        if type(tp.typ) is typedesc.PointerType:
            self.generate(tp.typ)
        elif type(tp.typ) in (typedesc.Union, typedesc.Structure):
            self.generate(tp.typ.get_head())
            self.more.add(tp.typ)
        elif type(tp.typ) is typedesc.Typedef:
            self.generate(tp.typ)
        else:
            self.generate(tp.typ)

    def CvQualifiedType(self, tp):
        self.generate(tp.typ)

    _variables = 0
    def Variable(self, tp):
        self._variables += 1
        if tp.init is None:
            # wtypes.h contains IID_IProcessInitControl, for example
            return
        try:
            value = self.init_value(tp.typ, tp.init)
        except (TypeError, ValueError), detail:
            print "Could not init", tp.name, tp.init, detail
##            raise
            return
        print >> self.stream, \
              "%s = %r # Variable %s" % (tp.name,
                                         value,
                                         self.type_name(tp.typ, False))
        self.names.add(tp.name)

    _enumvalues = 0
    def EnumValue(self, tp):
        value = int(tp.value)
        print >> self.stream, \
              "%s = %d" % (tp.name, value)
        self.names.add(tp.name)
        self._enumvalues += 1

    _enumtypes = 0
    def Enumeration(self, tp):
        self._enumtypes += 1
        print >> self.stream
        if tp.name:
            print >> self.stream, "# values for enumeration '%s'" % tp.name
        else:
            print >> self.stream, "# values for unnamed enumeration"
        # Some enumerations have the same name for the enum type
        # and an enum value.  Excel's XlDisplayShapes is such an example.
        # Since we don't have separate namespaces for the type and the values,
        # we generate the TYPE last, overwriting the value. XXX
        for item in tp.values:
            self.generate(item)
        if tp.name:
            print >> self.stream, "%s = c_int # enum" % tp.name
            self.names.add(tp.name)


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
                    print >> self.stream, "%s._pack_ = %s" % (body.struct.name, pack)
            except PackingError, details:
                # if packing fails, write a warning comment to the output.
                import warnings
                message = "Structure %s: %s" % (body.struct.name, details)
                warnings.warn(message, UserWarning)
                print >> self.stream, "# WARNING: %s" % details

        if fields:
            if body.struct.bases:
                assert len(body.struct.bases) == 1
                self.generate(body.struct.bases[0].get_body())
            # field definition normally span several lines.
            # Before we generate them, we need to 'import' everything they need.
            # So, call type_name for each field once,
            for f in fields:
                self.type_name(f.typ)
            print >> self.stream, "%s._fields_ = [" % body.struct.name
            if body.struct.location:
                print >> self.stream, "    # %s %s" % body.struct.location
            # unnamed fields will get autogenerated names "_", "_1". "_2", "_3", ...
            unnamed_index = 0
            for f in fields:
                if not f.name:
                    if unnamed_index:
                        fieldname = "_%d" % unnamed_index
                    else:
                        fieldname = "_"
                    unnamed_index += 1
                    print >> self.stream, "    # Unnamed field renamed to '%s'" % fieldname
                else:
                    fieldname = f.name
                if f.bits is None:
                    print >> self.stream, "    ('%s', %s)," % (fieldname, self.type_name(f.typ))
                else:
                    print >> self.stream, "    ('%s', %s, %s)," % (fieldname, self.type_name(f.typ), f.bits)
            print >> self.stream, "]"
            # generate assert statements for size and alignment
            if body.struct.size and body.struct.name not in dont_assert_size:
                size = body.struct.size // 8
                print >> self.stream, "assert sizeof(%s) == %s, sizeof(%s)" % \
                      (body.struct.name, size, body.struct.name)
                align = body.struct.align // 8
                print >> self.stream, "assert alignment(%s) == %s, alignment(%s)" % \
                      (body.struct.name, align, body.struct.name)

        if methods:
            # Ha! Autodetect ctypes.com or comtypes ;)
            if "COMMETHOD" in self.known_symbols:
                self.need_COMMETHOD()
            else:
                self.need_STDMETHOD()
            # method definitions normally span several lines.
            # Before we generate them, we need to 'import' everything they need.
            # So, call type_name for each field once,
            for m in methods:
                self.type_name(m.returns)
                for a in m.iterArgTypes():
                    self.type_name(a)
            if "COMMETHOD" in self.known_symbols:
                print >> self.stream, "%s._methods_ = [" % body.struct.name
            else:
                # ctypes.com needs baseclass methods listed as well
                if body.struct.bases:
                    basename = body.struct.bases[0].name
                    print >> self.stream, "%s._methods_ = %s._methods + [" % \
                          (body.struct.name, basename)
                else:
                    print >> self.stream, "%s._methods_ = [" % body.struct.name
            if body.struct.location:
                print >> self.stream, "# %s %s" % body.struct.location

            if "COMMETHOD" in self.known_symbols:
                for m in methods:
                    if m.location:
                        print >> self.stream, "    # %s %s" % m.location
                    print >> self.stream, "    COMMETHOD([], %s, '%s'," % (
                        self.type_name(m.returns),
                        m.name)
                    for a in m.iterArgTypes():
                        print >> self.stream, \
                              "               ( [], %s, )," % self.type_name(a)
                    print >> self.stream, "             ),"
            else:
                for m in methods:
                    args = [self.type_name(a) for a in m.iterArgTypes()]
                    print >> self.stream, "    STDMETHOD(%s, '%s', [%s])," % (
                        self.type_name(m.returns),
                        m.name,
                        ", ".join(args))
            print >> self.stream, "]"

    def find_dllname(self, func):
        if hasattr(func, "dllname"):
            return func.dllname
        name = func.name
        for dll in self.searched_dlls:
            try:
                getattr(dll, name)
            except AttributeError:
                pass
            else:
                return dll._name
##        if self.verbose:
        # warnings.warn, maybe?
##        print >> sys.stderr, "function %s not found in any dll" % name
        return None

    _loadedlibs = None
    def get_sharedlib(self, dllname):
        if self._loadedlibs is None:
            self._loadedlibs = {}
        try:
            return self._loadedlibs[dllname]
        except KeyError:
            pass
        import os
        basename = os.path.basename(dllname)
        name, ext = os.path.splitext(basename)
        self._loadedlibs[dllname] = name
        # This should be handled in another way!
##        print >> self.stream, "%s = CDLL(%r)" % (name, dllname)
        return name

    _STRING_defined = False
    def need_STRING(self):
        if self._STRING_defined:
            return
        print >> self.imports, "STRING = c_char_p"
        self._STRING_defined = True

    _WSTRING_defined = False
    def need_WSTRING(self):
        if self._WSTRING_defined:
            return
        print >> self.imports, "WSTRING = c_wchar_p"
        self._WSTRING_defined = True

    _STDMETHOD_defined = False
    def need_STDMETHOD(self):
        if self._STDMETHOD_defined:
            return
        print >> self.imports, "from ctypes.com import STDMETHOD"
        self._STDMETHOD_defined = True

    _COMMETHOD_defined = False
    def need_COMMETHOD(self):
        if self._COMMETHOD_defined:
            return
        print >> self.imports, "from comtypes import COMMETHOD"
        self._COMMETHOD_defined = True

    _GUID_defined = False
    def need_GUID(self):
        if self._GUID_defined:
            return
        self._GUID_defined = True
        modname = self.known_symbols.get("GUID")
        if modname:
            print >> self.imports, "from %s import GUID" % modname

    _functiontypes = 0
    _notfound_functiontypes = 0
    def Function(self, func):
        dllname = self.find_dllname(func)
        if dllname:
            self.generate(func.returns)
            self.generate_all(func.iterArgTypes())
            args = [self.type_name(a) for a in func.iterArgTypes()]
            if "__stdcall__" in func.attributes:
                cc = "stdcall"
            else:
                cc = "cdecl"
            libname = self.get_sharedlib(dllname)
            print >> self.stream
            if self.use_decorators:
                print >> self.stream, "@ %s(%s, '%s', [%s])" % \
                      (cc, self.type_name(func.returns), libname, ", ".join(args))
            argnames = [a or "p%d" % (i+1) for i, a in enumerate(func.iterArgNames())]
            # function definition
            print >> self.stream, "def %s(%s):" % (func.name, ", ".join(argnames))
            if func.location:
                print >> self.stream, "    # %s %s" % func.location
            print >> self.stream, "    return %s._api_(%s)" % (func.name, ", ".join(argnames))
            if not self.use_decorators:
                print >> self.stream, "%s = %s(%s, '%s', [%s]) (%s)" % \
                      (func.name, cc, self.type_name(func.returns), libname, ", ".join(args), func.name)
            print >> self.stream
            self.names.add(func.name)
            self._functiontypes += 1
        else:
            self._notfound_functiontypes += 1

    def FundamentalType(self, item):
        pass # we should check if this is known somewhere
##        name = ctypes_names[item.name]
##        if name !=  "None":
##            print >> self.stream, "from ctypes import %s" % name
##        self.done.add(item)

    ########

    def generate(self, item):
        if item in self.done:
            return
        if isinstance(item, typedesc.StructureHead):
            name = getattr(item.struct, "name", None)
        else:
            name = getattr(item, "name", None)
        if name in self.known_symbols:
            mod = self.known_symbols[name]
            print >> self.imports, "from %s import %s" % (mod, name)
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

    def cmpitems(a, b):
	a = getattr(a, "location", None)
	b = getattr(b, "location", None)
	if a is None: return -1
	if b is None: return 1
	return cmp(a[0],b[0]) or cmp(int(a[1]),int(b[1]))
    cmpitems = staticmethod(cmpitems)

    def generate_code(self, items):
        print >> self.imports, "from ctypes import *"
        items = set(items)
        loops = 0
        while items:
            loops += 1
            self.more = set()
            self.generate_all(sorted(items, self.cmpitems))

            items |= self.more
            items -= self.done

        self.output.write(self.imports.getvalue())
        self.output.write("\n\n")
        self.output.write(self.stream.getvalue())

        import textwrap
        text = "__all__ = [%s]" % ", ".join([repr(str(n)) for n in self.names])

        for line in textwrap.wrap(text,
                                  subsequent_indent="           "):
            print >> self.output, line
        return loops

    def print_stats(self, stream):
        total = self._structures + self._functiontypes + self._enumtypes + self._typedefs +\
                self._pointertypes + self._arraytypes
        print >> stream, "###########################"
        print >> stream, "# Symbols defined:"
        print >> stream, "#"
        print >> stream, "# Variables:          %5d" % self._variables
        print >> stream, "# Struct/Unions:      %5d" % self._structures
        print >> stream, "# Functions:          %5d" % self._functiontypes
        print >> stream, "# Enums:              %5d" % self._enumtypes
        print >> stream, "# Enum values:        %5d" % self._enumvalues
        print >> stream, "# Typedefs:           %5d" % self._typedefs
        print >> stream, "# Pointertypes:       %5d" % self._pointertypes
        print >> stream, "# Arraytypes:         %5d" % self._arraytypes
        print >> stream, "# unknown functions:  %5d" % self._notfound_functiontypes
        print >> stream, "#"
        print >> stream, "# Total symbols: %5d" % total
        print >> stream, "###########################"

################################################################

def generate_code(xmlfile,
                  outfile,
                  expressions=None,
                  symbols=None,
                  verbose=False,
                  use_decorators=False,
                  known_symbols=None,
                  searched_dlls=None,
                  types=None):
    # expressions is a sequence of compiled regular expressions,
    # symbols is a sequence of names
    from gccxmlparser import parse
    items = parse(xmlfile)

    # filter symbols to generate
    todo = []

    if types:
        items = [i for i in items if isinstance(i, types)]
    
    if symbols:
        syms = set(symbols)
        for i in items:
            if i.name in syms:
                todo.append(i)
                syms.remove(i.name)

        if syms:
            print "symbols not found", list(syms)

    if expressions:
        for i in items:
            for s in expressions:
                if i.name is None:
                    continue
                match = s.match(i.name)
                # we only want complete matches
                if match and match.group() == i.name:
                    todo.append(i)
                    break
    if symbols or expressions:
        items = todo

    ################
    gen = Generator(outfile,
                    use_decorators=use_decorators,
                    known_symbols=known_symbols,
                    searched_dlls=searched_dlls)

    loops = gen.generate_code(items)
    if verbose:
        gen.print_stats(sys.stderr)
        print >> sys.stderr, "needed %d loop(s)" % loops

