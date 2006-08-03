# Extended code generator able to generate code for everything
# contained in COM type libraries.

from comtypes.tools import typedesc
from comtypes.tools import codegenerator_base

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

class Generator(codegenerator_base.Generator):

    def __init__(self, ofi, make_module=None, name_module=None, *args, **kw):
        self._name_module = name_module
        self._make_module = make_module
        self._externals = {}
        super(Generator, self).__init__(ofi, *args, **kw)

    def generate_code(self, items, filename=None):
        if filename is None:
            filename = "<unable to determine filename>"
        print >> self.imports, "# typelib %s" % filename
        print >> self.imports, "_lcid = 0 # change this if required"
        return super(Generator, self).generate_code(items)

    def type_name(self, t, generate=True):
        # Return a string, containing an expression which can be used to
        # refer to the type. Assumes the * namespace is available.
        if isinstance(t, typedesc.SAFEARRAYType):
            return "_midlSAFEARRAY(%s)" % self.type_name(t.typ)
##        if isinstance(t, typedesc.CoClass):
##            return "%s._com_interfaces_[0]" % t.name
        return super(Generator, self).type_name(t, generate)

    _midlSAFEARRAY_defined = False
    def need_midlSAFEARRAY(self):
        if self._midlSAFEARRAY_defined:
            return
        print >> self.imports, "from comtypes.automation import _midlSAFEARRAY"
        self._midlSAFEARRAY_defined = True

    _CoClass_defined = False
    def need_CoClass(self):
        if self._CoClass_defined:
            return
        print >> self.imports, "from comtypes import CoClass"
        self._CoClass_defined = True

    _dispid_defined = False
    def need_dispid(self):
        if self._dispid_defined:
            return
        print >> self.imports, "from comtypes import dispid"
        self._dispid_defined = True

    def need_COMMETHOD(self):
        if self._COMMETHOD_defined:
            return
        print >> self.imports, "from comtypes import helpstring"
        super(Generator, self).need_COMMETHOD()

    _DISPMETHOD_defined = False
    def need_DISPMETHOD(self):
        if self._DISPMETHOD_defined:
            return
        print >> self.imports, "from comtypes import DISPMETHOD, DISPPROPERTY, helpstring"
        self._DISPMETHOD_defined = True

    ################################################################
    # top-level typedesc generators
    #
    def External(self, ext):
        # ext.docs - docstring of typelib
        # ext.symbol_name - symbol to generate
        # ext.tlib - the ITypeLib pointer to the typelibrary containing the symbols definition
        #
        # ext.name filled in here

        # XXX This code should be refactored to NOT require that
        # _make_module and _name_module must be set on the Generator
        # instance.  Should we pass a ModuleGenerator instance instead,
        # with make_module and name_module methods?

        # Or, are there other ways?
        
        libdesc = str(ext.tlib.GetLibAttr()) # str(TLIBATTR) is unique for a given typelib
        if libdesc in self._externals: # typelib wrapper already created
            modname = self._externals[libdesc]
            # we must fill in ext.name, it is used by self.type_name()
            ext.name = "%s.%s" % (modname, ext.symbol_name)
            return

        if self._name_module is None:
            ext.name = ext.symbol_name
            print "# Cannot create external symbol", ext.symbol_name
            return

        modname = self._name_module(ext.tlib)
        ext.name = "%s.%s" % (modname, ext.symbol_name)
        self._externals[libdesc] = modname
        print >> self.imports, "import", modname
        # it may be already generated in another instance of
        # Generator, but _make_module will catch this.
        if self._make_module is not None:
            self._make_module(ext.tlib)

    def Constant(self, tp):
        print >> self.stream, \
              "%s = %r # Constant %s" % (tp.name,
                                         tp.value,
                                         self.type_name(tp.typ, False))
        self.names.add(tp.name)

    def SAFEARRAYType(self, sa):
        self.generate(sa.typ)
        self.need_midlSAFEARRAY()

    def PointerType(self, tp):
        if type(tp.typ) is typedesc.ComInterface:
            # this defines the class
            self.generate(tp.typ.get_head())
            # this defines the _methods_
            self.more.add(tp.typ)
        elif type(tp.typ) is typedesc.PointerType:
            self.generate(tp.typ)
        else:
            super(Generator, self).PointerType(tp)

    def CoClass(self, coclass):
        self.need_GUID()
        self.need_CoClass()
        print >> self.stream, "class %s(CoClass):" % coclass.name
        doc = getattr(coclass, "doc", None)
        if doc:
            print >> self.stream, "    %r" % doc
        print >> self.stream, "    _reg_clsid_ = GUID(%r)" % coclass.clsid
        print >> self.stream, "    _idlflags_ = %s" % coclass.idlflags
##X        print >> self.stream, "POINTER(%s).__ctypes_from_outparam__ = wrap" % coclass.name

        libid = coclass.tlibattr.guid
        wMajor, wMinor = coclass.tlibattr.wMajorVerNum, coclass.tlibattr.wMinorVerNum
        print >> self.stream, "    _reg_typelib_ = (%r, %s, %s)" % (str(libid), wMajor, wMinor)

        for itf, idlflags in coclass.interfaces:
            self.generate(itf.get_head())
        implemented = [i[0].name for i in coclass.interfaces
                       if i[1] & 2 == 0]
        sources = [i[0].name for i in coclass.interfaces
                       if i[1] & 2 == 2]
        if implemented:
            print >> self.stream, "%s._com_interfaces_ = [%s]" % (coclass.name, ", ".join(implemented))
        if sources:
            print >> self.stream, "%s._outgoing_interfaces_ = [%s]" % (coclass.name, ", ".join(sources))
        print >> self.stream
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
        print >> self.stream, "class %s(%s):" % (head.itf.name, basename)
        print >> self.stream, "    _case_insensitive_ = True"
        doc = getattr(head.itf, "doc", None)
        if doc:
            print >> self.stream, "    %r" % doc
        print >> self.stream, "    _iid_ = GUID(%r)" % head.itf.iid
        print >> self.stream, "    _idlflags_ = %s" % head.itf.idlflags

        if self._is_enuminterface(head.itf):
            print >> self.stream, "    def __iter__(self):"
            print >> self.stream, "        return self"
            print >> self.stream

            # Well, not sure if they are really broken, but sometimes
            # the last parameter to Next is marked [in, out],
            # sometimes it is only [out].
            NextIsBroken = False
            for mth in head.itf.members:
                if mth.name == "Next":
                    NextIsBroken = 'in' in mth.arguments[-1][2]
                    break

            print >> self.stream, "    def next(self):"
            if NextIsBroken:
                print >> self.stream, "        item, fetched = self.Next(1, 0)"
            else:
                print >> self.stream, "        item, fetched = self.Next(1)"
            print >> self.stream, "        if fetched:"
            print >> self.stream, "            return item"
            print >> self.stream, "        raise StopIteration"
            print >> self.stream

            print >> self.stream, "    def __getitem__(self, index):"
            print >> self.stream, "        self.Reset()"
            print >> self.stream, "        self.Skip(index)"
            if NextIsBroken:
                print >> self.stream, "        item, fetched = self.Next(1, 0)"
            else:
                print >> self.stream, "        item, fetched = self.Next(1)"
            print >> self.stream, "        if fetched:"
            print >> self.stream, "            return item"
            print >> self.stream, "        raise IndexError, index"
            print >> self.stream

    def ComInterfaceBody(self, body):
        # The base class must be fully generated, including the
        # _methods_ list.
        self.generate(body.itf.base)

        # make sure we can generate the body
        for m in body.itf.members:
            for a in m.arguments:
                self.generate(a[0])
            self.generate(m.returns)

        self.need_COMMETHOD()
        self.need_dispid()
        print >> self.stream, "%s._methods_ = [" % body.itf.name
        for m in body.itf.members:
            if isinstance(m, typedesc.ComMethod):
                self.make_ComMethod(m, "dual" in body.itf.idlflags)
            else:
                raise TypeError, "what's this?"
        print >> self.stream, "]"

    def DispInterface(self, itf):
        self.generate(itf.get_head())
        self.generate(itf.get_body())
        self.names.add(itf.name)

    def DispInterfaceHead(self, head):
        self.generate(head.itf.base)
        basename = self.type_name(head.itf.base)

        self.need_GUID()
        print >> self.stream, "class %s(%s):" % (head.itf.name, basename)
        print >> self.stream, "    _case_insensitive_ = True"
        doc = getattr(head.itf, "doc", None)
        if doc:
            print >> self.stream, "    %r" % doc
        print >> self.stream, "    _iid_ = GUID(%r)" % head.itf.iid
        print >> self.stream, "    _idlflags_ = %s" % head.itf.idlflags
        print >> self.stream, "    _methods_ = []"

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
                raise TypeError, m

        self.need_dispid()
        self.need_DISPMETHOD()
        print >> self.stream, "%s._disp_methods_ = [" % body.itf.name
        for m in body.itf.members:
            if isinstance(m, typedesc.DispMethod):
                self.make_DispMethod(m)
            elif isinstance(m, typedesc.DispProperty):
                self.make_DispProperty(m)
            else:
                raise TypeError, m
        print >> self.stream, "]"

    ################################################################
    # non-toplevel method generators
    #
    def make_ComMethod(self, m, isdual):
        # typ, name, idlflags, default
        if isdual:
            idlflags = [dispid(m.memid)] + m.idlflags
        else:
            # We don't include the dispid for non-dispatch COM interfaces
            idlflags = m.idlflags
        if __debug__ and m.doc:
            idlflags.insert(1, helpstring(m.doc))
        code = "    COMMETHOD(%r, %s, '%s'" % (
            idlflags,
            self.type_name(m.returns),
            m.name)

        if not m.arguments:
            print >> self.stream, "%s)," % code
        else:
            print >> self.stream, "%s," % code
            self.stream.write("              ")
            arglist = []
            for typ, name, idlflags, default in m.arguments:
                if 'lcid' in idlflags:# and 'in' in idlflags:
                    default = lcid
                if default is not None:
                    arglist.append("( %r, %s, '%s', %r )" % (
                        idlflags,
                        self.type_name(typ),
                        name,
                        default))
                else:
                    arglist.append("( %r, %s, '%s' )" % (
                        idlflags,
                        self.type_name(typ),
                        name))
            self.stream.write(",\n              ".join(arglist))
            print >> self.stream, "),"

    def make_DispMethod(self, m):
        idlflags = [dispid(m.dispid)] + m.idlflags
        if __debug__ and m.doc:
            idlflags.insert(1, helpstring(m.doc))
        # typ, name, idlflags, default
        code = "    DISPMETHOD(%r, %s, '%s'" % (
            idlflags,
            self.type_name(m.returns),
            m.name)

        if not m.arguments:
            print >> self.stream, "%s)," % code
        else:
            print >> self.stream, "%s," % code
            self.stream.write("               ")
            arglist = []
            for typ, name, idlflags, default in m.arguments:
                if default is not None:
                    arglist.append("( %r, %s, '%s', %r )" % (
                        idlflags,
                        self.type_name(typ),
                        name,
                        default))
                else:
                    arglist.append("( %r, %s, '%s' )" % (
                        idlflags,
                        self.type_name(typ),
                        name,
                        ))
            self.stream.write(",\n               ".join(arglist))
            print >> self.stream, "),"

    def make_DispProperty(self, prop):
        idlflags = [dispid(prop.dispid)] + prop.idlflags
        if __debug__ and prop.doc:
            idlflags.insert(1, helpstring(prop.doc))
        print >> self.stream, "    DISPPROPERTY(%r, %s, '%s')," % (
            idlflags,
            self.type_name(prop.typ),
            prop.name)

# shortcut for development
if __name__ == "__main__":
    import tlbparser
    tlbparser.main()
