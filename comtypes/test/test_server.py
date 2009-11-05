import os
import atexit
import comtypes.typeinfo, comtypes.client

class TypeLib(object):
    """This class collects IDL code fragments and eventually writes
    them into a .IDL file.  The compile() method compiles the IDL file
    into a typelibrary and registers it.  A function is also
    registered with atexit that will unregister the typelib at program
    exit.
    """
    def __init__(self, lib):
        self.lib = lib
        self.interfaces = []

    def interface(self, header):
        itf = Interface(header)
        self.interfaces.append(itf)
        return itf

    def __str__(self):
        header = '''import "oaidl.idl";
                    import "ocidl.idl";
                    %s {''' % self.lib
        body = "\n".join([str(itf) for itf in self.interfaces])
        footer = "}"
        return "\n".join((header, body, footer))

    def compile(self):
        """Compile and register the typelib"""
        code = str(self)
        curdir = os.path.dirname(__file__)
        idl_path = os.path.join(curdir, "mylib.idl")
        tlb_path = os.path.join(curdir, "mylib.tlb")
        if not os.path.isfile(idl_path) or open(idl_path, "r").read() != code:
            open(idl_path, "w").write(code)
            os.system(r'call "%%VS71COMNTOOLS%%vsvars32.bat" && '
                      r'midl /nologo %s /tlb %s' % (idl_path, tlb_path))
        # Register the typelib...
        tlib = comtypes.typeinfo.LoadTypeLib(tlb_path)
        # create the wrapper module...
        comtypes.client.GetModule(tlb_path)
        # Unregister the typelib at interpreter exit...
        attr = tlib.GetLibAttr()
        guid, major, minor = attr.guid, attr.wMajorVerNum, attr.wMinorVerNum
        atexit.register(lambda:
                        comtypes.typeinfo.UnRegisterTypeLib(guid, major, minor))
        return tlb_path
    
class Interface(object):
    def __init__(self, header):
        self.header = header
        self.code = ""

    def add(self, text):
        self.code += text + "\n"
        return self

    def __str__(self):
        return self.header + " {\n" + self.code + "}\n"

################################################################
import comtypes
from comtypes.client import wrap

tlb = TypeLib("[uuid(f4f74946-4546-44bd-a073-9ea6f9fe78cb)] library TestLib")

itf = tlb.interface("""[object,
                        // oleautomation,
                        uuid(ed978f5f-cc45-4fcc-a7a6-751ffa8dfedd)]
                       interface IMyInterface : IDispatch""")

# The purpose of the MyServer class is to locate three separate code
# section snippets closely together:
#
# 1. The IDL method definition for a COM interface method
# 2. The Python implementation of the COM method
# 3. The unittest(s) for the COM method.
#
class MyServer(comtypes.CoClass):
    _reg_typelib_ = ('{f4f74946-4546-44bd-a073-9ea6f9fe78cb}', 0, 0)

    ################
    # definition
    itf.add("""[propget] HRESULT Name([out, retval] BSTR *pname);
               [propput] HRESULT Name([in] BSTR name);""")
    # implementation
    Name = "foo"
    # test
    def test_Name(self):
        p = wrap(self.create())
        self.assertEqual((p.Name, p.name, p.nAME), ("foo",) * 3)
        p.NAME = "spam"
        self.assertEqual((p.Name, p.name, p.nAME), ("spam",) * 3)

    ################
    # definition
    itf.add("HRESULT MixedInOut([in] int a, [out] int *b, [in] int c, [out] int *d);")
    # implementation
    def MixedInOut(self, a, c):
        return a+1, c+1
    #test
    def test_MixedInOut(self):
        p = wrap(self.create())
        self.assertEqual(p.MixedInOut(1, 2), (2, 3))

    ################
    # definition
    itf.add("HRESULT MultiInOutArgs([in, out] int *pa, [in, out] int *pb);")
    # implementation
    def MultiInOutArgs(self, pa, pb):
        return pa[0] * 3, pb[0] * 4
    # test
    def test_MultiInOutArgs(self):
        p = wrap(self.create())
        self.assertEqual(p.MultiInOutArgs(1, 2), (3, 8))

    ################
    # definition
    itf.add("HRESULT MultiInOutArgs2([in, out] int *pa, [out] int *pb);")
##    # implementation
##    def MultiInOutArgs2(self, pa):
##        return pa[0] * 3, pa[0] * 4
##    # test
##    def test_MultiInOutArgs2(self):
##        p = wrap(self.create())
##        self.assertEqual(p.MultiInOutArgs2(42), (126, 168))

    ################
    # definition
    itf.add("HRESULT MultiInOutArgs3([out] int *pa, [out] int *pb);")
    # implementation
    def MultiInOutArgs3(self):
        return 42, 43
    # test
    def test_MultiInOutArgs3(self):
        p = wrap(self.create())
        self.assertEqual(p.MultiInOutArgs3(), (42, 43))

    ################
    # definition
    itf.add("HRESULT MultiInOutArgs4([out] int *pa, [in, out] int *pb);")
    # implementation
    def MultiInOutArgs4(self, pb):
        return pb[0] + 3, pb[0] + 4
    # test
    def test_MultiInOutArgs4(self):
        p = wrap(self.create())
        res = p.MultiInOutArgs4(pb=32)
##        print "MultiInOutArgs4", res

    itf.add("""HRESULT GetStackTrace([in] ULONG FrameOffset,
                                     [in, out] INT *Frames,
                                     [in] ULONG FramesSize,
                                     [out, optional] ULONG *FramesFilled);""")
    def GetStackTrace(self, this, *args):
##        print "GetStackTrace", args
        return 0
    def test_GetStackTrace(self):
        p = wrap(self.create())
        from ctypes import c_int, POINTER, pointer
        frames = (c_int * 5)()
        res = p.GetStackTrace(42, frames, 5)
##        print "RES_1", res

        frames = pointer(c_int(5))
        res = p.GetStackTrace(42, frames, 0)
##        print "RES_2", res

path = tlb.compile()
from comtypes.gen import TestLib

MyServer._com_interfaces_ = [TestLib.IMyInterface]

################################################################

import unittest

class Test(unittest.TestCase, MyServer):
    def create(self):
        obj = MyServer()
        return obj.QueryInterface(comtypes.IUnknown)


if __name__ == "__main__":
    unittest.main()
