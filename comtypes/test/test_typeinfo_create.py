import tempfile
import unittest
from pathlib import Path

from comtypes import GUID, typeinfo
from comtypes.typeinfo import CreateTypeLib, LoadTypeLibEx


class Test_CreateTypeLib(unittest.TestCase):
    def setUp(self):
        td = tempfile.TemporaryDirectory()
        self.addCleanup(td.cleanup)
        self.tmpdir = Path(td.name)
        self.typelib_path = self.tmpdir / "test.tlb"

    def test_Documentation(self):
        ctlib = CreateTypeLib(str(self.typelib_path))
        libname = "MyTestTypeLib"
        docstring = "This is a test type library docstring."
        helpctx = 123
        helpfile = "myhelp.chm"
        ctlib.SetName(libname)
        ctlib.SetDocString(docstring)
        ctlib.SetHelpContext(helpctx)
        ctlib.SetHelpFileName(helpfile)
        ctlib.SaveAllChanges()
        # Verify by loading the created typelib
        tlib = LoadTypeLibEx(str(self.typelib_path))
        doc = tlib.GetDocumentation(-1)
        self.assertEqual(
            doc[:-1] + (Path(doc[-1].strip("\0")).name,),  # type: ignore
            (libname, docstring, helpctx, helpfile),
        )

    def test_LibAttr(self):
        ctlib = CreateTypeLib(str(self.typelib_path))
        libid = GUID.create_new()
        lcid = 1033  # English (United States)
        LIBFLAG_FRESTRICTED = 0x1
        major_version = 1
        minor_version = 2
        ctlib.SetGuid(libid)
        ctlib.SetLcid(lcid)
        ctlib.SetVersion(major_version, minor_version)
        ctlib.SetLibFlags(LIBFLAG_FRESTRICTED)
        ctlib.SaveAllChanges()
        # Verify by loading the created typelib
        tlib = LoadTypeLibEx(str(self.typelib_path))
        la = tlib.GetLibAttr()
        self.assertEqual(la.guid, libid)
        self.assertEqual(la.lcid, lcid)
        self.assertTrue(la.wLibFlags & LIBFLAG_FRESTRICTED)
        self.assertEqual(la.wMajorVerNum, major_version)
        self.assertEqual(la.wMinorVerNum, minor_version)

    def test_CreateTypeInfo(self):
        ctlib = CreateTypeLib(str(self.typelib_path))
        # Create a type info
        self.assertIsInstance(
            ctlib.CreateTypeInfo("IMyInterface", typeinfo.TKIND_INTERFACE),
            typeinfo.ICreateTypeInfo,
        )


class Test_ICreateTypeInfo(unittest.TestCase):
    def setUp(self):
        td = tempfile.TemporaryDirectory()
        self.addCleanup(td.cleanup)
        self.tmpdir = Path(td.name)
        self.typelib_path = self.tmpdir / "test.tlb"
        self.ctlib = CreateTypeLib(str(self.typelib_path))

    def test_Documentaion(self):
        name = "IMyInterface"
        docstring = "My test interface"
        helpctx = 123
        ctinfo = self.ctlib.CreateTypeInfo(name, typeinfo.TKIND_INTERFACE)
        ctinfo.SetDocString(docstring)
        ctinfo.SetHelpContext(helpctx)
        # `Layout` must be called before `SaveAllChanges`.
        ctinfo.LayOut()
        self.ctlib.SaveAllChanges()
        # Load the typelib and verify the type info
        tlib = LoadTypeLibEx(str(self.typelib_path))
        _, tinfo = tlib.FindName(name)  # type: ignore
        doc = tinfo.GetDocumentation(-1)
        self.assertEqual(doc, (name, docstring, helpctx, None))
