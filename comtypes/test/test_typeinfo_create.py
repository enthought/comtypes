import tempfile
import unittest
from ctypes import HRESULT, POINTER, Structure
from ctypes.wintypes import DWORD, INT, ULONG
from pathlib import Path

from comtypes import BSTR, COMMETHOD, GUID, typeinfo
from comtypes.automation import LCID, VARIANT, VARIANTARG
from comtypes.typeinfo import CreateTypeLib, ITypeLib, LoadTypeLibEx


class tagCUSTDATAITEM(Structure):
    # https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-custdataitem
    _fields_ = [
        ("guid", GUID),
        ("varValue", VARIANTARG),
    ]


CUSTDATAITEM = tagCUSTDATAITEM


class tagCUSTDATA(Structure):
    # https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-custdata
    _fields_ = [
        ("cElems", DWORD),
        ("pCustData", POINTER(CUSTDATAITEM)),
    ]


CUSTDATA = tagCUSTDATA


class ITypeLib2(ITypeLib):
    _iid_ = GUID("{00020411-0000-0000-C000-000000000046}")

    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "GetCustData",
            (["in"], POINTER(GUID), "guid"),
            (["out"], POINTER(VARIANT), "pVarVal"),
        ),
        COMMETHOD(
            [], HRESULT, "GetAllCustData", (["out"], POINTER(CUSTDATA), "pCustData")
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetDocumentation2",
            (["in"], INT, "index"),
            (["in"], LCID, "lcid"),
            (["out"], POINTER(BSTR), "pbstrHelpString"),
            (["out"], POINTER(DWORD), "pdwHelpStringContext"),
            (["out"], POINTER(BSTR), "pbstrHelpStringDll"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "GetLibStatistics",
            (["out"], POINTER(ULONG), "pcUniqueNames"),
            (["out"], POINTER(ULONG), "pcchUniqueNames"),
        ),
    ]


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

    def test_GetDocumentation2(self):
        ctlib = CreateTypeLib(str(self.typelib_path))
        helpstring = "This is a test type library helpstring."
        helpctx = 123
        helpdll = "myhelp.dll"
        ctlib.SetDocString(helpstring)
        ctlib.SetHelpStringContext(helpctx)
        ctlib.SetHelpStringDll(helpdll)
        ctlib.SaveAllChanges()
        tlib2 = ctlib.QueryInterface(ITypeLib2)
        self.assertEqual(tlib2.GetDocumentation2(-1, 0), (helpstring, helpctx, helpdll))

    def test_GetCustData(self):
        ctlib = CreateTypeLib(str(self.typelib_path))
        guid = GUID.create_new()
        val = "Custom Library Data"
        ctlib.SetCustData(guid, val)
        ctlib.SaveAllChanges()
        tlib2 = ctlib.QueryInterface(ITypeLib2)
        self.assertEqual(tlib2.GetCustData(guid), val)


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
