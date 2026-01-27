import tempfile
import unittest
from ctypes import HRESULT, POINTER, Structure, pointer
from ctypes.wintypes import DWORD, INT, ULONG
from pathlib import Path

from comtypes import BSTR, COMMETHOD, GUID, automation, typeinfo
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

    def test_Type_Documentation(self):
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

    def test_Func_Documentation(self):
        func_name = "MyFunction"
        func_docstring = "This is my function's docstring."
        func_helpctx = 999
        func_memid = 42  # Arbitrary member ID
        itf_name = "IMyInterface"
        # Create a new typeinfo for the function test to avoid interference
        itf_ctinfo = self.ctlib.CreateTypeInfo(itf_name, typeinfo.TKIND_INTERFACE)
        itf_ctinfo.AddFuncDesc(
            0,
            typeinfo.FUNCDESC(
                memid=func_memid,
                funckind=typeinfo.FUNC_PUREVIRTUAL,
                invkind=typeinfo.INVOKE_FUNC,
                callconv=typeinfo.CC_STDCALL,
                cParams=0,
                elemdescFunc=typeinfo.ELEMDESC(
                    tdesc=typeinfo.TYPEDESC(vt=automation.VT_HRESULT)
                ),
            ),
        )
        itf_ctinfo.SetFuncAndParamNames(0, func_name)
        itf_ctinfo.SetFuncDocString(0, func_docstring)
        itf_ctinfo.SetFuncHelpContext(0, func_helpctx)
        itf_ctinfo.LayOut()
        self.ctlib.SaveAllChanges()
        # Load and verify function documentation
        tlib_func = LoadTypeLibEx(str(self.typelib_path))
        _, tinfo_func = tlib_func.FindName(itf_name)  # type: ignore
        doc = tinfo_func.GetDocumentation(func_memid)
        self.assertEqual(doc, (func_name, func_docstring, func_helpctx, None))

    def test_Var_Documentation(self):
        mod_name = "MyModule"
        var_name = "MyDocVar"
        var_value = "MyValue"
        var_docstring = "This is my variable's docstring."
        var_helpctx = 888
        var_memid = 102  # Arbitrary member ID
        mod_ctinfo = self.ctlib.CreateTypeInfo(mod_name, typeinfo.TKIND_MODULE)
        vardesc = typeinfo.VARDESC(
            memid=var_memid,
            varkind=typeinfo.VAR_CONST,
            elemdescVar=typeinfo.ELEMDESC(
                tdesc=typeinfo.TYPEDESC(vt=automation.VT_BSTR)
            ),
        )
        vardesc._.lpvarValue = pointer(automation.VARIANT(var_value))
        mod_ctinfo.AddVarDesc(0, vardesc)
        mod_ctinfo.SetVarName(0, var_name)
        mod_ctinfo.SetVarDocString(0, var_docstring)
        mod_ctinfo.SetVarHelpContext(0, var_helpctx)
        mod_ctinfo.LayOut()
        self.ctlib.SaveAllChanges()
        tlib = LoadTypeLibEx(str(self.typelib_path))
        _, tinfo = tlib.FindName(var_name)  # type: ignore
        doc = tinfo.GetDocumentation(var_memid)
        self.assertEqual(doc, (var_name, var_docstring, var_helpctx, None))
