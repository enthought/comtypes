import tempfile
import unittest
from ctypes import HRESULT, POINTER, Structure, pointer
from ctypes.wintypes import DWORD, INT, ULONG
from pathlib import Path

from comtypes import BSTR, COMMETHOD, GUID, IUnknown, automation, typeinfo
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

    def test_Interface_TYPEATTR(self):
        # Create the base type info
        base_name = "IUnknown"
        base_iid = IUnknown._iid_
        base_ctinfo = self.ctlib.CreateTypeInfo(base_name, typeinfo.TKIND_INTERFACE)
        base_ctinfo.SetGuid(base_iid)
        base_ctinfo.LayOut()
        # Create the derived type info
        derived_name = "IMyDerived"
        derived_iid = GUID.create_new()
        typeflags = typeinfo.TYPEFLAG_FHIDDEN
        major_version = 2
        minor_version = 1
        alignment = 8
        derived_ctinfo = self.ctlib.CreateTypeInfo(
            derived_name, typeinfo.TKIND_INTERFACE
        )
        derived_ctinfo.SetGuid(derived_iid)
        derived_ctinfo.SetTypeFlags(typeflags)
        derived_ctinfo.SetVersion(major_version, minor_version)
        derived_ctinfo.SetAlignment(alignment)
        derived_ctinfo.AddImplType(0, derived_ctinfo.AddRefTypeInfo(base_ctinfo))
        derived_ctinfo.LayOut()
        self.ctlib.SaveAllChanges()
        # Load the typelib and verify the type info's GUID
        tlib = LoadTypeLibEx(str(self.typelib_path))
        # Get the base type info
        _, base_tinfo = tlib.FindName(base_name)  # type: ignore
        base_ta = base_tinfo.GetTypeAttr()
        self.assertEqual(base_ta.cImplTypes, 0)  # has NO referenced type
        self.assertEqual(base_ta.guid, base_iid)
        self.assertEqual(base_ta.wTypeFlags, 0)
        self.assertEqual(base_ta.cbAlignment, alignment)
        # Get the derived type info
        _, derived_tinfo = tlib.FindName(derived_name)  # type: ignore
        derived_ta = derived_tinfo.GetTypeAttr()
        self.assertEqual(derived_ta.cImplTypes, 1)  # has a referenced type
        self.assertEqual(derived_ta.wTypeFlags, typeflags)
        self.assertEqual(derived_ta.wMajorVerNum, major_version)
        self.assertEqual(derived_ta.wMinorVerNum, minor_version)
        self.assertEqual(derived_ta.guid, derived_iid)
        # Get the referenced type info
        ref_tinfo = derived_tinfo.GetRefTypeInfo(derived_tinfo.GetRefTypeOfImplType(0))
        ref_ta = ref_tinfo.GetTypeAttr()
        self.assertEqual(ref_ta.guid, base_iid)

    def test_Alias_TYPEATTR(self):
        alias_name = "MyAlias"
        ctinfo_alias = self.ctlib.CreateTypeInfo(alias_name, typeinfo.TKIND_ALIAS)
        ctinfo_alias.SetTypeDescAlias(typeinfo.TYPEDESC(vt=automation.VT_INT))
        ctinfo_alias.LayOut()
        self.ctlib.SaveAllChanges()
        # Load the typelib and verify the alias
        tlib = LoadTypeLibEx(str(self.typelib_path))
        _, tinfo_alias = tlib.FindName(alias_name)
        typeattr_alias = tinfo_alias.GetTypeAttr()
        self.assertEqual(typeattr_alias.typekind, typeinfo.TKIND_ALIAS)
        self.assertEqual(typeattr_alias.tdescAlias.vt, automation.VT_INT)

    def test_FUNCDESC(self):
        func_name = "MyFunction"
        itf_name = "IMyInterface"
        itf_ctinfo = self.ctlib.CreateTypeInfo(itf_name, typeinfo.TKIND_INTERFACE)
        func_memid = 42  # Arbitrary member ID
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
        itf_ctinfo.LayOut()
        self.ctlib.SaveAllChanges()
        # Load the typelib and verify the function
        tlib = LoadTypeLibEx(str(self.typelib_path))
        _, tinfo = tlib.FindName(itf_name)  # type: ignore
        ta = tinfo.GetTypeAttr()
        self.assertEqual(ta.cFuncs, 1)  # Should have one function
        fd = tinfo.GetFuncDesc(0)
        self.assertEqual(fd.memid, func_memid)

    def test_VARDESC(self):
        mod_name = "MyModule"
        var_name = "MyDocVar"
        var_value = "MyValue"
        var_memid = 102  # Arbitrary member ID
        mod_ctinfo = self.ctlib.CreateTypeInfo(mod_name, typeinfo.TKIND_MODULE)
        vardesc = typeinfo.VARDESC(
            memid=var_memid,
            varkind=typeinfo.VAR_CONST,
            elemdescVar=typeinfo.ELEMDESC(
                tdesc=typeinfo.TYPEDESC(vt=automation.VT_BSTR)
            ),
            wVarFlags=typeinfo.VARFLAG_FDEFAULTBIND,
        )
        vardesc._.lpvarValue = pointer(automation.VARIANT(var_value))
        mod_ctinfo.AddVarDesc(0, vardesc)
        mod_ctinfo.SetVarName(0, var_name)
        mod_ctinfo.LayOut()
        self.ctlib.SaveAllChanges()
        # Load the typelib
        tlib = LoadTypeLibEx(str(self.typelib_path))
        _, tinfo = tlib.FindName(var_name)  # type: ignore
        ta = tinfo.GetTypeAttr()
        self.assertEqual(ta.cVars, 1)  # Should have one variable
        vd = tinfo.GetVarDesc(0)
        self.assertEqual(vd.memid, var_memid)
        self.assertEqual(vd.varkind, typeinfo.VAR_CONST)
        self.assertEqual(vd.wVarFlags, typeinfo.VARFLAG_FDEFAULTBIND)
        self.assertEqual(vd.elemdescVar.tdesc.vt, automation.VT_BSTR)
        self.assertEqual(vd._.lpvarValue[0].value, var_value)
