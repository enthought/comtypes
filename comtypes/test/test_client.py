import contextlib
from ctypes import POINTER, byref
import os
import sys
import unittest as ut

import comtypes.client
from comtypes import COSERVERINFO, CLSCTX_INPROC_SERVER

# create the typelib wrapper and import it
comtypes.client.GetModule("scrrun.dll")
from comtypes.gen import Scripting


class Test_GetModule(ut.TestCase):
    def test_tlib_string(self):
        mod = comtypes.client.GetModule("scrrun.dll")
        self.assertIs(mod, Scripting)

    def test_abspath(self):
        mod = comtypes.client.GetModule(Scripting.typelib_path)
        self.assertIs(mod, Scripting)

    @ut.skipUnless(
        os.path.splitdrive(Scripting.typelib_path)[0]
        == os.path.splitdrive(__file__)[0],
        "This depends on typelib and test module are in same drive",
    )
    def test_relpath(self):
        relpath = os.path.relpath(Scripting.typelib_path, __file__)
        mod = comtypes.client.GetModule(relpath)
        self.assertIs(mod, Scripting)

    def test_libid_and_version_numbers(self):
        mod = comtypes.client.GetModule(Scripting.Library._reg_typelib_)
        self.assertIs(mod, Scripting)

    def test_one_length_sequence_containing_libid(self):
        libid, _, _ = Scripting.Library._reg_typelib_
        mod = comtypes.client.GetModule((libid,))
        self.assertIs(mod, Scripting)

    def test_obj_has_reg_libid_and_reg_version(self):
        typelib = Scripting.Library._reg_typelib_
        libid, version = typelib[0], typelib[1:]
        # HACK: Prefer to use Mock, but `unittest.mock` is not available in py27!
        info = type("info", (object,), dict(_reg_libid_=libid, _reg_version_=version))
        mod = comtypes.client.GetModule(info)
        self.assertIs(mod, Scripting)

    def test_clsid(self):
        clsid = comtypes.GUID.from_progid("MediaPlayer.MediaPlayer")
        mod = comtypes.client.GetModule(clsid)
        self.assertEqual(mod.MediaPlayer._reg_clsid_, clsid)

    def test_ptr_itypelib(self):
        from comtypes import typeinfo

        mod = comtypes.client.GetModule(typeinfo.LoadTypeLibEx("scrrun.dll"))
        self.assertIs(mod, Scripting)

    def test_imports_IEnumVARIANT_from_other_generated_modules(self):
        # NOTE: `codegenerator` generates code that contains unused imports,
        # but removing them are attracting wierd bugs in library-wrappers
        # which depend on externals.
        # NOTE: `mscorlib`, which imports `IEnumVARIANT` from `stdole`.
        comtypes.client.GetModule(("{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}",))

    def test_no_replacing_Patch_namespace(self):
        # NOTE: An object named `Patch` is defined in some dll.
        # Depending on how the namespace is defined in the static module,
        # `Patch` in generated modules will be replaced with
        # `comtypes.patcher.Patch`, and generating module will fail.
        # NOTE: `WindowsInstaller`, which has `Patch` definition in dll.
        comtypes.client.GetModule("msi.dll")

    def test_abstracted_wrapper_module_in_friendly_module(self):
        mod = comtypes.client.GetModule("scrrun.dll")
        self.assertTrue(hasattr(mod, "__wrapper_module__"))

    def test_raises_typerror_if_takes_unsupported(self):
        with self.assertRaises(TypeError):
            comtypes.client.GetModule(object())


class Test_KnownSymbols(ut.TestCase):
    # It is guaranteed that each element of `__known_symbols__` is in
    # each module's namespace.
    # If this test fails, `ImportError` or `AttributeError` may be raised
    # when generating a `comtypes.gen._xxx...` in runtime.
    def _doit(self, mod):
        for s in mod.__known_symbols__:
            self.assertTrue(hasattr(mod, s))

    def test_symbols_in_comtypes(self):
        import comtypes

        self._doit(comtypes)

    def test_symbols_in_comtypes_automation(self):
        import comtypes.automation

        self._doit(comtypes.automation)

    def test_symbols_in_comtypes_typeinfo(self):
        import comtypes.typeinfo

        self._doit(comtypes.typeinfo)

    def test_symbols_in_comtypes_persist(self):
        import comtypes.persist

        self._doit(comtypes.persist)


class Test_CreateObject(ut.TestCase):
    def test_progid(self):
        # create from ProgID
        obj = comtypes.client.CreateObject("Scripting.Dictionary")
        self.assertTrue(isinstance(obj, POINTER(Scripting.IDictionary)))

    def test_clsid(self):
        # create from the CoClass' clsid
        obj = comtypes.client.CreateObject(Scripting.Dictionary)
        self.assertTrue(isinstance(obj, POINTER(Scripting.IDictionary)))

    def test_clsid_string(self):
        # create from string clsid
        comtypes.client.CreateObject(str(Scripting.Dictionary._reg_clsid_))

    def test_remote(self):
        comtypes.client.GetModule("UIAutomationCore.dll")
        from comtypes.gen.UIAutomationClient import (
            CUIAutomation,
            IUIAutomation,
            IUIAutomationElement,
        )

        iuia = comtypes.client.CreateObject(
            CUIAutomation().IPersist_GetClassID(),
            interface=IUIAutomation,
            clsctx=CLSCTX_INPROC_SERVER,
            machine="localhost",
        )
        self.assertIsInstance(iuia, POINTER(IUIAutomation))
        self.assertIsInstance(iuia, IUIAutomation)
        self.assertIsInstance(iuia.GetRootElement(), POINTER(IUIAutomationElement))
        self.assertIsInstance(iuia.GetRootElement(), IUIAutomationElement)

    def test_server_info(self):
        comtypes.client.GetModule("UIAutomationCore.dll")
        from comtypes.gen.UIAutomationClient import (
            CUIAutomation,
            IUIAutomation,
            IUIAutomationElement,
        )

        serverinfo = COSERVERINFO()
        serverinfo.pwszName = "localhost"
        pServerInfo = byref(serverinfo)
        with self.assertRaises(ValueError):
            # cannot set both the machine name and server info
            comtypes.client.CreateObject(
                CUIAutomation().IPersist_GetClassID(),
                interface=IUIAutomation,
                clsctx=CLSCTX_INPROC_SERVER,
                machine="localhost",
                pServerInfo=pServerInfo,
            )
        iuia = comtypes.client.CreateObject(
            CUIAutomation().IPersist_GetClassID(),
            interface=IUIAutomation,
            clsctx=CLSCTX_INPROC_SERVER,
            pServerInfo=pServerInfo,
        )
        self.assertIsInstance(iuia, POINTER(IUIAutomation))
        self.assertIsInstance(iuia, IUIAutomation)
        self.assertIsInstance(iuia.GetRootElement(), POINTER(IUIAutomationElement))
        self.assertIsInstance(iuia.GetRootElement(), IUIAutomationElement)


class Test_Constants(ut.TestCase):
    def test_punk(self):
        obj = comtypes.client.CreateObject(Scripting.Dictionary)
        consts = comtypes.client.Constants(obj)
        self.assertEqual(consts.BinaryCompare, Scripting.BinaryCompare)
        self.assertEqual(consts.TextCompare, Scripting.TextCompare)
        self.assertEqual(consts.DatabaseCompare, Scripting.DatabaseCompare)
        with self.assertRaises(AttributeError):
            consts.Foo
        CompareMethod = consts.CompareMethod
        self.assertEqual(CompareMethod.BinaryCompare, Scripting.BinaryCompare)
        self.assertEqual(CompareMethod.TextCompare, Scripting.TextCompare)
        self.assertEqual(CompareMethod.DatabaseCompare, Scripting.DatabaseCompare)
        with self.assertRaises(AttributeError):
            CompareMethod.Foo
        with self.assertRaises(AttributeError):
            CompareMethod.TextCompare = 1
        with self.assertRaises(AttributeError):
            CompareMethod.Foo = 1
        with self.assertRaises(TypeError):
            CompareMethod["Foo"] = 1
        with self.assertRaises(TypeError):
            del CompareMethod["Foo"]
        with self.assertRaises(TypeError):
            CompareMethod |= {"Foo": 3}
        with self.assertRaises(TypeError):
            CompareMethod.clear()
        with self.assertRaises(TypeError):
            CompareMethod.pop("TextCompare")
        with self.assertRaises(TypeError):
            CompareMethod.popitem()
        with self.assertRaises(TypeError):
            CompareMethod.setdefault("Bar", 3)

    def test_alias(self):
        obj = comtypes.client.CreateObject(Scripting.FileSystemObject)
        consts = comtypes.client.Constants(obj)
        StandardStreamTypes = consts.StandardStreamTypes
        real_name = "__MIDL___MIDL_itf_scrrun_0001_0001_0003"
        self.assertEqual(StandardStreamTypes, getattr(consts, real_name))
        self.assertEqual(StandardStreamTypes.StdIn, Scripting.StdIn)
        self.assertEqual(StandardStreamTypes.StdOut, Scripting.StdOut)
        self.assertEqual(StandardStreamTypes.StdErr, Scripting.StdErr)

    def test_progid(self):
        consts = comtypes.client.Constants("scrrun.dll")
        self.assertEqual(consts.BinaryCompare, Scripting.BinaryCompare)
        self.assertEqual(consts.TextCompare, Scripting.TextCompare)
        self.assertEqual(consts.DatabaseCompare, Scripting.DatabaseCompare)

    def test_returns_other_than_enum_members(self):
        obj = comtypes.client.CreateObject("SAPI.SpVoice")
        from comtypes.gen import SpeechLib as sapi

        consts = comtypes.client.Constants(obj)
        # int (Constant c_int)
        self.assertEqual(consts.Speech_Max_Word_Length, sapi.Speech_Max_Word_Length)
        # str (Constant BSTR)
        self.assertEqual(
            consts.SpeechVoiceSkipTypeSentence, sapi.SpeechVoiceSkipTypeSentence
        )
        self.assertEqual(
            consts.SpeechAudioFormatGUIDWave, sapi.SpeechAudioFormatGUIDWave
        )
        self.assertEqual(
            consts.SpeechRegistryLocalMachineRoot, sapi.SpeechRegistryLocalMachineRoot
        )
        self.assertEqual(
            consts.SpeechGrammarTagDictation, sapi.SpeechGrammarTagDictation
        )
        # float (Constant c_float)
        self.assertEqual(consts.Speech_Default_Weight, sapi.Speech_Default_Weight)

    def test_munged_definitions(self):
        with contextlib.redirect_stdout(None):  # supress warnings
            MSVidCtlLib = comtypes.client.GetModule("msvidctl.dll")
            consts = comtypes.client.Constants("msvidctl.dll")
        # `None` is a Python3 keyword.
        self.assertEqual(consts.MSVidCCService.None_, consts.None_)
        self.assertEqual(MSVidCtlLib.None_, consts.None_)


if __name__ == "__main__":
    ut.main()
