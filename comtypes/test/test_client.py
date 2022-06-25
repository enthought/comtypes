from ctypes import POINTER, byref
import os
import sys
import unittest as ut

import comtypes.client
from comtypes import COSERVERINFO

# create the typelib wrapper and import it
comtypes.client.GetModule("scrrun.dll")
from comtypes.gen import Scripting


if sys.version_info >= (3, 0):
    text_type = str
else:
    text_type = unicode


class Test_GetModule(ut.TestCase):
    def test_tlib_string(self):
        mod = comtypes.client.GetModule("scrrun.dll")
        self.assertIs(mod, Scripting)

    def test_abspath(self):
        mod = comtypes.client.GetModule(Scripting.typelib_path)
        self.assertIs(mod, Scripting)

    @ut.skipUnless(
        os.path.splitdrive(Scripting.typelib_path)[0] == os.path.splitdrive(__file__)[0],
        "This depends on typelib and test module are in same drive")
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
        comtypes.client.CreateObject(text_type(Scripting.Dictionary._reg_clsid_))
        comtypes.client.CreateObject(str(Scripting.Dictionary._reg_clsid_))

    @ut.skip(
            "This test uses IE which is not available on all machines anymore. "
            "Find another API to use."
    )
    def test_remote(self):
        ie = comtypes.client.CreateObject("InternetExplorer.Application",
                                          machine="localhost")
        self.assertEqual(ie.Visible, False)
        ie.Visible = 1
        # on a remote machine, this may not work.  Probably depends on
        # how the server is run.
        self.assertEqual(ie.Visible, True)
        self.assertEqual(0, ie.Quit()) # 0 == S_OK

    @ut.skip(
            "This test uses IE which is not available on all machines anymore. "
            "Find another API to use."
    )
    def test_server_info(self):
        serverinfo = COSERVERINFO()
        serverinfo.pwszName = 'localhost'
        pServerInfo = byref(serverinfo)

        self.assertRaises(ValueError, comtypes.client.CreateObject,
                "InternetExplorer.Application", machine='localhost',
                pServerInfo=pServerInfo)
        ie = comtypes.client.CreateObject("InternetExplorer.Application",
                                          pServerInfo=pServerInfo)
        self.assertEqual(ie.Visible, False)
        ie.Visible = 1
        # on a remote machine, this may not work.  Probably depends on
        # how the server is run.
        self.assertEqual(ie.Visible, True)
        self.assertEqual(0, ie.Quit()) # 0 == S_OK


class Test_Constants(ut.TestCase):
    def test_punk(self):
        obj = comtypes.client.CreateObject(Scripting.Dictionary)
        consts = comtypes.client.Constants(obj)
        self.assertEqual(consts.BinaryCompare, Scripting.BinaryCompare)
        self.assertEqual(consts.TextCompare, Scripting.TextCompare)
        self.assertEqual(consts.DatabaseCompare, Scripting.DatabaseCompare)
        with self.assertRaises(AttributeError):
            consts.CompareMethod

    def test_returns_other_than_int(self):
        obj = comtypes.client.CreateObject("SAPI.SpVoice")
        from comtypes.gen import SpeechLib as sapi
        consts = comtypes.client.Constants(obj)
        # str (Constant BSTR)
        self.assertEqual(consts.SpeechVoiceSkipTypeSentence, sapi.SpeechVoiceSkipTypeSentence)
        self.assertEqual(consts.SpeechAudioFormatGUIDWave, sapi.SpeechAudioFormatGUIDWave)
        self.assertEqual(consts.SpeechRegistryLocalMachineRoot, sapi.SpeechRegistryLocalMachineRoot)
        self.assertEqual(consts.SpeechGrammarTagDictation, sapi.SpeechGrammarTagDictation)
        # float (Constant c_float)
        self.assertEqual(consts.Speech_Default_Weight, sapi.Speech_Default_Weight)


if __name__ == "__main__":
    ut.main()
