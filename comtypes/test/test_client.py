import sys
import unittest as ut
import comtypes.client
from comtypes import COSERVERINFO
from ctypes import POINTER, byref

# create the typelib wrapper and import it
comtypes.client.GetModule("scrrun.dll")
from comtypes.gen import Scripting


if sys.version_info >= (3, 0):
    text_type = str
else:
    text_type = unicode

class Test(ut.TestCase):
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

    def test_GetModule_clsid(self):
        clsid = comtypes.GUID.from_progid("MediaPlayer.MediaPlayer")
        tlib = comtypes.client.GetModule(clsid)

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


if __name__ == "__main__":
    ut.main()
