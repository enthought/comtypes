import os
import tempfile
import unittest
from _ctypes import COMError
from ctypes import HRESULT, POINTER, OleDLL, byref, c_ubyte
from ctypes.wintypes import DWORD, PWCHAR
from pathlib import Path
from typing import Optional

import comtypes
import comtypes.client

comtypes.client.GetModule("portabledeviceapi.dll")
from comtypes.gen.PortableDeviceApiLib import IStorage, tagSTATSTG

STGTY_STORAGE = 1

STATFLAG_DEFAULT = 0
STGC_DEFAULT = 0
STGM_CREATE = 0x00001000
STGM_DIRECT = 0x00000000
STGM_DELETEONRELEASE = 0x04000000
STGM_READ = 0x00000000
STGM_READWRITE = 0x00000002
STGM_SHARE_EXCLUSIVE = 0x00000010
STGM_TRANSACTED = 0x00010000
STGMOVE_MOVE = 0
STREAM_SEEK_SET = 0

STG_E_PATHNOTFOUND = -2147287038
STG_E_INVALIDFLAG = -2147286785

_ole32 = OleDLL("ole32")

_StgCreateDocfile = _ole32.StgCreateDocfile
_StgCreateDocfile.argtypes = [PWCHAR, DWORD, DWORD, POINTER(POINTER(IStorage))]
_StgCreateDocfile.restype = HRESULT


class Test_IStorage(unittest.TestCase):
    RW_EXCLUSIVE = STGM_READWRITE | STGM_SHARE_EXCLUSIVE
    RW_EXCLUSIVE_TX = RW_EXCLUSIVE | STGM_TRANSACTED
    RW_EXCLUSIVE_CREATE = RW_EXCLUSIVE | STGM_CREATE
    CREATE_TESTDOC = STGM_DIRECT | STGM_CREATE | RW_EXCLUSIVE
    CREATE_TEMP_TESTDOC = CREATE_TESTDOC | STGM_DELETEONRELEASE

    def _create_docfile(self, mode: int, name: Optional[str] = None) -> IStorage:
        stg = POINTER(IStorage)()
        _StgCreateDocfile(name, mode, 0, byref(stg))
        return stg  # type: ignore

    def test_CreateStream(self):
        storage = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        # When created with `StgCreateDocfile(NULL, ...)`, `pwcsName` is a
        # temporary filename. The file really exists on disk because Windows
        # creates an actual temporary file for the compound storage.
        filepath = Path(storage.Stat(STATFLAG_DEFAULT).pwcsName)
        self.assertTrue(filepath.exists())
        stream = storage.CreateStream("example", self.RW_EXCLUSIVE_CREATE, 0, 0)
        test_data = b"Some data"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        stream.RemoteWrite(pv, len(test_data))
        stream.Commit(STGC_DEFAULT)
        stream.RemoteSeek(0, STREAM_SEEK_SET)
        read_buffer, data_read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(read_buffer)[0:data_read], test_data)
        storage.Commit(STGC_DEFAULT)
        self.assertTrue(filepath.exists())
        del storage
        self.assertFalse(filepath.exists())

    # TODO: Auto-generated methods based on type info are remote-side and hard
    #       to call from the client.
    #       If a proper invocation method or workaround is found, testing
    #       becomes possible.
    #       See: https://github.com/enthought/comtypes/issues/607
    # def test_RemoteOpenStream(self):
    #     pass

    def test_CreateStorage(self):
        parent = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        child = parent.CreateStorage("child", self.RW_EXCLUSIVE_TX, 0, 0)
        self.assertEqual("child", child.Stat(STATFLAG_DEFAULT).pwcsName)

    def test_OpenStorage(self):
        parent = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        with self.assertRaises(COMError) as cm:
            parent.OpenStorage("child", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual(cm.exception.hresult, STG_E_PATHNOTFOUND)
        parent.CreateStorage("child", self.RW_EXCLUSIVE_TX, 0, 0)
        child = parent.OpenStorage("child", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual("child", child.Stat(STATFLAG_DEFAULT).pwcsName)

    def test_RemoteCopyTo(self):
        src_stg = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        src_stg.CreateStorage("child", self.RW_EXCLUSIVE_TX, 0, 0)
        dst_stg = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        src_stg.RemoteCopyTo(0, None, None, dst_stg)
        src_stg.Commit(STGC_DEFAULT)
        del src_stg
        opened_stg = dst_stg.OpenStorage("child", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual("child", opened_stg.Stat(STATFLAG_DEFAULT).pwcsName)

    def test_MoveElementTo(self):
        src_stg = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        src_stg.CreateStorage("foo", self.RW_EXCLUSIVE_TX, 0, 0)
        dst_stg = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        src_stg.MoveElementTo("foo", dst_stg, "bar", STGMOVE_MOVE)
        opened_stg = dst_stg.OpenStorage("bar", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual("bar", opened_stg.Stat(STATFLAG_DEFAULT).pwcsName)
        with self.assertRaises(COMError) as cm:
            src_stg.OpenStorage("foo", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual(cm.exception.hresult, STG_E_PATHNOTFOUND)

    def test_Revert(self):
        storage = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        foo = storage.CreateStorage("foo", self.RW_EXCLUSIVE_TX, 0, 0)
        foo.CreateStorage("bar", self.RW_EXCLUSIVE_TX, 0, 0)
        bar = foo.OpenStorage("bar", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual("bar", bar.Stat(STATFLAG_DEFAULT).pwcsName)
        foo.Revert()
        with self.assertRaises(COMError) as cm:
            foo.OpenStorage("bar", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual(cm.exception.hresult, STG_E_PATHNOTFOUND)

    # TODO: Auto-generated methods based on type info are remote-side and hard
    #       to call from the client.
    #       If a proper invocation method or workaround is found, testing
    #       becomes possible.
    #       See: https://github.com/enthought/comtypes/issues/607
    # def test_RemoteEnumElements(self):
    #     pass

    def test_DestroyElement(self):
        storage = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        storage.CreateStorage("example", self.RW_EXCLUSIVE_TX, 0, 0)
        storage.DestroyElement("example")
        with self.assertRaises(COMError) as cm:
            storage.OpenStorage("example", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual(cm.exception.hresult, STG_E_PATHNOTFOUND)

    def test_RenameElement(self):
        storage = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        storage.CreateStorage("example", self.RW_EXCLUSIVE_TX, 0, 0)
        storage.RenameElement("example", "sample")
        sample = storage.OpenStorage("sample", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual("sample", sample.Stat(STATFLAG_DEFAULT).pwcsName)
        with self.assertRaises(COMError) as cm:
            storage.OpenStorage("example", None, self.RW_EXCLUSIVE_TX, None, 0)
        self.assertEqual(cm.exception.hresult, STG_E_PATHNOTFOUND)

    def test_SetClass(self):
        storage = self._create_docfile(mode=self.CREATE_TEMP_TESTDOC)
        # Initial value is CLSID_NULL.
        self.assertEqual(storage.Stat(STATFLAG_DEFAULT).clsid, comtypes.GUID())
        new_clsid = comtypes.GUID.create_new()
        storage.SetClass(new_clsid)
        self.assertEqual(storage.Stat(STATFLAG_DEFAULT).clsid, new_clsid)
        # Re-set CLSID to CLSID_NULL and verify it is correctly set.
        storage.SetClass(comtypes.GUID())
        self.assertEqual(storage.Stat(STATFLAG_DEFAULT).clsid, comtypes.GUID())

    def test_Stat(self):
        with tempfile.TemporaryDirectory() as t:
            tmpdir = Path(t)
            tmpfile = tmpdir / "test_docfile.cfs"
            self.assertFalse(tmpfile.exists())
            # When created with `StgCreateDocfile(filepath_string, ...)`, the
            # compound file is created at that location.
            storage = self._create_docfile(
                name=str(tmpfile), mode=self.CREATE_TEMP_TESTDOC
            )
            self.assertTrue(tmpfile.exists())
            with self.assertRaises(COMError) as cm:
                storage.Stat(0xFFFFFFFF)  # Invalid flag
            self.assertEqual(cm.exception.hresult, STG_E_INVALIDFLAG)
            stat = storage.Stat(STATFLAG_DEFAULT)
            self.assertIsInstance(stat, tagSTATSTG)
            self.assertEqual(
                os.path.normcase(os.path.normpath(Path(stat.pwcsName))),
                os.path.normcase(os.path.normpath(tmpfile)),
            )
            del storage  # Release the storage to prevent 'cannot access the file ...'
        self.assertEqual(stat.type, STGTY_STORAGE)
        # Due to header overhead and file system allocation, the size may be
        # greater than 0 bytes.
        self.assertGreaterEqual(stat.cbSize, 0)
        # `grfMode` should reflect the access mode flags from creation.
        self.assertEqual(stat.grfMode, self.RW_EXCLUSIVE | STGM_DIRECT)
        self.assertEqual(stat.grfLocksSupported, 0)
        self.assertEqual(stat.clsid, comtypes.GUID())  # CLSID_NULL for new creation.
        self.assertEqual(stat.grfStateBits, 0)
