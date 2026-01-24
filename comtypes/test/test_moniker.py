import contextlib
import ctypes
import os
import tempfile
import unittest
from _ctypes import COMError
from ctypes import POINTER, WinDLL, byref
from ctypes.wintypes import DWORD, LPCWSTR, LPWSTR, MAX_PATH
from pathlib import Path

from comtypes import GUID, hresult
from comtypes.client import CreateObject, GetModule
from comtypes.persist import IPersistFile
from comtypes.test.monikers_helper import (
    MK_E_NEEDGENERIC,
    MKSYS_FILEMONIKER,
    MKSYS_GENERICCOMPOSITE,
    MKSYS_ITEMMONIKER,
    ROTFLAGS_ALLOWANYCLIENT,
    CLSID_AntiMoniker,
    CLSID_CompositeMoniker,
    CLSID_FileMoniker,
    CLSID_ItemMoniker,
    _CreateBindCtx,
    _CreateFileMoniker,
    _CreateGenericComposite,
    _CreateItemMoniker,
    _GetRunningObjectTable,
)

with contextlib.redirect_stdout(None):  # supress warnings
    GetModule("msvidctl.dll")
from comtypes.gen import MSVidCtlLib as msvidctl
from comtypes.gen.MSVidCtlLib import (
    IBindCtx,
    IEnumMoniker,
    IMoniker,
    IRunningObjectTable,
)

_kernel32 = WinDLL("kernel32")

_GetLongPathNameW = _kernel32.GetLongPathNameW
_GetLongPathNameW.argtypes = [LPCWSTR, LPWSTR, DWORD]
_GetLongPathNameW.restype = DWORD


def _get_long_path_name(path: str) -> str:
    """Converts a path to its long form using GetLongPathNameW."""
    buffer = ctypes.create_unicode_buffer(MAX_PATH)
    length = _GetLongPathNameW(path, buffer, MAX_PATH)
    return buffer.value[:length]


def _create_generic_composite(mk_first: IMoniker, mk_rest: IMoniker) -> IMoniker:
    mon = POINTER(IMoniker)()
    _CreateGenericComposite(mk_first, mk_rest, byref(mon))
    return mon  # type: ignore


def _create_file_moniker(path: str) -> IMoniker:
    mon = POINTER(IMoniker)()
    _CreateFileMoniker(path, byref(mon))
    return mon  # type: ignore


def _create_item_moniker(delim: str, item: str) -> IMoniker:
    mon = POINTER(IMoniker)()
    _CreateItemMoniker(delim, item, byref(mon))
    return mon  # type: ignore


def _create_bctx() -> IBindCtx:
    bctx = POINTER(IBindCtx)()
    # The first parameter is reserved and must be 0.
    _CreateBindCtx(0, byref(bctx))
    return bctx  # type: ignore


def _create_rot() -> IRunningObjectTable:
    rot = POINTER(IRunningObjectTable)()
    # The first parameter is reserved and must be 0.
    _GetRunningObjectTable(0, byref(rot))
    return rot  # type: ignore


class Test_IsSystemMoniker_GetDisplayName_Inverse(unittest.TestCase):
    def test_generic_composite(self):
        item_id1 = str(GUID.create_new())
        item_id2 = str(GUID.create_new())
        mon = _create_generic_composite(
            _create_item_moniker("!", item_id1),
            _create_item_moniker("!", item_id2),
        )
        self.assertEqual(mon.IsSystemMoniker(), MKSYS_GENERICCOMPOSITE)
        bctx = _create_bctx()
        self.assertEqual(mon.GetDisplayName(bctx, None), f"!{item_id1}!{item_id2}")
        self.assertEqual(mon.GetClassID(), CLSID_CompositeMoniker)
        self.assertEqual(mon.Inverse().GetClassID(), CLSID_CompositeMoniker)

    def test_file(self):
        with tempfile.NamedTemporaryFile() as f:
            mon = _create_file_moniker(f.name)
            self.assertEqual(mon.IsSystemMoniker(), MKSYS_FILEMONIKER)
            bctx = _create_bctx()
            self.assertEqual(
                os.path.normcase(
                    os.path.normpath(
                        _get_long_path_name(mon.GetDisplayName(bctx, None))
                    )
                ),
                os.path.normcase(os.path.normpath(_get_long_path_name(f.name))),
            )
            self.assertEqual(mon.GetClassID(), CLSID_FileMoniker)
            self.assertEqual(mon.Inverse().GetClassID(), CLSID_AntiMoniker)

    def test_item(self):
        item_id = str(GUID.create_new())
        mon = _create_item_moniker("!", item_id)
        self.assertEqual(mon.IsSystemMoniker(), MKSYS_ITEMMONIKER)
        bctx = _create_bctx()
        self.assertEqual(mon.GetDisplayName(bctx, None), f"!{item_id}")
        self.assertEqual(mon.GetClassID(), CLSID_ItemMoniker)
        self.assertEqual(mon.Inverse().GetClassID(), CLSID_AntiMoniker)


class Test_ComposeWith(unittest.TestCase):
    def test_item(self):
        item_id = str(GUID.create_new())
        mon = _create_item_moniker("!", item_id)
        item_mon2 = _create_item_moniker("!", str(GUID.create_new()))
        self.assertEqual(
            mon.ComposeWith(item_mon2, False).GetClassID(),
            CLSID_CompositeMoniker,
        )
        with self.assertRaises(COMError) as cm:
            mon.ComposeWith(item_mon2, True)
        self.assertEqual(cm.exception.hresult, MK_E_NEEDGENERIC)


class Test_IsEqual(unittest.TestCase):
    def test_item(self):
        item_id = str(GUID.create_new())
        mon1 = _create_item_moniker("!", item_id)
        mon2 = _create_item_moniker("!", item_id)  # Should be equal
        mon3 = _create_item_moniker("!", str(GUID.create_new()))  # Should not be equal
        self.assertEqual(mon1.IsEqual(mon2), hresult.S_OK)
        self.assertEqual(mon1.IsEqual(mon3), hresult.S_FALSE)


class Test_Hash(unittest.TestCase):
    def test_item(self):
        item_id = str(GUID.create_new())
        mon1 = _create_item_moniker("!", item_id)
        mon2 = _create_item_moniker("!", item_id)  # Should be equal
        mon3 = _create_item_moniker("!", str(GUID.create_new()))  # Should not be equal
        self.assertEqual(mon1.Hash(), mon2.Hash())
        self.assertNotEqual(mon1.Hash(), mon3.Hash())
        self.assertNotEqual(mon2.Hash(), mon3.Hash())


class Test_IsRunning(unittest.TestCase):
    def test_item(self):
        vidctl = CreateObject(msvidctl.MSVidCtl, interface=msvidctl.IMSVidCtl)
        item_id = str(GUID.create_new())
        mon = _create_item_moniker("!", item_id)
        rot = _create_rot()
        bctx = _create_bctx()
        # Before registering: should NOT be running
        self.assertEqual(mon.IsRunning(bctx, None, None), hresult.S_FALSE)
        dw_reg = rot.Register(ROTFLAGS_ALLOWANYCLIENT, vidctl, mon)
        # After registering: should be running
        self.assertEqual(mon.IsRunning(bctx, None, None), hresult.S_OK)
        rot.Revoke(dw_reg)
        # After revoking: should NOT be running again
        self.assertEqual(mon.IsRunning(bctx, None, None), hresult.S_FALSE)


class Test_CommonPrefixWith(unittest.TestCase):
    def test_file(self):
        bctx = _create_bctx()
        # Create temporary directories and files for realistic File Monikers
        with tempfile.TemporaryDirectory() as t:
            tmpdir = Path(t)
            dir_a = tmpdir / "dir_a"
            dir_b = tmpdir / "dir_a" / "dir_b"
            dir_b.mkdir(parents=True)
            file1 = dir_a / "file1.txt"
            file2 = dir_b / "file2.txt"
            file3 = tmpdir / "file3.txt"
            mon1 = _create_file_moniker(str(file1))  # tmpdir/dir_a/file1.txt
            mon2 = _create_file_moniker(str(file2))  # tmpdir/dir_a/dir_b/file2.txt
            mon3 = _create_file_moniker(str(file3))  # tmpdir/file3.txt
            # Common prefix between mon1 and mon2 (tmpdir/dir_a)
            self.assertEqual(
                os.path.normcase(
                    os.path.normpath(
                        mon1.CommonPrefixWith(mon2).GetDisplayName(bctx, None)
                    )
                ),
                os.path.normcase(os.path.normpath(dir_a)),
            )
            # Common prefix between mon1 and mon3 (tmpdir)
            self.assertEqual(
                os.path.normcase(
                    os.path.normpath(
                        mon1.CommonPrefixWith(mon3).GetDisplayName(bctx, None)
                    )
                ),
                os.path.normcase(os.path.normpath(tmpdir)),
            )


class Test_RelativePathTo(unittest.TestCase):
    def test_file(self):
        bctx = _create_bctx()
        with tempfile.TemporaryDirectory() as t:
            tmpdir = Path(t)
            dir_a = tmpdir / "dir_a"
            dir_b = tmpdir / "dir_b"
            dir_a.mkdir()
            dir_b.mkdir()
            file1 = dir_a / "file1.txt"
            file2 = dir_b / "file2.txt"
            mon_from = _create_file_moniker(str(file1))  # tmpdir/dir_a/file1.txt
            mon_to = _create_file_moniker(str(file2))  # tmpdir/dir_b/file2.txt
            # The COM API returns paths with backslashes on Windows, so we normalize.
            self.assertEqual(
                # Check the display name of the relative moniker
                # The moniker's `RelativePathTo` method calculates the path from
                # the base of the `mon_from` to the target `mon_to`.
                os.path.normcase(
                    os.path.normpath(
                        mon_from.RelativePathTo(mon_to).GetDisplayName(bctx, None)
                    )
                ),
                # Calculate the relative path from the directory of file1 to file2
                os.path.normcase(
                    os.path.normpath(file2.relative_to(file1, walk_up=True))
                ),
            )


class Test_Enum(unittest.TestCase):
    def test_generic_composite(self):
        item_id1 = str(GUID.create_new())
        item_id2 = str(GUID.create_new())
        item_mon1 = _create_item_moniker("!", item_id1)
        item_mon2 = _create_item_moniker("!", item_id2)
        # Create a composite moniker to ensure multiple elements for enumeration
        comp_mon = _create_generic_composite(item_mon1, item_mon2)
        enum_moniker = comp_mon.Enum(True)  # True for forward enumeration
        self.assertIsInstance(enum_moniker, IEnumMoniker)


class Test_RemoteBindToObject(unittest.TestCase):
    def test_file(self):
        bctx = _create_bctx()
        with tempfile.TemporaryDirectory() as t:
            tmpdir = Path(t)
            tmpfile = tmpdir / "tmp.lnk"
            tmpfile.touch()
            mon = _create_file_moniker(str(tmpfile))
            bound_obj = mon.RemoteBindToObject(bctx, None, IPersistFile._iid_)
            pf = bound_obj.QueryInterface(IPersistFile)
            self.assertEqual(
                os.path.normcase(os.path.normpath(_get_long_path_name(str(tmpfile)))),
                os.path.normcase(
                    os.path.normpath(_get_long_path_name(pf.GetCurFile()))
                ),
            )
