import contextlib
import unittest
from _ctypes import COMError
from ctypes import POINTER, byref

from comtypes import GUID, hresult
from comtypes.client import CreateObject, GetModule
from comtypes.test.monikers_helper import (
    MK_E_NEEDGENERIC,
    MKSYS_ITEMMONIKER,
    ROTFLAGS_ALLOWANYCLIENT,
    CLSID_AntiMoniker,
    CLSID_CompositeMoniker,
    CLSID_ItemMoniker,
    _CreateBindCtx,
    _CreateItemMoniker,
    _GetRunningObjectTable,
)

with contextlib.redirect_stdout(None):  # supress warnings
    GetModule("msvidctl.dll")
from comtypes.gen import MSVidCtlLib as msvidctl
from comtypes.gen.MSVidCtlLib import IBindCtx, IMoniker, IRunningObjectTable


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
