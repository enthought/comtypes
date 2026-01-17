import contextlib
import unittest
from ctypes import POINTER, byref

from comtypes import GUID, hresult
from comtypes.client import CreateObject, GetModule
from comtypes.test.monikers_helper import (
    MKSYS_ITEMMONIKER,
    ROTFLAGS_ALLOWANYCLIENT,
    CLSID_AntiMoniker,
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
