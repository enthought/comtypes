import contextlib
import unittest
from _ctypes import COMError
from ctypes import POINTER, byref

from comtypes import GUID, hresult
from comtypes.client import CreateObject, GetModule
from comtypes.test.monikers_helper import (
    ROTFLAGS_ALLOWANYCLIENT,
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


class Test_EnumObjectParam(unittest.TestCase):
    def test_cannot_call(self):
        bctx = _create_bctx()
        with self.assertRaises(COMError) as cm:
            # calling `EnumObjectParam` results in a return value of E_NOTIMPL.
            # https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ibindctx-enumobjectparam#notes-to-callers
            bctx.EnumObjectParam()
        self.assertEqual(cm.exception.hresult, hresult.E_NOTIMPL)


class Test_GetRunningObjectTable(unittest.TestCase):
    def test_returns_rot(self):
        vidctl = CreateObject(msvidctl.MSVidCtl, interface=msvidctl.IMSVidCtl)
        item_id = str(GUID.create_new())
        mon = _create_item_moniker("!", item_id)
        bctx = _create_bctx()
        # Before registering: should NOT be running
        rot_from_bctx = bctx.GetRunningObjectTable()
        self.assertIsInstance(rot_from_bctx, IRunningObjectTable)
        rot_from_func = _create_rot()
        dw_reg = rot_from_func.Register(ROTFLAGS_ALLOWANYCLIENT, vidctl, mon)
        # After registering: should be running
        self.assertEqual(rot_from_bctx.IsRunning(mon), hresult.S_OK)
        rot_from_func.Revoke(dw_reg)
        # After revoking: should NOT be running again
        self.assertEqual(rot_from_bctx, rot_from_func)
