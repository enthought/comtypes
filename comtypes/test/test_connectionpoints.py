import contextlib
import unittest as ut
from ctypes import byref

import comtypes.client
from comtypes.connectionpoints import IConnectionPointContainer

with contextlib.redirect_stdout(None):  # supress warnings
    comtypes.client.GetModule("msvidctl.dll")
from comtypes.gen import MSVidCtlLib as msvidctl


class Test_IConnectionPointContainer(ut.TestCase):
    EVENT_IID = msvidctl._IMSVidCtlEvents._iid_

    def setUp(self):
        self.impl = comtypes.client.CreateObject(
            msvidctl.MSVidCtl, interface=msvidctl.IMSVidCtl
        )
        self.cpc = self.impl.QueryInterface(IConnectionPointContainer)

    def test_FindConnectionPoint(self):
        cp = self.cpc.FindConnectionPoint(byref(self.EVENT_IID))
        self.assertEqual(cp.GetConnectionPointContainer(), self.cpc)
