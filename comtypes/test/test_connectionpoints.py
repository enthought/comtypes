import contextlib
import unittest as ut
from ctypes import byref

import comtypes.client
from comtypes import COMObject
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

    def test_EnumConnectionPoints(self):
        conn_pts = list(self.cpc.EnumConnectionPoints())
        self.assertGreater(len(conn_pts), 0)
        self.assertTrue(
            all(pt.GetConnectionPointContainer() == self.cpc for pt in conn_pts)
        )

    def test_FindConnectionPoint(self):
        cp = self.cpc.FindConnectionPoint(byref(self.EVENT_IID))
        self.assertEqual(cp.GetConnectionPointContainer(), self.cpc)


class Sink(COMObject):
    _com_interfaces_ = [msvidctl._IMSVidCtlEvents]


class Test_IConnectionPoint(ut.TestCase):
    EVENT_IID = msvidctl._IMSVidCtlEvents._iid_

    def setUp(self):
        self.impl = comtypes.client.CreateObject(
            msvidctl.MSVidCtl, interface=msvidctl.IMSVidCtl
        )
        self.cpc = self.impl.QueryInterface(IConnectionPointContainer)
        self.cp = self.cpc.FindConnectionPoint(byref(self.EVENT_IID))

    def test_GetConnectionInterface(self):
        self.assertEqual(self.cp.GetConnectionInterface(), self.EVENT_IID)

    def test_GetConnectionPointContainer(self):
        self.assertEqual(self.cp.GetConnectionPointContainer(), self.cpc)

    def test_Advise_Unadvise(self):
        self.assertEqual(len(list(self.cp.EnumConnections())), 0)
        sink = Sink()
        # Since `POINTER(IUnknown).from_param`(`_compointer_base.from_param`)
        # can accept a `COMObject` instance, `IConnectionPoint.Advise` can
        # take either a COM object or a COM interface pointer.
        cookie = self.cp.Advise(sink)
        # Verify the connection exists
        self.assertEqual(len(list(self.cp.EnumConnections())), 1)
        self.cp.Unadvise(cookie)
        self.assertEqual(len(list(self.cp.EnumConnections())), 0)
