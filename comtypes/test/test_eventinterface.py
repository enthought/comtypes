import gc
import time
import unittest as ut
from ctypes import HRESULT, byref
from ctypes.wintypes import MSG

from comtypes import COMMETHOD, GUID, IUnknown
from comtypes.automation import DISPID
from comtypes.client import CreateObject, GetEvents
from comtypes.messageloop import (
    PM_REMOVE,
    DispatchMessage,
    PeekMessage,
    TranslateMessage,
)

# The primary goal is to verify how `GetEvents` behaves when the
# `interface` argument is explicitly specified versus when it is omitted,
# using an object that has multiple outgoing event interfaces.


class IPropertyNotifySink(IUnknown):
    # https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-ipropertynotifysink
    _iid_ = GUID("{9BFBBC02-EFF1-101A-84ED-00AA00341D07}")
    _methods_ = [
        COMMETHOD([], HRESULT, "OnChanged", (["in"], DISPID, "dispid")),
        COMMETHOD([], HRESULT, "OnRequestEdit", (["in"], DISPID, "dispid")),
    ]


class MSXMLDocumentSink:
    def __init__(self):
        self._events = []

    def onreadystatechange(self, this, *args):
        self._events.append("onreadystatechange")

    def ondataavailable(self, this, *args):
        self._events.append("ondataavailable")

    def OnChanged(self, this, *args):
        self._events.append("OnChanged")

    def OnRequestEdit(self, this, *args):
        self._events.append("OnRequestEdit")


def PumpWaitingMessages():
    msg = MSG()
    while PeekMessage(byref(msg), 0, 0, 0, PM_REMOVE):
        TranslateMessage(byref(msg))
        DispatchMessage(byref(msg))


class Test_MSXML(ut.TestCase):
    def setUp(self):
        self.doc = CreateObject("Msxml2.DOMDocument")
        self.doc.async_ = True

    def tearDown(self):
        del self.doc
        gc.collect()
        time.sleep(2)

    def test_default_eventinterface(self):
        sink = MSXMLDocumentSink()
        conn = GetEvents(self.doc, sink)
        self.doc.loadXML("<root/>")

        for _ in range(50):
            PumpWaitingMessages()
            time.sleep(0.1)
        self.assertIn("onreadystatechange", sink._events)
        self.assertNotIn("OnChanged", sink._events)

        del conn

    def test_nondefault_eventinterface(self):
        sink = MSXMLDocumentSink()

        conn = GetEvents(self.doc, sink, interface=IPropertyNotifySink)

        self.doc.loadXML("<root/>")

        for _ in range(50):
            PumpWaitingMessages()
            time.sleep(0.1)
        self.assertNotIn("onreadystatechange", sink._events)
        self.assertIn("OnChanged", sink._events)

        del conn


if __name__ == "__main__":
    ut.main()
