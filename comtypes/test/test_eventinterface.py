import unittest as ut
from ctypes import byref
from ctypes.wintypes import MSG

from comtypes.client import CreateObject, GetEvents
from comtypes.messageloop import (
    PM_REMOVE,
    DispatchMessage,
    PeekMessage,
    TranslateMessage,
)

# FIXME: External test dependencies like this seem bad.  Find a different
# built-in win32 API to use.
# The primary goal is to verify how `GetEvents` behaves when the
# `interface` argument is explicitly specified versus when it is omitted,
# using an object that has multiple outgoing event interfaces.


class EventSink:
    def __init__(self):
        self._events = []

    # some DWebBrowserEvents
    def OnVisible(self, this, *args):
        # print "OnVisible", args
        self._events.append("OnVisible")

    def BeforeNavigate(self, this, *args):
        # print "BeforeNavigate", args
        self._events.append("BeforeNavigate")

    def NavigateComplete(self, this, *args):
        # print "NavigateComplete", args
        self._events.append("NavigateComplete")

    # some DWebBrowserEvents2
    def BeforeNavigate2(self, this, *args):
        # print "BeforeNavigate2", args
        self._events.append("BeforeNavigate2")

    def NavigateComplete2(self, this, *args):
        # print "NavigateComplete2", args
        self._events.append("NavigateComplete2")

    def DocumentComplete(self, this, *args):
        # print "DocumentComplete", args
        self._events.append("DocumentComplete")


def PumpWaitingMessages():
    msg = MSG()
    while PeekMessage(byref(msg), 0, 0, 0, PM_REMOVE):
        TranslateMessage(byref(msg))
        DispatchMessage(byref(msg))


class Test(ut.TestCase):
    def tearDown(self):
        import gc

        gc.collect()
        import time

        time.sleep(2)

    @ut.skip(
        "External test dependencies like this seem bad.  Find a different built-in "
        "win32 API to use."
    )
    def test_default_eventinterface(self):
        sink = EventSink()
        ie = CreateObject("InternetExplorer.Application")
        conn = GetEvents(ie, sink=sink)
        ie.Visible = True
        ie.Navigate2(URL="http://docs.python.org/", Flags=0)
        import time

        for i in range(50):
            PumpWaitingMessages()
            time.sleep(0.1)
        ie.Visible = False
        ie.Quit()

        self.assertEqual(
            sink._events,
            [
                "OnVisible",
                "BeforeNavigate2",
                "NavigateComplete2",
                "DocumentComplete",
                "OnVisible",
            ],
        )

        del ie
        del conn

    @ut.skip(
        "External test dependencies like this seem bad.  Find a different built-in "
        "win32 API to use."
    )
    def test_nondefault_eventinterface(self):
        sink = EventSink()
        ie = CreateObject("InternetExplorer.Application")
        import comtypes.gen.SHDocVw as mod

        conn = GetEvents(ie, sink, interface=mod.DWebBrowserEvents)

        ie.Visible = True
        ie.Navigate2(Flags=0, URL="http://docs.python.org/")
        import time

        for i in range(50):
            PumpWaitingMessages()
            time.sleep(0.1)
        ie.Visible = False
        ie.Quit()

        self.assertEqual(sink._events, ["BeforeNavigate", "NavigateComplete"])
        del ie


if __name__ == "__main__":
    ut.main()
