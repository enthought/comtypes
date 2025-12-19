import contextlib
import time
import unittest
from ctypes import POINTER

import comtypes
from comtypes import GUID
from comtypes.client import CreateObject, GetModule

with contextlib.redirect_stdout(None):  # supress warnings
    GetModule("mshtml.tlb")
import comtypes.gen.MSHTML as mshtml

SID_SHTMLEditServices = GUID("{3050F7F9-98B5-11CF-BB82-00AA00BDCE0B}")


class TestCase(unittest.TestCase):
    def test(self):
        doc = CreateObject(mshtml.HTMLDocument, interface=mshtml.IHTMLDocument2)
        doc.designMode = "On"
        doc.write("<html><body><div id='test'>Hello</div></body></html>")
        doc.close()
        while doc.readyState != "complete":
            time.sleep(0.01)
        sp = doc.QueryInterface(comtypes.IServiceProvider)
        # This behavior is described in Microsoft documentation:
        # https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa704048(v=vs.85)
        es = sp.QueryService(SID_SHTMLEditServices, mshtml.IHTMLEditServices)
        self.assertIsInstance(es, POINTER(mshtml.IHTMLEditServices))
        mc = doc.QueryInterface(mshtml.IMarkupContainer)
        ss = es.GetSelectionServices(mc)
        self.assertIsInstance(ss, POINTER(mshtml.ISelectionServices))


if __name__ == "__main__":
    unittest.main()
