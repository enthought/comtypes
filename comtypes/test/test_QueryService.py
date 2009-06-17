import unittest
from ctypes import POINTER
import comtypes
import comtypes.client

class TestCase(unittest.TestCase):
    def test(self):
        comtypes.client.GetModule('oleacc.dll')
        from comtypes.gen.Accessibility import IAccessible

        ie = comtypes.client.CreateObject('InternetExplorer.application')
        ie.navigate2("about:blank",0)
        sp = ie.Document.Body.QueryInterface(comtypes.IServiceProvider)
        pacc = sp.QueryService(IAccessible._iid_, IAccessible)
        ie.Quit()
        self.failUnlessEqual(type(pacc), POINTER(IAccessible))

if __name__ == "__main__":
    unittest.main()

