import unittest
from comtypes.client import CreateObject
from comtypes.client.dynamic import _Dispatch

class Test(unittest.TestCase):
    def test(self):
        d = CreateObject("MSScriptControl.ScriptControl")
        d.Language = "jscript"
        d.AddCode('function x() { return [3, "spam foo", 3.14]; };')
        result = d.Run("x", [])
        self.failUnless(isinstance(result, _Dispatch))
        self.failUnlessEqual(result[0], 3)
        self.failUnlessEqual(result[1], "spam foo")
        self.failUnlessEqual(result[2], 3.14)
        self.assertRaises(IndexError, lambda: result[3])

    def test_with_args(self):
        d = CreateObject("MSScriptControl.ScriptControl")
        d.Language = "jscript"
        d.AddCode('function x(a1, a2) { return [3, "spam foo", 3.14, a1, a2]; };')
        result = d.Run("x", [42, 96])
        self.failUnless(isinstance(result, _Dispatch))
        self.failUnlessEqual(result[0], 3)
        self.failUnlessEqual(result[1], "spam foo")
        self.failUnlessEqual(result[2], 3.14)
        self.failUnlessEqual(result[3], 42)
        self.failUnlessEqual(result[4], 96)
        self.assertRaises(IndexError, lambda: result[5])

if __name__ == "__main__":
    unittest.main()
