import unittest, os
from ctypes import *
from comtypes import BSTR
from comtypes.test import requires

##requires("memleaks")

from comtypes.test.find_memleak import find_memleak

class Test(unittest.TestCase):
    def check_leaks(self, func, limit=0):
        bytes = find_memleak(func)
        self.assertFalse(bytes > limit, "Leaks %d bytes" % bytes)

    def test_creation(self):
        def doit():
            BSTR("abcdef" * 100)
        # It seems this test is unreliable.  Sometimes it leaks 4096
        # bytes, sometimes not.  Try to workaround that...
        self.check_leaks(doit, limit=4096)

    def test_from_param(self):
        def doit():
            BSTR.from_param("abcdef")
        self.check_leaks(doit)

    def test_paramflags(self):
        prototype = WINFUNCTYPE(c_void_p, BSTR)
        func = prototype(("SysStringLen", oledll.oleaut32))
        func.restype = c_void_p
        func.argtypes = (BSTR, )
        def doit():
            func("abcdef")
            func("abc xyz")
            func(BSTR("abc def"))
        self.check_leaks(doit)

    def test_inargs(self):
        SysStringLen = windll.oleaut32.SysStringLen
        SysStringLen.argtypes = BSTR,
        SysStringLen.restype = c_uint

        self.assertEqual(SysStringLen("abc xyz"), 7)
        def doit():
            SysStringLen("abc xyz")
            SysStringLen("abc xyz")
            SysStringLen(BSTR("abc def"))
        self.check_leaks(doit)

if __name__ == "__main__":
    unittest.main()
