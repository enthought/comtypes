import unittest
from ctypes import WinDLL
from ctypes.wintypes import UINT

from comtypes import BSTR
from comtypes.test.find_memleak import find_memleak


class Test(unittest.TestCase):
    def check_leaks(self, func, limit=0):
        # BSTRs are allocated with `SysAllocString`, which goes through the COM
        # task allocator, so `mallocspy` counts them exactly. This avoids the
        # `working_set` false positive where the first call in a process faults
        # in pages that look like a leak
        # (see https://github.com/enthought/comtypes/issues/971).
        bytes = find_memleak(func, target="mallocspy")
        self.assertFalse(bytes > limit, f"Leaks {bytes} bytes")

    def test_creation(self):
        def doit():
            BSTR("abcdef" * 100)

        self.check_leaks(doit)

    def test_from_param(self):
        def doit():
            BSTR.from_param("abcdef")

        self.check_leaks(doit)

    def test_inargs(self):
        oleaut32 = WinDLL("oleaut32")

        SysStringLen = oleaut32.SysStringLen
        SysStringLen.argtypes = [BSTR]
        SysStringLen.restype = UINT

        self.assertEqual(SysStringLen("abc xyz"), 7)

        def doit():
            SysStringLen("abc xyz")
            SysStringLen("abc xyz")
            SysStringLen(BSTR("abc def"))

        self.check_leaks(doit)


if __name__ == "__main__":
    unittest.main()
