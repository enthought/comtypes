import unittest
from ctypes import POINTER, addressof, c_char, c_double, c_int, cast

import comtypes.util


class ByrefAtTest(unittest.TestCase):
    def test_ctypes(self):
        for ctype, value in [
            (c_int, 42),
            (c_double, 3.14),
            (c_char, b"A"),
        ]:
            with self.subTest(ctype=ctype, value=value):
                obj = ctype(value)
                # Test with zero offset - should point to the same location
                ref = comtypes.util.byref_at(obj, 0)
                ptr = cast(ref, POINTER(ctype))
                # byref objects don't have contents, but we can cast them to pointers
                self.assertEqual(addressof(ptr.contents), addressof(obj))
                self.assertEqual(ptr.contents.value, value)
