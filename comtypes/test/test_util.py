import unittest
from ctypes import POINTER, Structure, addressof, c_char, c_double, c_int, cast, sizeof

import comtypes.util
from comtypes import GUID, CoCreateInstance, IUnknown, shelllink


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

    def test_array_offsets(self):
        elms = [10, 20, 30, 40]
        arr = (c_int * 4)(*elms)  # Create an array
        # Test accessing different elements via offset
        for i, expected in enumerate(elms):
            with self.subTest(index=i, expected=expected):
                ref = comtypes.util.byref_at(arr, offset=sizeof(c_int) * i)
                ptr = cast(ref, POINTER(c_int))
                self.assertEqual(ptr.contents.value, expected)

    def test_pointer_arithmetic(self):
        # Test that byref_at behaves like C pointer arithmetic

        class TestStruct(Structure):
            _fields_ = [
                ("field1", c_int),
                ("field2", c_double),
                ("field3", c_char),
            ]

        struct = TestStruct(123, 3.14, b"X")
        for fname, ftype, expected in [
            ("field1", c_int, 123),
            ("field2", c_double, 3.14),
            ("field3", c_char, b"X"),
        ]:
            with self.subTest(field=fname, type=ftype, expected=expected):
                offset = getattr(TestStruct, fname).offset
                ref = comtypes.util.byref_at(struct, offset)
                ptr = cast(ref, POINTER(ftype))
                self.assertEqual(ptr.contents.value, expected)

    def test_com_interface(self):
        CLSID_ShellLink = GUID("{00021401-0000-0000-C000-000000000046}")
        sc = CoCreateInstance(CLSID_ShellLink, interface=shelllink.IShellLinkA)
        ref = comtypes.util.byref_at(sc, 0)
        ptr = cast(ref, POINTER(POINTER(IUnknown)))
        self.assertEqual(addressof(ptr.contents), addressof(sc))

    def test_large_offset(self):
        # Create a large array to test with large offsets
        arr = (c_int * 100)(*range(100))
        # Test accessing element at index 50 (offset = 50 * sizeof(c_int))
        offset = 50 * sizeof(c_int)
        ref = comtypes.util.byref_at(arr, offset)
        ptr = cast(ref, POINTER(c_int))
        self.assertEqual(ptr.contents.value, 50)

    def test_memory_safety(self):
        for initial in [111, 222, 333, 444]:
            with self.subTest(initial=initial):
                obj = c_int(initial)
                ref = comtypes.util.byref_at(obj, 0)
                ptr = cast(ref, POINTER(c_int))
                # Verify initial value
                self.assertEqual(ptr.contents.value, initial)
                # Modify original objects and verify references still work
                obj.value = 333
                # Verify reference still works after modification
                self.assertEqual(ptr.contents.value, 333)
