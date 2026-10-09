import ctypes
import unittest

from comtypes import automation, typeinfo
from comtypes.tools.codegenerator.helpers import TypeNamer
from comtypes.tools.tlbparser import Parser


class ByteTypeTest(unittest.TestCase):
    def make_ctype(self, desc):
        # Exercise the ctypes expression used in generated COM signatures.
        typ = Parser().make_type(desc, None)
        namespace = dict(vars(ctypes), STRING=ctypes.c_char_p)
        return eval(TypeNamer()(typ), namespace)

    def test_signed_byte(self):
        desc = typeinfo.TYPEDESC()
        desc.vt = automation.VT_I1
        ctype = self.make_ctype(desc)
        for value in (-128, -7, 0, 127):
            with self.subTest(value=value):
                self.assertEqual(ctype(value).value, value)

    def test_signed_byte_pointer(self):
        element = typeinfo.TYPEDESC()
        element.vt = automation.VT_I1
        desc = typeinfo.TYPEDESC()
        desc.vt = automation.VT_PTR
        desc._.lptdesc = ctypes.pointer(element)
        ctype = self.make_ctype(desc)
        value = ctypes.c_byte(-7)
        pointer = ctypes.cast(ctypes.pointer(value), ctype)
        self.assertEqual(pointer.contents.value, -7)

    def test_unsigned_byte(self):
        desc = typeinfo.TYPEDESC()
        desc.vt = automation.VT_UI1
        ctype = self.make_ctype(desc)
        for value in (0, 42, 255):
            with self.subTest(value=value):
                self.assertEqual(ctype(value).value, value)

    def test_narrow_string(self):
        desc = typeinfo.TYPEDESC()
        desc.vt = automation.VT_LPSTR
        ctype = self.make_ctype(desc)
        self.assertEqual(ctype(b"hello").value, b"hello")


if __name__ == "__main__":
    unittest.main()
