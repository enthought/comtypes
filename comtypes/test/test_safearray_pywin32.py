import unittest
from ctypes import POINTER, PyDLL, byref, py_object
from ctypes.wintypes import BOOL

from comtypes.automation import VARIANT


raise unittest.SkipTest("This depends on 'pywin32'.")

import pythoncom
# pywin32 not installed...

# pywin32 is available.  The pythoncom dll contains two handy
# exported functions that allow to create a VARIANT from a Python
# object, also a function that unpacks a VARIANT into a Python
# object.
#
# This allows us to create und unpack SAFEARRAY instances
# contained in VARIANTs, and check for consistency with the
# comtypes code.

_dll = PyDLL(pythoncom.__file__)

# c:/sf/pywin32/com/win32com/src/oleargs.cpp 213
# PyObject *PyCom_PyObjectFromVariant(const VARIANT *var)
unpack = _dll.PyCom_PyObjectFromVariant
unpack.restype = py_object
unpack.argtypes = (POINTER(VARIANT),)

# c:/sf/pywin32/com/win32com/src/oleargs.cpp 54
# BOOL PyCom_VariantFromPyObject(PyObject *obj, VARIANT *var)
_pack = _dll.PyCom_VariantFromPyObject
_pack.argtypes = py_object, POINTER(VARIANT)
_pack.restype = BOOL


def pack(obj):
    var = VARIANT()
    _pack(obj, byref(var))
    return var


class PyWinTest(unittest.TestCase):
    def test_1dim(self):
        data = (1, 2, 3)
        variant = pack(data)
        self.assertEqual(variant.value, data)
        self.assertEqual(unpack(variant), data)

    def test_2dim(self):
        data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        variant = pack(data)
        self.assertEqual(variant.value, data)
        self.assertEqual(unpack(variant), data)

    def test_3dim(self):
        data = (((1, 2), (3, 4), (5, 6)), ((7, 8), (9, 10), (11, 12)))
        variant = pack(data)
        self.assertEqual(variant.value, data)
        self.assertEqual(unpack(variant), data)

    def test_4dim(self):
        data = (
            (((1, 2), (3, 4)), ((5, 6), (7, 8))),
            (((9, 10), (11, 12)), ((13, 14), (15, 16))),
        )
        variant = pack(data)
        self.assertEqual(variant.value, data)
        self.assertEqual(unpack(variant), data)


if __name__ == "__main__":
    unittest.main()
