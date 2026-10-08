"""Numeric SAFEARRAYs returned by the first-party out-of-process C++ server."""

import unittest

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.safearray import safearray_as_ndarray

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import (
        CoNumericSafearrayTest,
        INumericSafearrayTest,
    )

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True

try:
    import numpy

    IMPORT_NUMPY_FAILED = False
except ImportError:
    IMPORT_NUMPY_FAILED = True


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class NumericSafearrayTest(unittest.TestCase):
    def setUp(self):
        self.server = CreateObject(
            CoNumericSafearrayTest,
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=INumericSafearrayTest,
        )
        # Fixed expected values come from the C++ fixture, not Python input.
        self.cases_1d = [
            (self.server.GetUint8Array, "uint8", (0, 42, 255, 0, 42, 255)),
            (self.server.GetInt16Array, "int16", (-7, 0, 42, -7, 0, 42)),
            (self.server.GetUint16Array, "uint16", (0, 42, 65535, 0, 42, 65535)),
            (self.server.GetInt32Array, "int32", (-7, 0, 42, -7, 0, 42)),
            (
                self.server.GetUint32Array,
                "uint32",
                (0, 42, 4294967295, 0, 42, 4294967295),
            ),
            (self.server.GetInt64Array, "int64", (-7, 0, 42, -7, 0, 42)),
            (
                self.server.GetUint64Array,
                "uint64",
                (0, 42, 4294967297, 0, 42, 4294967297),
            ),
            (
                self.server.GetFloat32Array,
                "float32",
                (1.25, -2.5, 3.75, 1.25, -2.5, 3.75),
            ),
            (
                self.server.GetFloat64Array,
                "float64",
                (1.25, -2.5, 3.75, 1.25, -2.5, 3.75),
            ),
        ]
        # SAFEARRAY storage is column-major; the fixture has two rows.
        self.cases_2d = [
            (self.server.GetUint8Array, "uint8", ((0, 255, 42), (42, 0, 255))),
            (self.server.GetInt16Array, "int16", ((-7, 42, 0), (0, -7, 42))),
            (self.server.GetUint16Array, "uint16", ((0, 65535, 42), (42, 0, 65535))),
            (self.server.GetInt32Array, "int32", ((-7, 42, 0), (0, -7, 42))),
            (
                self.server.GetUint32Array,
                "uint32",
                ((0, 4294967295, 42), (42, 0, 4294967295)),
            ),
            (self.server.GetInt64Array, "int64", ((-7, 42, 0), (0, -7, 42))),
            (
                self.server.GetUint64Array,
                "uint64",
                ((0, 4294967297, 42), (42, 0, 4294967297)),
            ),
            (
                self.server.GetFloat32Array,
                "float32",
                ((1.25, 3.75, -2.5), (-2.5, 1.25, 3.75)),
            ),
            (
                self.server.GetFloat64Array,
                "float64",
                ((1.25, 3.75, -2.5), (-2.5, 1.25, 3.75)),
            ),
        ]

    def test_tuple_values_1d(self):
        for method, dtype, expected in self.cases_1d:
            with self.subTest(dtype=dtype):
                result = method(1, 6)
                self.assertIsInstance(result, tuple)
                self.assertEqual(result, expected)

    @unittest.skipIf(IMPORT_NUMPY_FAILED, "Requires NumPy.")
    def test_ndarray_values_1d(self):
        for method, dtype, expected in self.cases_1d:
            with self.subTest(dtype=dtype):
                with safearray_as_ndarray:
                    result = method(1, 6)
                self.assertIsInstance(result, numpy.ndarray)
                self.assertEqual(result.shape, (6,))
                numpy.testing.assert_array_equal(result, expected)
                self.assertEqual(result.dtype, numpy.dtype(dtype))
                # COM's returned SAFEARRAY is already destroyed at this point.
                # A second call and mutation must not affect the first copy.
                with safearray_as_ndarray:
                    other = method(1, 6)
                other.flat[0] = 99
                numpy.testing.assert_array_equal(result, expected)
                self.assertEqual(method(1, 6), expected)

    def test_tuple_values_2d(self):
        for method, dtype, expected in self.cases_2d:
            with self.subTest(dtype=dtype):
                result = method(2, 6)
                self.assertIsInstance(result, tuple)
                self.assertEqual(result, expected)

    @unittest.skipIf(IMPORT_NUMPY_FAILED, "Requires NumPy.")
    def test_ndarray_values_2d(self):
        for method, dtype, expected in self.cases_2d:
            with self.subTest(dtype=dtype):
                with safearray_as_ndarray:
                    result = method(2, 6)
                self.assertIsInstance(result, numpy.ndarray)
                self.assertEqual(result.shape, (2, 3))
                numpy.testing.assert_array_equal(result, expected)
                self.assertEqual(result.dtype, numpy.dtype(dtype))
                # COM's returned SAFEARRAY is already destroyed at this point.
                # A second call and mutation must not affect the first copy.
                with safearray_as_ndarray:
                    other = method(2, 6)
                other.flat[0] = 99
                numpy.testing.assert_array_equal(result, expected)
                self.assertEqual(method(2, 6), expected)

    def test_int8_codegen_1d(self):
        # VT_I1 is currently parsed as c_char (see #935), whose pointer slice
        # returns bytes. Check the real COM/code-generator path and raw bytes;
        # signed int8/NumPy dtype assertions await the separate c_char/c_byte fix.
        result = self.server.GetInt8Array(1, 6)
        self.assertEqual(result, (249, 0, 42, 249, 0, 42))

    def test_int8_codegen_2d(self):
        # As above, 249 is the raw byte for -7, not a signed dtype assertion.
        result = self.server.GetInt8Array(2, 6)
        self.assertEqual(result, ((249, 42, 0), (0, 249, 42)))


if __name__ == "__main__":
    unittest.main()
