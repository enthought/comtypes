"""Numeric SAFEARRAYs returned by the first-party out-of-process C++ server."""

import unittest

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.safearray import safearray_as_ndarray

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import (
        CoComtypesDispSafearrayParamTest,
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


# These are fixed expectations for the C++ fixture, not values sent to COM.
CASES = [
    ("GetUint8Array", "uint8", (0, 42, 255)),
    ("GetInt16Array", "int16", (-7, 0, 42)),
    ("GetUint16Array", "uint16", (0, 42, 65535)),
    ("GetInt32Array", "int32", (-7, 0, 42)),
    ("GetUint32Array", "uint32", (0, 42, 4294967295)),
    ("GetInt64Array", "int64", (-7, 0, 42)),
    ("GetUint64Array", "uint64", (0, 42, 4294967297)),
    ("GetFloat32Array", "float32", (1.25, -2.5, 3.75)),
    ("GetFloat64Array", "float64", (1.25, -2.5, 3.75)),
]


def expected_values(values, dimensions):
    a, b, c = values
    if dimensions == 1:
        return (a, b, c, a, b, c)
    # SAFEARRAY storage is column-major; the fixture has two rows.
    return ((a, c, b), (b, a, c))


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class NumericSafearrayTest(unittest.TestCase):
    def setUp(self):
        self.server = CreateObject(
            CoComtypesDispSafearrayParamTest,
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=INumericSafearrayTest,
        )

    def test_tuple_values(self):
        for method, dtype, values in CASES:
            for dimensions in (1, 2):
                with self.subTest(dtype=dtype, dimensions=dimensions):
                    result = getattr(self.server, method)(dimensions, 6)
                    self.assertIsInstance(result, tuple)
                    self.assertEqual(result, expected_values(values, dimensions))

    @unittest.skipIf(IMPORT_NUMPY_FAILED, "Requires NumPy.")
    def test_ndarray_values(self):
        # This compatibility test also runs against the pre-fix production code.
        for method, dtype, values in CASES:
            for dimensions in (1, 2):
                with self.subTest(dtype=dtype, dimensions=dimensions):
                    with safearray_as_ndarray:
                        result = getattr(self.server, method)(dimensions, 6)
                    self.assertIsInstance(result, numpy.ndarray)
                    self.assertEqual(result.shape, (6,) if dimensions == 1 else (2, 3))
                    expected = expected_values(values, dimensions)
                    expected_list = (
                        list(expected)
                        if dimensions == 1
                        else [list(row) for row in expected]
                    )
                    self.assertEqual(result.tolist(), expected_list)
                    # COM's returned SAFEARRAY is already destroyed at this point.
                    # A second call and mutation must not affect the first copy.
                    with safearray_as_ndarray:
                        other = getattr(self.server, method)(dimensions, 6)
                    other.flat[0] = 99
                    self.assertEqual(result.tolist(), expected_list)
                    self.assertEqual(
                        getattr(self.server, method)(dimensions, 6), expected
                    )

    @unittest.skipIf(IMPORT_NUMPY_FAILED, "Requires NumPy.")
    def test_ndarray_dtype(self):
        for method, dtype, values in CASES:
            for dimensions in (1, 2):
                with self.subTest(dtype=dtype, dimensions=dimensions):
                    with safearray_as_ndarray:
                        result = getattr(self.server, method)(dimensions, 6)
                    self.assertEqual(result.dtype, numpy.dtype(dtype))


if __name__ == "__main__":
    unittest.main()
