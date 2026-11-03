"""Ownership and element-type tests for the native string SAFEARRAY double."""

import itertools
import sys
import unittest
from _ctypes import COMError

from comtypes import BSTR, CLSCTX_LOCAL_SERVER, hresult
from comtypes.automation import VARIANT, _midlSAFEARRAY
from comtypes.client import CreateObject, GetModule

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True


def _create_tester() -> "ComtypesCppTestSrvLib.IStringSafearrayInsideVariantTest":
    return CreateObject(
        "Comtypes.CoStringSafearrayInsideVariantTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IStringSafearrayInsideVariantTest,
    )


_BstrArrType = _midlSAFEARRAY(BSTR)
_VariantArrType = _midlSAFEARRAY(VARIANT)

PY_VER = "Python {0}.{1}.{2}".format(*sys.version_info[:3])
# It seems to be the same root cause we discussed in https://github.com/enthought/comtypes/issues/212
IS_SKIP_PY_VER = sys.version_info[:2] == (
    (3, 9)
    or (sys.version_info[:2] == (3, 10) and sys.version_info < (3, 10, 10))
    or (sys.version_info[:2] == (3, 11) and sys.version_info < (3, 11, 2))
)


@unittest.skipIf(IS_SKIP_PY_VER, f"This fails in {PY_VER}.")
@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestVariantArrayToBstrArray(unittest.TestCase):
    def test_returns_bstr_array(self):
        tester = _create_tester()
        for factory, values in itertools.product(
            (_VariantArrType.create, _VariantArrType.from_param),
            (("foo", "bar"), ["foo", "bar"]),
        ):
            with self.subTest(factory=factory.__name__, values=values):
                self.assertEqual(
                    tester.VariantArrayToBstrArray(factory(values)), ("foo", "bar")
                )

    def test_raises_element_type_mismatch(self):
        tester = _create_tester()
        pa = _BstrArrType.from_param(["foo", "bar"])
        with self.assertRaises(COMError) as cm:
            tester.VariantArrayToBstrArray(pa)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)


@unittest.skipIf(IS_SKIP_PY_VER, f"This fails in {PY_VER}.")
@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestBstrArrayToVariantArray(unittest.TestCase):
    def test_returns_variant_array(self):
        tester = _create_tester()
        for factory, values in itertools.product(
            (_BstrArrType.create, _BstrArrType.from_param),
            (("foo", "bar"), ["foo", "bar"]),
        ):
            with self.subTest(factory=factory.__name__, values=values):
                self.assertEqual(
                    tester.BstrArrayToVariantArray(factory(values)), ("foo", "bar")
                )

    def test_raises_element_type_mismatch(self):
        tester = _create_tester()
        pa = _VariantArrType.from_param(["foo", "bar"])
        with self.assertRaises(COMError) as cm:
            tester.BstrArrayToVariantArray(pa)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)


@unittest.skipIf(IS_SKIP_PY_VER, f"This fails in {PY_VER}.")
@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestRepeatVariantArray(unittest.TestCase):
    def test_returns_repeated_variant_array(self):
        tester = _create_tester()
        for factory, values in itertools.product(
            (_VariantArrType.create, _VariantArrType.from_param),
            (("foo", "bar"), ["foo", "bar"]),
        ):
            for repeat in range(4):
                with self.subTest(
                    factory=factory.__name__, values=values, repeat=repeat
                ):
                    self.assertEqual(
                        tester.RepeatVariantArray(factory(values), repeat),
                        ("foo", "bar") * repeat,
                    )

    def test_raises_element_type_mismatch(self):
        tester = _create_tester()
        pa = _BstrArrType.from_param(["foo", "bar"])
        with self.assertRaises(COMError) as cm:
            tester.RepeatVariantArray(pa, 1)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)


@unittest.skipIf(IS_SKIP_PY_VER, f"This fails in {PY_VER}.")
@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestRepeatBstrArray(unittest.TestCase):
    def test_returns_repeated_bstr_array(self):
        tester = _create_tester()
        for factory, values in itertools.product(
            (_BstrArrType.create, _BstrArrType.from_param),
            (("foo", "bar"), ["foo", "bar"]),
        ):
            for repeat in range(4):
                with self.subTest(
                    factory=factory.__name__, values=values, repeat=repeat
                ):
                    self.assertEqual(
                        tester.RepeatBstrArray(factory(values), repeat),
                        ("foo", "bar") * repeat,
                    )

    def test_raises_element_type_mismatch(self):
        tester = _create_tester()
        pa = _VariantArrType.from_param(["foo", "bar"])
        with self.assertRaises(COMError) as cm:
            tester.RepeatBstrArray(pa, 1)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)


if __name__ == "__main__":
    unittest.main()
