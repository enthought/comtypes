"""Fixed point number (VT_DECIMAL, VT_CY) round-trip tests via the native C++ COM test server."""

import decimal
import unittest
from _ctypes import COMError
from ctypes import POINTER, byref, pointer, sizeof

from comtypes import CLSCTX_LOCAL_SERVER, hresult
from comtypes.automation import VARIANT, VT_CY, VT_DECIMAL, VT_NULL, tagCY, tagDEC
from comtypes.client import CreateObject, GetModule

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import (
        CoFixedPointNumberTest,
        IFixedPointNumberTest,
    )

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True


_CY_BITS = sizeof(tagCY) * 8
_LLONG_MIN = -(1 << (_CY_BITS - 1))
_LLONG_MAX = (1 << (_CY_BITS - 1)) - 1
_CY_SCALE = decimal.Decimal("10000")
_CY_MIN = decimal.Decimal(_LLONG_MIN) / _CY_SCALE
_CY_MAX = decimal.Decimal(_LLONG_MAX) / _CY_SCALE


def _create_server() -> "IFixedPointNumberTest":
    return CreateObject(
        CoFixedPointNumberTest,
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=IFixedPointNumberTest,
    )


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestToDecVariant(unittest.TestCase):
    """Tests for ToDecVariant: string -> VT_DECIMAL VARIANT (returns Decimal directly)."""

    def setUp(self):
        self.server = _create_server()

    def test_positive_integer(self):
        result = self.server.ToDecVariant("12345")
        self.assertEqual(result, decimal.Decimal("12345"))

    def test_negative_integer(self):
        result = self.server.ToDecVariant("-12345")
        self.assertEqual(result, decimal.Decimal("-12345"))

    def test_positive_decimal(self):
        result = self.server.ToDecVariant("123.45")
        self.assertEqual(result, decimal.Decimal("123.45"))

    def test_negative_decimal(self):
        result = self.server.ToDecVariant("-123.45")
        self.assertEqual(result, decimal.Decimal("-123.45"))

    def test_zero(self):
        result = self.server.ToDecVariant("0")
        self.assertEqual(result, decimal.Decimal("0"))

    def test_zero_with_decimal_point(self):
        result = self.server.ToDecVariant("0.0")
        self.assertEqual(result, decimal.Decimal("0.0"))

    def test_small_decimal(self):
        result = self.server.ToDecVariant("0.0001")
        self.assertEqual(result, decimal.Decimal("0.0001"))

    def test_many_decimal_places(self):
        result = self.server.ToDecVariant("123.45678901234567890123456789")
        self.assertEqual(result, decimal.Decimal("123.45678901234567890123456789"))

    def test_scientific_notation(self):
        result = self.server.ToDecVariant("1.23E+2")
        self.assertEqual(result, decimal.Decimal("123"))

    def test_negative_scientific_notation(self):
        result = self.server.ToDecVariant("-1.23E-2")
        self.assertEqual(result, decimal.Decimal("-0.0123"))


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestToCyVariant(unittest.TestCase):
    """Tests for ToCyVariant: string -> VT_CY VARIANT (returns Decimal directly)."""

    def setUp(self):
        self.server = _create_server()

    def test_positive_integer(self):
        result = self.server.ToCyVariant("12345")
        self.assertEqual(result, decimal.Decimal("12345"))

    def test_negative_integer(self):
        result = self.server.ToCyVariant("-12345")
        self.assertEqual(result, decimal.Decimal("-12345"))

    def test_positive_decimal(self):
        result = self.server.ToCyVariant("123.45")
        self.assertEqual(result, decimal.Decimal("123.45"))

    def test_negative_decimal(self):
        result = self.server.ToCyVariant("-123.45")
        self.assertEqual(result, decimal.Decimal("-123.45"))

    def test_zero(self):
        result = self.server.ToCyVariant("0")
        self.assertEqual(result, decimal.Decimal("0"))

    def test_four_decimal_places(self):
        result = self.server.ToCyVariant("123.4567")
        self.assertEqual(result, decimal.Decimal("123.4567"))

    def test_more_than_four_decimal_places_rounded(self):
        result = self.server.ToCyVariant("123.45678")
        self.assertEqual(result, decimal.Decimal("123.4568"))

    def test_max_cy_value(self):
        result = self.server.ToCyVariant(str(_CY_MAX))
        self.assertEqual(result, _CY_MAX)

    def test_min_cy_value(self):
        result = self.server.ToCyVariant(str(_CY_MIN))
        self.assertEqual(result, _CY_MIN)

    def test_above_max_cy_value_raises_overflow(self):
        """Value exceeding CY max should raise DISP_E_OVERFLOW."""
        # CY max is 922337203685477.5807 (2^63-1 / 10000)
        # Adding 1 to the integer part exceeds the range
        over_max = str(_CY_MAX + decimal.Decimal("1"))
        with self.assertRaises(COMError) as cm:
            self.server.ToCyVariant(over_max)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_OVERFLOW)

    def test_below_min_cy_value_raises_overflow(self):
        """Value below CY min should raise DISP_E_OVERFLOW."""
        # CY min is -922337203685477.5808 (-2^63 / 10000)
        # Subtracting 1 from the integer part exceeds the range
        under_min = str(_CY_MIN - decimal.Decimal("1"))
        with self.assertRaises(COMError) as cm:
            self.server.ToCyVariant(under_min)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_OVERFLOW)


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestDecomposeCyVariant(unittest.TestCase):
    """Tests for DecomposeCyVariant: VT_CY VARIANT -> int64."""

    def setUp(self):
        self.server = _create_server()

    def test_positive_value(self):
        cy_var = VARIANT(tagCY(int64=123456789))
        result = self.server.DecomposeCyVariant(cy_var)
        self.assertEqual(result, 123456789)

    def test_negative_value(self):
        cy_var = VARIANT(tagCY(int64=-123456789))
        result = self.server.DecomposeCyVariant(cy_var)
        self.assertEqual(result, -123456789)

    def test_zero(self):
        cy_var = VARIANT(tagCY(int64=0))
        result = self.server.DecomposeCyVariant(cy_var)
        self.assertEqual(result, 0)

    def test_max_int64(self):
        cy_var = VARIANT(tagCY(int64=_LLONG_MAX))
        result = self.server.DecomposeCyVariant(cy_var)
        self.assertEqual(result, _LLONG_MAX)

    def test_min_int64(self):
        cy_var = VARIANT(tagCY(int64=_LLONG_MIN))
        result = self.server.DecomposeCyVariant(cy_var)
        self.assertEqual(result, _LLONG_MIN)

    def test_raises_type_mismatch_for_non_cy(self):
        for value in [
            VARIANT(123),
            VARIANT("hello"),
            VARIANT(),
            VARIANT(123.45),
            VARIANT(tagDEC(scale=0, sign=0, Hi32=0, Lo64=0)),
        ]:
            with self.subTest(value=value):
                with self.assertRaises(COMError) as cm:
                    self.server.DecomposeCyVariant(value)
                self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestDecomposeDecVariant(unittest.TestCase):
    """Tests for DecomposeDecVariant: VT_DECIMAL VARIANT -> (sign, digits, exponent)."""

    def setUp(self):
        self.server = _create_server()

    def test_positive_integer(self):
        var = VARIANT(tagDEC(scale=0, sign=0, Hi32=0, Lo64=12345))
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 0)
        self.assertEqual(list(digits), [1, 2, 3, 4, 5])
        self.assertEqual(exponent, 0)

    def test_negative_integer(self):
        var = VARIANT(tagDEC(scale=0, sign=0x80, Hi32=0, Lo64=12345))
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 1)
        self.assertEqual(list(digits), [1, 2, 3, 4, 5])
        self.assertEqual(exponent, 0)

    def test_positive_decimal(self):
        var = VARIANT(tagDEC(scale=2, sign=0, Hi32=0, Lo64=12345))
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 0)
        self.assertEqual(list(digits), [1, 2, 3, 4, 5])
        self.assertEqual(exponent, -2)

    def test_negative_decimal(self):
        var = VARIANT(tagDEC(scale=2, sign=0x80, Hi32=0, Lo64=12345))
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 1)
        self.assertEqual(list(digits), [1, 2, 3, 4, 5])
        self.assertEqual(exponent, -2)

    def test_zero(self):
        var = VARIANT(tagDEC(scale=0, sign=0, Hi32=0, Lo64=0))
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 0)
        self.assertEqual(list(digits), [0])
        self.assertEqual(exponent, 0)

    def test_small_decimal(self):
        var = VARIANT(tagDEC(scale=4, sign=0, Hi32=0, Lo64=1))
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 0)
        self.assertEqual(list(digits), [0, 0, 0, 0, 1])
        self.assertEqual(exponent, -4)

    def test_many_decimal_places(self):
        var = VARIANT(
            tagDEC(scale=26, sign=0, Hi32=669260594, Lo64=5097733592125636885)
        )
        sign, digits, exponent = self.server.DecomposeDecVariant(var)
        self.assertEqual(sign, 0)
        self.assertEqual(list(digits), (list(range(10)) * 3)[1:])
        self.assertEqual(exponent, -26)

    def test_raises_type_mismatch_for_non_decimal(self):
        for value in [
            VARIANT(123),
            VARIANT("hello"),
            VARIANT(),
            VARIANT(123.45),
            VARIANT(tagCY(int64=0)),
        ]:
            with self.subTest(value=value):
                with self.assertRaises(COMError) as cm:
                    self.server.DecomposeDecVariant(value)
                self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestInitCyVariantWith3dot14(unittest.TestCase):
    def setUp(self):
        self.server = _create_server()

    def test_empty_variant_becomes_decimal(self):
        """VARIANT() passed by value: return value is Decimal.

        In comtypes, VARIANT passed by value is passed as [in, out] VARIANT*.
        The server modifies the VARIANT in place. The return value is the
        modified VARIANT converted to Decimal. The original `var` is cleared
        (becomes VT_NULL) by the return value processing.
        """
        var = VARIANT()
        result = self.server.InitCyVariantWith3dot14(var)
        self.assertIsInstance(result, decimal.Decimal)
        self.assertEqual(result, decimal.Decimal("3.14"))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_incoming_bstr_is_replaced(self):
        """The server releases the incoming content before reflecting the value.

        The original VARIANT is cleared by return value processing (VT_NULL).
        """
        var = VARIANT("not a number")
        result = self.server.InitCyVariantWith3dot14(var)
        self.assertEqual(result, decimal.Decimal("3.14"))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_incoming_decimal_is_replaced(self):
        """The server releases the incoming content before reflecting the value.

        The original VARIANT is cleared by return value processing (VT_NULL).
        """
        var = VARIANT(decimal.Decimal("1.5"))
        result = self.server.InitCyVariantWith3dot14(var)
        self.assertEqual(result, decimal.Decimal("3.14"))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_pointer_to_variant_is_written_through(self):
        """A POINTER(VARIANT) is passed as it is, so the caller's VARIANT is
        filled.
        Return value is the pointer to VARIANT.
        """
        var = VARIANT()
        result = self.server.InitCyVariantWith3dot14(pointer(var))
        # Return value is the pointer (ctypes pointer object)
        self.assertIsInstance(result, POINTER(VARIANT))
        # The pointed-to VARIANT is mutated
        self.assertEqual(result.contents.vt, VT_CY)
        self.assertEqual(result.contents.value, decimal.Decimal("3.14"))
        # Argument is mutated
        self.assertEqual(var.vt, VT_CY)
        self.assertEqual(var.value, decimal.Decimal("3.14"))

    def test_byref_variant_writes_through(self):
        """`byref(VARIANT)` is wrapped into a new VARIANT (VT_BYREF | VT_VARIANT);
        the server writes through to the referenced VARIANT.
        Return value is the outer VARIANT (VT_NULL), argument var is mutated.
        """
        var = VARIANT()
        result = self.server.InitCyVariantWith3dot14(byref(var))
        # Return value is the outer VARIANT (VT_NULL because server doesn't modify it)
        self.assertEqual(result.vt, VT_NULL)
        # Argument var is mutated through the reference
        self.assertEqual(var.vt, VT_CY)
        self.assertEqual(var.value, decimal.Decimal("3.14"))

    def test_byref_cy_is_written_through(self):
        """`byref(tagCY)` becomes a VT_BYREF | VT_CY VARIANT; the server writes
        through it. Return value is the outer VARIANT (VT_NULL).
        """
        cy = tagCY(0)
        result = self.server.InitCyVariantWith3dot14(byref(cy))
        # Return value is the outer VARIANT (VT_NULL)
        self.assertEqual(result.vt, VT_NULL)
        # Argument is mutated
        self.assertEqual(cy.int64, 31400)

    def test_pointer_to_cy_is_written_through(self):
        """`pointer(tagCY)` becomes a VT_BYREF | VT_CY VARIANT; the server
        writes through it. Return value is the outer VARIANT (VT_NULL).
        """
        cy = tagCY(0)
        result = self.server.InitCyVariantWith3dot14(pointer(cy))
        # Return value is the outer VARIANT (VT_NULL) created to wrap the pointer
        self.assertEqual(result.vt, VT_NULL)
        # Argument is mutated
        self.assertEqual(cy.int64, 31400)


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestInitDecVariantWith3dot14(unittest.TestCase):
    def setUp(self):
        self.server = _create_server()

    def test_empty_variant_becomes_decimal(self):
        """VARIANT() passed by value: return value is Decimal.

        In comtypes, VARIANT passed by value is passed as [in, out] VARIANT*.
        The server modifies the VARIANT in place. The return value is the
        modified VARIANT converted to Decimal. The original `var` is cleared
        (becomes VT_NULL) by the return value processing.
        """
        var = VARIANT()
        result = self.server.InitDecVariantWith3dot14(var)
        self.assertIsInstance(result, decimal.Decimal)
        self.assertEqual(result, decimal.Decimal("3.14"))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_scale_is_preserved(self):
        """VT_DECIMAL keeps its scale `(2)`, so the tuple is exactly
        `(0, (3, 1, 4), -2)`.
        """
        var = VARIANT()
        result = self.server.InitDecVariantWith3dot14(var)
        self.assertEqual(result.as_tuple(), decimal.DecimalTuple(0, (3, 1, 4), -2))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_incoming_bstr_is_replaced(self):
        """The server releases the incoming content before reflecting the value.

        The original VARIANT is cleared by return value processing (VT_NULL).
        """
        var = VARIANT("not a number")
        result = self.server.InitDecVariantWith3dot14(var)
        self.assertEqual(result, decimal.Decimal("3.14"))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_incoming_decimal_is_replaced(self):
        """The server releases the incoming content before reflecting the value.

        The original VARIANT is cleared by return value processing (VT_NULL).
        """
        var = VARIANT(decimal.Decimal("1.5"))
        result = self.server.InitDecVariantWith3dot14(var)
        self.assertEqual(result, decimal.Decimal("3.14"))
        # The original var is cleared by return value processing (VT_NULL)
        self.assertEqual(var.vt, VT_NULL)

    def test_pointer_to_variant_is_written_through(self):
        """A POINTER(VARIANT) is passed as it is, so the caller's VARIANT is
        filled.
        Return value is the pointer to VARIANT.
        """
        var = VARIANT()
        result = self.server.InitDecVariantWith3dot14(pointer(var))
        # Return value is the pointer (ctypes pointer object)
        self.assertIsInstance(result, POINTER(VARIANT))
        # The pointed-to VARIANT is mutated
        self.assertEqual(result.contents.vt, VT_DECIMAL)
        self.assertEqual(result.contents.value, decimal.Decimal("3.14"))
        # Argument is mutated
        self.assertEqual(var.vt, VT_DECIMAL)
        self.assertEqual(var.value, decimal.Decimal("3.14"))

    def test_byref_variant_writes_through(self):
        """`byref(VARIANT)` is wrapped into a new VARIANT (VT_BYREF | VT_VARIANT);
        the server writes through to the referenced VARIANT.
        Return value is the outer VARIANT (VT_NULL), argument var is mutated.
        """
        var = VARIANT()
        result = self.server.InitDecVariantWith3dot14(byref(var))
        # Return value is the outer VARIANT (VT_NULL because server doesn't modify it)
        self.assertEqual(result.vt, VT_NULL)
        # Argument var is mutated through the reference
        self.assertEqual(var.vt, VT_DECIMAL)
        self.assertEqual(var.value, decimal.Decimal("3.14"))

    def test_byref_decimal_is_written_through(self):
        """`byref(tagDEC)` becomes a VT_BYREF | VT_DECIMAL VARIANT; the server
        writes through it. Return value is the outer VARIANT (VT_NULL).
        """
        dec = tagDEC()
        result = self.server.InitDecVariantWith3dot14(byref(dec))
        # Return value is the outer VARIANT (VT_NULL)
        self.assertEqual(result.vt, VT_NULL)
        # Argument is mutated
        self.assertEqual(dec.as_decimal(), decimal.Decimal("3.14"))
        self.assertEqual((dec.scale, dec.sign, dec.Hi32, dec.Lo64), (2, 0, 0, 314))

    def test_pointer_to_decimal_is_written_through(self):
        """`pointer(tagDEC)` becomes a VT_BYREF | VT_DECIMAL VARIANT; the server
        writes through it. Return value is the outer VARIANT (VT_NULL).
        """
        dec = tagDEC()
        result = self.server.InitDecVariantWith3dot14(pointer(dec))
        # Return value is the outer VARIANT (VT_NULL) created to wrap the pointer
        self.assertEqual(result.vt, VT_NULL)
        # Argument is mutated
        self.assertEqual(dec.as_decimal(), decimal.Decimal("3.14"))
        self.assertEqual((dec.scale, dec.sign, dec.Hi32, dec.Lo64), (2, 0, 0, 314))


if __name__ == "__main__":
    unittest.main()
