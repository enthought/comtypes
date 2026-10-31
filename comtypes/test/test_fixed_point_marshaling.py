"""Unit tests for fixed-point number (VT_CY / VT_DECIMAL) marshaling helpers.

These tests do not require a COM server and test the tagCY/tagDEC conversion
methods directly.
"""

import decimal
import unittest
from decimal import Decimal

from comtypes.automation import (
    _CY_MAX,
    _CY_MIN,
    _DECIMAL_MAX_MANTISSA,
    VARIANT,
    tagCY,
    tagDEC,
)


class TestTagCY(unittest.TestCase):
    """Tests for tagCY conversion methods."""

    def test_repr(self):
        """Test tagCY.__repr__()."""
        test_cases = [
            (tagCY(int64=1234500), "tagCY(int64=1234500)"),
            (tagCY(int64=-1234500), "tagCY(int64=-1234500)"),
            (tagCY(int64=0), "tagCY(int64=0)"),
        ]
        for cy, expected in test_cases:
            with self.subTest(cy=cy):
                self.assertEqual(repr(cy), expected)

    def test_as_decimal(self):
        """Test tagCY.as_decimal()."""
        cy = tagCY(int64=1234500)
        self.assertEqual(cy.as_decimal(), Decimal("123.45"))

    def test_as_decimal_negative(self):
        """Test tagCY.as_decimal() with negative value."""
        cy = tagCY(int64=-1234500)
        self.assertEqual(cy.as_decimal(), Decimal("-123.45"))

    def test_from_decimal_exact(self):
        """Test tagCY.from_decimal() with exact 4-decimal values."""
        test_cases = [
            (Decimal("123.45"), 1234500),
            (Decimal("123.4500"), 1234500),
            (Decimal("123.450000"), 1234500),
            (Decimal("3.14"), 31400),
            (Decimal("1.2300"), 12300),
            (Decimal("0"), 0),
            (Decimal("-3.14"), -31400),
            (Decimal("-1.2300"), -12300),
            (Decimal("-0"), 0),
        ]
        for val, int64 in test_cases:
            with self.subTest(val=val):
                cy = tagCY.from_decimal(val)
                self.assertEqual(cy.int64, int64)

    def test_from_decimal_inexact_raises(self):
        """Test tagCY.from_decimal() raises OverflowError for values with >4 decimal places.

        These values would be silently rounded if we allowed them (e.g., 123.45678
        would become 123.4567). Instead, we require the caller to explicitly round
        to 4 decimal places before passing the value, giving them control over
        the rounding behavior.
        """
        test_cases = [
            Decimal("123.45678"),  # would round to 123.4567
            Decimal("0.00005"),  # would round to 0
            Decimal("0.00015"),  # would round to 0.0001
            Decimal("1.23456"),  # would round to 1.2345
        ]
        for val in test_cases:
            with self.subTest(val=val):
                with self.assertRaises(OverflowError) as cm:
                    tagCY.from_decimal(val)
                self.assertIn("more than 4 decimal places", str(cm.exception))
                self.assertIn("Round explicitly before calling", str(cm.exception))

    def test_from_decimal_overflow(self):
        """Test tagCY.from_decimal() raises OverflowError for out-of-range values."""
        test_cases = [
            _CY_MAX + Decimal("0.0001"),  # above CY_MAX
            _CY_MIN - Decimal("0.0001"),  # below CY_MIN
            Decimal("1E+30"),  # very large value
        ]
        for val in test_cases:
            with self.subTest(val=val):
                with self.assertRaises(OverflowError) as cm:
                    tagCY.from_decimal(val)
                self.assertIn("out of range for VT_CY", str(cm.exception))

    def test_round_trip(self):
        """Test round-trip Decimal -> tagCY -> Decimal."""
        test_cases = [
            Decimal("0"),
            Decimal("123.45"),
            Decimal("-123.45"),
            Decimal("0.0001"),
            Decimal("-0.0001"),
            Decimal("922337203685477.5807"),  # CY_MAX
            Decimal("-922337203685477.5808"),  # CY_MIN
        ]
        for val in test_cases:
            with self.subTest(val=val):
                cy = tagCY.from_decimal(val)
                self.assertEqual(cy.as_decimal(), val)

    def test_tagCY_from_decimal_context_independent(self):
        """Test tagCY.from_decimal is context-independent."""
        with decimal.localcontext(decimal.Context(prec=5)):
            cy = tagCY.from_decimal(Decimal("123.4567"))
            self.assertEqual(cy.int64, 1234567)

    def test_tagCY_from_decimal_traps_independent(self):
        """Test tagCY.from_decimal is not affected by context traps."""
        with decimal.localcontext(
            decimal.Context(traps=[decimal.Inexact, decimal.Rounded])
        ):
            # Should not raise
            cy = tagCY.from_decimal(Decimal("123.4567"))
            self.assertEqual(cy.int64, 1234567)

    def test_tagCY_from_decimal_rejects_nan(self):
        """Test tagCY.from_decimal rejects NaN."""
        with self.assertRaises(ValueError) as cm:
            tagCY.from_decimal(Decimal("NaN"))
        self.assertIn("NaN cannot be converted", str(cm.exception))

    def test_tagCY_from_decimal_rejects_infinity(self):
        """Test tagCY.from_decimal rejects Infinity."""
        with self.assertRaises(OverflowError) as cm:
            tagCY.from_decimal(Decimal("Infinity"))
        self.assertIn("Infinity cannot be converted", str(cm.exception))

    def test_tagCY_from_decimal_rejects_non_decimal(self):
        """Test tagCY.from_decimal rejects non-decimal."""
        with self.assertRaises(TypeError) as cm:
            tagCY.from_decimal(object())
        self.assertIn("only Decimal is accepted", str(cm.exception))


class TestTagDEC(unittest.TestCase):
    """Tests for tagDEC conversion methods."""

    def test_repr(self):
        """Test tagDEC.__repr__()."""
        test_cases = [
            (
                tagDEC(scale=2, sign=0, Hi32=0, Lo64=314),
                "tagDEC(scale=2, sign=0, Hi32=0, Lo64=314)",
            ),
            (
                tagDEC(scale=4, sign=0x80, Hi32=0, Lo64=987654321),
                "tagDEC(scale=4, sign=128, Hi32=0, Lo64=987654321)",
            ),
            (
                tagDEC(scale=0, sign=0, Hi32=0, Lo64=0),
                "tagDEC(scale=0, sign=0, Hi32=0, Lo64=0)",
            ),
        ]
        for dec, expected in test_cases:
            with self.subTest(dec=dec):
                self.assertEqual(repr(dec), expected)

    def test_as_decimal(self):
        """Test tagDEC.as_decimal()."""
        dec = tagDEC(scale=2, sign=0, Hi32=0, Lo64=314)
        self.assertEqual(dec.as_decimal(), Decimal("3.14"))

    def test_as_decimal_negative(self):
        """Test tagDEC.as_decimal() with negative value."""
        dec = tagDEC(scale=2, sign=0x80, Hi32=0, Lo64=314)
        self.assertEqual(dec.as_decimal(), Decimal("-3.14"))

    def test_from_decimal_exact(self):
        """Test tagDEC.from_decimal() with exact values."""
        test_cases = [
            (Decimal("3.14"), Decimal("3.14"), 2, 0),
            (Decimal("1.2300"), Decimal("1.2300"), 4, 0),
            (Decimal("0"), Decimal("0"), 0, 0),
            (Decimal("-3.14"), Decimal("-3.14"), 2, 0x80),
            (Decimal("-1.2300"), Decimal("-1.2300"), 4, 0x80),
            (Decimal("-0"), Decimal("-0"), 0, 0x80),
        ]
        for val, expected, scale, sign in test_cases:
            with self.subTest(val=val):
                dec = tagDEC.from_decimal(val)
                self.assertEqual(dec.as_decimal(), expected)
                self.assertEqual(dec.scale, scale)
                self.assertEqual(dec.sign, sign)

    def test_from_decimal_max_mantissa(self):
        """Test tagDEC.from_decimal() with max 96-bit mantissa."""
        # 2**96 - 1 is the max mantissa
        val = Decimal(_DECIMAL_MAX_MANTISSA)
        dec = tagDEC.from_decimal(val)
        mantissa = (dec.Hi32 << 64) + dec.Lo64
        self.assertEqual(mantissa, _DECIMAL_MAX_MANTISSA)

    def test_from_decimal_acceptable_many_decimal_places(self):
        val = Decimal("1.2345678901234567890123456789")  # 28 decimal places - OK
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.scale, 28)

    def test_from_decimal_overflow(self):
        """Test tagDEC.from_decimal() raises OverflowError for too-large values."""
        # Value exceeding 96-bit mantissa at the required scale
        val = Decimal("12345678901234567890123456789.0123456789")
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("exceeds 96-bit mantissa", str(cm.exception))
        self.assertIn("Reduce precision explicitly", str(cm.exception))

    def test_from_decimal_too_many_decimal_places(self):
        """Test tagDEC.from_decimal() raises OverflowError for >28 decimal places."""
        # Use a value with 29 non-zero decimal places (no trailing zeros to strip)
        val = Decimal("1.23456789012345678901234567891")
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("more than 28 decimal places", str(cm.exception))
        self.assertIn("Round explicitly before calling", str(cm.exception))

    def test_from_decimal_trailing_zeros(self):
        """Test tagDEC.from_decimal() preserves trailing zeros in scale."""
        # 1.2300 should have scale 4
        dec = tagDEC.from_decimal(Decimal("1.2300"))
        self.assertEqual(dec.scale, 4)
        self.assertEqual(dec.as_decimal(), Decimal("1.2300"))

    def test_from_decimal_trailing_zeros_stripped(self):
        """Test tagDEC.from_decimal() accepts values with many trailing zeros beyond 28 places."""
        # 1.23 with 30 trailing zeros (32 decimal places total) - should be accepted
        # because trailing zeros are stripped until exp >= -28, reducing scale to 28
        val = Decimal("1.23" + "0" * 30)
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.scale, 28)
        self.assertEqual(dec.as_decimal(), Decimal("1.23" + "0" * 28))

    def test_from_decimal_zero_with_many_decimal_places(self):
        """Test tagDEC.from_decimal() accepts zero with many decimal places."""
        # 0E-30 should be accepted (zero is special-cased, scale preserved)
        val = Decimal("0E-30")
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.scale, 28)  # scale preserved from exponent
        self.assertEqual(dec.as_decimal(), Decimal("0E-28"))

    def test_from_decimal_huge_positive_exponent(self):
        """Test tagDEC.from_decimal() rejects huge positive exponent early."""
        # 1E1000000 would cause huge integer allocation if not checked early
        val = Decimal("1E1000000")
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("exceeds 96-bit mantissa", str(cm.exception))
        self.assertIn("Reduce precision explicitly", str(cm.exception))

    def test_from_decimal_negative_zero(self):
        """Test tagDEC.from_decimal() with negative zero."""
        dec = tagDEC.from_decimal(Decimal("-0"))
        self.assertEqual(dec.sign, 0x80)
        self.assertEqual(dec.as_decimal(), Decimal("-0"))

    def test_round_trip(self):
        """Test round-trip Decimal -> tagDEC -> Decimal."""
        test_cases = [
            Decimal("0"),
            Decimal("123.45"),
            Decimal("-123.45"),
            Decimal("1E-28"),
            Decimal("-1E-28"),
            Decimal("1.2300"),
            Decimal("-1.2300"),
            Decimal(str(_DECIMAL_MAX_MANTISSA)),
        ]
        for val in test_cases:
            with self.subTest(val=val):
                dec = tagDEC.from_decimal(val)
                self.assertEqual(dec.as_decimal(), val)

    def test_tagDEC_from_decimal_context_independent(self):
        """Test tagDEC.from_decimal is context-independent."""
        with decimal.localcontext(decimal.Context(prec=5)):
            val = Decimal("123.45678901234567890")
            dec = tagDEC.from_decimal(val)
            # Should not be affected by context precision - round-trip must match
            self.assertEqual(dec.as_decimal(), val)

    def test_tagDEC_from_decimal_traps_independent(self):
        """Test tagDEC.from_decimal is not affected by context traps."""
        with decimal.localcontext(
            decimal.Context(traps=[decimal.Inexact, decimal.Rounded])
        ):
            val = Decimal("123.45678901234567890")
            dec = tagDEC.from_decimal(val)
            # Should not raise and round-trip must match
            self.assertEqual(dec.as_decimal(), val)

    def test_tagDEC_from_decimal_positive_exponent_zero(self):
        """Test tagDEC.from_decimal with positive exponent zero (0E+5)."""
        val = Decimal("0E+5")
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.scale, 0)
        self.assertEqual(dec.as_decimal(), Decimal("0"))

    def test_tagDEC_from_decimal_positive_exponent(self):
        """Test tagDEC.from_decimal with positive exponent (1E+1)."""
        val = Decimal("1E+1")  # = 10
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.scale, 0)
        self.assertEqual(dec.as_decimal(), Decimal("10"))

        # Larger positive exponent
        val = Decimal("123E+2")  # = 12300
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.scale, 0)
        self.assertEqual(dec.as_decimal(), Decimal("12300"))

    def test_tagDEC_from_decimal_rejects_nan(self):
        """Test tagDEC.from_decimal rejects NaN."""
        with self.assertRaises(ValueError) as cm:
            tagDEC.from_decimal(Decimal("NaN"))
        self.assertIn("NaN cannot be converted", str(cm.exception))

    def test_tagDEC_from_decimal_rejects_snan(self):
        """Test tagDEC.from_decimal rejects signaling NaN."""
        with self.assertRaises(ValueError) as cm:
            tagDEC.from_decimal(Decimal("sNaN"))
        self.assertIn("NaN cannot be converted", str(cm.exception))

    def test_tagDEC_from_decimal_rejects_infinity(self):
        """Test tagDEC.from_decimal rejects Infinity."""
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(Decimal("Infinity"))
        self.assertIn("Infinity cannot be converted", str(cm.exception))

    def test_tagDEC_from_decimal_rejects_negative_infinity(self):
        """Test tagDEC.from_decimal rejects -Infinity."""
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(Decimal("-Infinity"))
        self.assertIn("Infinity cannot be converted", str(cm.exception))

    def test_tagDEC_from_decimal_rejects_non_decimal(self):
        """Test tagDEC.from_decimal rejects non-decimal."""
        with self.assertRaises(TypeError) as cm:
            tagDEC.from_decimal(object())
        self.assertIn("only Decimal is accepted", str(cm.exception))

    def test_tagDEC_from_decimal_boundary_values(self):
        """Test tagDEC.from_decimal with boundary values."""
        # 1E+28 (29 digits) - should be accepted (fits in 96-bit mantissa)
        val = Decimal("1E+28")
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.as_decimal(), val)

        # 1E+29 (30 digits) - should be rejected (exceeds 96-bit mantissa)
        val = Decimal("1E+29")
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("exceeds 96-bit mantissa", str(cm.exception))

        # Max mantissa value (2**96 - 1) - should be accepted
        val = Decimal(str(_DECIMAL_MAX_MANTISSA))
        dec = tagDEC.from_decimal(val)
        mantissa = (dec.Hi32 << 64) + dec.Lo64
        self.assertEqual(mantissa, _DECIMAL_MAX_MANTISSA)

        # Max mantissa + 1 - should be rejected
        val = Decimal(str(_DECIMAL_MAX_MANTISSA + 1))
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("exceeds 96-bit mantissa", str(cm.exception))

        # Negative boundary values
        # -1E+28 - should be accepted
        val = Decimal("-1E+28")
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.as_decimal(), val)
        self.assertEqual(dec.sign, 0x80)

        # -1E+29 - should be rejected
        val = Decimal("-1E+29")
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("exceeds 96-bit mantissa", str(cm.exception))

        # -Max mantissa - should be accepted
        val = Decimal(str(-_DECIMAL_MAX_MANTISSA))
        dec = tagDEC.from_decimal(val)
        mantissa = (dec.Hi32 << 64) + dec.Lo64
        self.assertEqual(mantissa, _DECIMAL_MAX_MANTISSA)
        self.assertEqual(dec.sign, 0x80)

        # -Max mantissa - 1 - should be rejected
        val = Decimal(str(-_DECIMAL_MAX_MANTISSA - 1))
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("exceeds 96-bit mantissa", str(cm.exception))

        # 1.5E-27 (28 decimal places, non-zero) - should be accepted
        val = Decimal("1.5E-27")
        dec = tagDEC.from_decimal(val)
        self.assertEqual(dec.as_decimal(), val)

        # 1E-29 (29 decimal places) - should be rejected
        val = Decimal("1E-29")
        with self.assertRaises(OverflowError) as cm:
            tagDEC.from_decimal(val)
        self.assertIn("more than 28 decimal places", str(cm.exception))


class TestCyVariantValueProperty(unittest.TestCase):
    """Tests for VARIANT.value property with VT_CY."""

    def test_value_property_returns_decimal(self):
        with self.assertRaises(TypeError) as cm:
            VARIANT(tagCY(int64=1234567))
        self.assertIn("Cannot put tagCY", str(cm.exception))

    def test_value_property_negative(self):
        with self.assertRaises(TypeError) as cm:
            VARIANT(tagCY(int64=-1234567))
        self.assertIn("Cannot put tagCY", str(cm.exception))

    def test_value_property_zero(self):
        with self.assertRaises(TypeError) as cm:
            VARIANT(tagCY(int64=0))
        self.assertIn("Cannot put tagCY", str(cm.exception))


class TestDecVariantValueProperty(unittest.TestCase):
    """Tests for VARIANT.value property with VT_DECIMAL."""

    def test_value_property_returns_decimal(self):
        with self.assertRaises(TypeError) as cm:
            VARIANT(tagDEC(scale=26, sign=0, Hi32=669260594, Lo64=5097733592125636885))
        self.assertIn("Cannot put tagDEC", str(cm.exception))

    def test_value_property_negative(self):
        with self.assertRaises(TypeError) as cm:
            VARIANT(
                tagDEC(scale=26, sign=0x80, Hi32=669260594, Lo64=5097733592125636885)
            )
        self.assertIn("Cannot put tagDEC", str(cm.exception))

    def test_value_property_zero(self):
        with self.assertRaises(TypeError) as cm:
            VARIANT(tagDEC(scale=0, sign=0, Hi32=0, Lo64=0))
        self.assertIn("Cannot put tagDEC", str(cm.exception))


if __name__ == "__main__":
    unittest.main()
