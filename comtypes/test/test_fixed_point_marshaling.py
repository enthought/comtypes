"""Unit tests for fixed-point number (VT_CY / VT_DECIMAL) marshaling helpers.

These tests do not require a COM server and test the tagCY/tagDEC conversion
methods directly.
"""

import unittest
from decimal import Decimal

from comtypes.automation import tagCY, tagDEC


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


if __name__ == "__main__":
    unittest.main()
