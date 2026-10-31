"""Unit tests for fixed-point number (VT_CY / VT_DECIMAL) marshaling helpers.

These tests do not require a COM server and test the tagCY/tagDEC conversion
methods directly.
"""

import unittest

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


if __name__ == "__main__":
    unittest.main()
