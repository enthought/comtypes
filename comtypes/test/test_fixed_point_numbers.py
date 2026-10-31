"""Fixed point number (VT_DECIMAL, VT_CY) round-trip tests via the native C++ COM test server."""

import decimal
import unittest
from ctypes import POINTER, byref, pointer, sizeof

from comtypes import CLSCTX_LOCAL_SERVER
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

        NOTE: This test currently expects a KeyError because `tagCY` is not
        registered in `comtypes.automation._ctype_to_vartype`. The mapping
        should include:
            tagCY: VT_CY,
        Once this is added to _ctype_to_vartype, this test should be updated
        to verify the correct behavior (the server writes 3.14 through the
        byref).
        """
        cy = tagCY(0)
        with self.assertRaises(KeyError) as cm:
            self.server.InitCyVariantWith3dot14(byref(cy))
        self.assertIn("tagCY", str(cm.exception))

    def test_pointer_to_cy_is_written_through(self):
        """`pointer(tagCY)` becomes a VT_BYREF | VT_CY VARIANT; the server
        writes through it. Return value is the outer VARIANT (VT_NULL).

        NOTE: This test currently expects a KeyError because `tagCY` is not
        registered in `comtypes.automation._ctype_to_vartype`. The mapping
        should include:
            tagCY: VT_CY,
        Once this is added to _ctype_to_vartype, this test should be updated
        to verify the correct behavior (the server writes 3.14 through the
        pointer).
        """
        cy = tagCY(0)
        with self.assertRaises(KeyError) as cm:
            self.server.InitCyVariantWith3dot14(pointer(cy))
        self.assertIn("tagCY", str(cm.exception))


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

        NOTE: This test currently expects a KeyError because `tagDEC` is not
        registered in `comtypes.automation._ctype_to_vartype`. The mapping
        should include:
            tagDEC: VT_DECIMAL,
        Once this is added to _ctype_to_vartype, this test should be updated
        to verify the correct behavior (the server writes 3.14 through the
        byref).
        """
        dec = tagDEC()
        with self.assertRaises(KeyError) as cm:
            self.server.InitDecVariantWith3dot14(byref(dec))
        self.assertIn("tagDEC", str(cm.exception))

    def test_pointer_to_decimal_is_written_through(self):
        """`pointer(tagDEC)` becomes a VT_BYREF | VT_DECIMAL VARIANT; the server
        writes through it. Return value is the outer VARIANT (VT_NULL).

        NOTE: This test currently expects a KeyError because `tagDEC` is not
        registered in `comtypes.automation._ctype_to_vartype`. The mapping
        should include:
            tagDEC: VT_DECIMAL,
        Once this is added to _ctype_to_vartype, this test should be updated
        to verify the correct behavior (the server writes 3.14 through the
        pointer).
        """
        dec = tagDEC()
        with self.assertRaises(KeyError) as cm:
            self.server.InitDecVariantWith3dot14(pointer(dec))
        self.assertIn("tagDEC", str(cm.exception))


if __name__ == "__main__":
    unittest.main()
