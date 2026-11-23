"""Tests for parameterized (named) properties via the native C++ COM test server.

A COM property that takes arguments is exposed by `comtypes` through its
`named_property` descriptor.  The bound form of that descriptor
(`bound_named_property`) supports `obj.Prop[index]`, `obj.Prop[index] = value`,
`obj.Prop(index)` and tuple indexing `obj.Prop[a, b]`.  These tests exercise
that behavior against a real out-of-process COM server, so the argument and
return-value marshalling of parameterized properties is covered as well.

The test double `CoNamedPropertyTest` exposes the same set of parameterized
properties through two interfaces:

* `IDualNamedPropertyTest` - reached through the dual vtable.
* `IDispNamedPropertyTest` - reached through `IDispatch::Invoke`.

Both interfaces are tested because `comtypes` builds the `named_property`
descriptor from two different code paths (`ComPropertyGenerator` for dual
interfaces and `DispPropertyGenerator` for dispinterfaces).

The properties follow this contract:

* Reading an index that has not been assigned returns a VARIANT whose scode is
  DISP_E_BADINDEX, so "not set" can be told apart from a stored value.
* `Range` takes two integer indices that denote a range.  Reading
  `obj.Range[start, end]` returns one element per index in `[start, end]`,
  filling an index that has not been assigned with a VT_NULL VARIANT (which
  comtypes converts to `None`).  Writing `obj.Range[start, end] = values`
  assigns `values[i]` to index `start + i`; the number of values must equal
  `end - start + 1`.
"""

import unittest
from _ctypes import COMError

from comtypes import CLSCTX_LOCAL_SERVER, hresult
from comtypes.client import CreateObject, GetModule

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    # Server type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDualNamedProperty(unittest.TestCase):
    """Parameterized properties reached through the dual vtable interface."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.NamedPropertyTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDualNamedPropertyTest,
        )

    def test_getitem_unset_index_returns_badindex_variant(self):
        """Reading an index that was never assigned must report DISP_E_BADINDEX.

        The getter returns a VARIANT instead of a plain long so that "not set"
        can be told apart from a stored value.  The VARIANT carries
        DISP_E_BADINDEX in its scode field.  (For the dual vtable path comtypes
        clears the VARIANT to VT_NULL after reading the out parameter, but the
        scode field is preserved, which is the same behavior the VT_ERROR
        round-trip tests rely on.)
        """
        result = self.obj.Item[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_setitem_then_getitem_round_trips(self):
        """`obj.Item[index] = value` must store the value for that index.

        This is the basic read/write contract of a parameterized property:
        the value assigned through `__setitem__` must be returned by a later
        `__getitem__` for the same index.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item[3], 42)

    def test_call_passes_index_to_getter(self):
        """`obj.Item(index)` must forward `index` to the property getter.

        `bound_named_property.__call__` is an alternative spelling of
        `__getitem__`; it must reach the same getter with the same argument.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item(3), 42)

    def test_getitem_range_unset_returns_nulls(self):
        """Reading a range that was never assigned must return VT_NULL elements.

        `Range[start, end]` returns one element per index in [start, end].  An
        index that has not been assigned is reported as a VT_NULL VARIANT,
        which comtypes converts to `None`.  This lets a caller tell "not set"
        apart from a stored value without inventing a sentinel.
        """
        self.assertEqual(self.obj.Range[1, 3], (None, None, None))

    def test_setitem_range_then_getitem_round_trips(self):
        """`obj.Range[start, end] = values` must assign values to the range.

        The setter receives the two indices followed by the assigned sequence;
        `values[i]` must be stored under index `start + i`.  A wrong argument
        order or an off-by-one would store the values under the wrong indices.
        """
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 3], (9, 8, 7))

    def test_call_with_multiple_args(self):
        """`obj.Range(start, end)` must forward both arguments to the getter.

        This is the call spelling of tuple indexing for a property with two
        arguments.
        """
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range(1, 3), (9, 8, 7))

    def test_getitem_range_fills_unset_indices_with_null(self):
        """Reading a range wider than what was assigned must fill with VT_NULL.

        After assigning only indices 1..3, reading 1..5 must return the three
        stored values followed by two VT_NULL elements (converted to `None`).
        This is the behavior that replaces the old pair-keyed Matrix property.
        """
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 5], (9, 8, 7, None, None))

    def test_setitem_range_wrong_length_raises_com_error(self):
        """Assigning a sequence whose length differs from the range must fail.

        `Range[1, 3]` covers three indices, so a two-element sequence cannot be
        assigned.  The server rejects the mismatch instead of silently storing
        a partial range.
        """
        with self.assertRaises(COMError) as cm:
            self.obj.Range[1, 3] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDispNamedProperty(unittest.TestCase):
    """Parameterized properties reached through the dispinterface."""

    def setUp(self):
        self.obj = CreateObject(
            "Comtypes.NamedPropertyTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDispNamedPropertyTest,
        )

    def test_getitem_unset_index_returns_badindex_variant(self):
        """Reading an index that was never assigned must report DISP_E_BADINDEX.

        The getter returns a VARIANT instead of a plain long so that "not set"
        can be told apart from a stored value.  The VARIANT carries
        DISP_E_BADINDEX in its scode field.
        """
        result = self.obj.Item[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_setitem_then_getitem_round_trips(self):
        """`obj.Item[index] = value` must store the value for that index.

        This is the basic read/write contract of a parameterized property:
        the value assigned through `__setitem__` must be returned by a later
        `__getitem__` for the same index.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item[3], 42)

    def test_call_passes_index_to_getter(self):
        """`obj.Item(index)` must forward `index` to the property getter.

        `bound_named_property.__call__` is an alternative spelling of
        `__getitem__`; it must reach the same getter with the same argument.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item(3), 42)

    def test_getitem_range_unset_returns_nulls(self):
        """Reading a range that was never assigned must return VT_NULL elements.

        `Range[start, end]` returns one element per index in [start, end].  An
        index that has not been assigned is reported as a VT_NULL VARIANT,
        which comtypes converts to `None`.  This lets a caller tell "not set"
        apart from a stored value without inventing a sentinel.
        """
        self.assertEqual(self.obj.Range[1, 3], (None, None, None))

    def test_setitem_range_then_getitem_round_trips(self):
        """`obj.Range[start, end] = values` must assign values to the range.

        The setter receives the two indices followed by the assigned sequence;
        `values[i]` must be stored under index `start + i`.  A wrong argument
        order or an off-by-one would store the values under the wrong indices.
        """
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 3], (9, 8, 7))

    def test_call_with_multiple_args(self):
        """`obj.Range(start, end)` must forward both arguments to the getter.

        This is the call spelling of tuple indexing for a property with two
        arguments.
        """
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range(1, 3), (9, 8, 7))

    def test_getitem_range_fills_unset_indices_with_null(self):
        """Reading a range wider than what was assigned must fill with VT_NULL.

        After assigning only indices 1..3, reading 1..5 must return the three
        stored values followed by two VT_NULL elements (converted to `None`).
        This is the behavior that replaces the old pair-keyed Matrix property.
        """
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 5], (9, 8, 7, None, None))

    def test_setitem_range_wrong_length_raises_com_error(self):
        """Assigning a sequence whose length differs from the range must fail.

        `Range[1, 3]` covers three indices, so a two-element sequence cannot be
        assigned.  The server rejects the mismatch instead of silently storing
        a partial range.  The dispinterface setter reaches `IDispatch::Invoke`,
        and `ITypeInfo::Invoke` wraps the failure HRESULT returned by the
        method into DISP_E_EXCEPTION, so that is what surfaces here (unlike the
        dual vtable setter, which reports the method's own HRESULT).
        """
        with self.assertRaises(COMError) as cm:
            self.obj.Range[1, 3] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)


if __name__ == "__main__":
    unittest.main()
