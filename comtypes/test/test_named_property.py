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
* `ReadOnlyItem` and `WriteOnlyItem` share the same storage as `Item`; they
  only read or only write that single map.
* `OptionalItem` can exchange the whole map through the all-slice spelling
  (`obj.OptionalItem[:]`) or the empty tuple (`obj.OptionalItem[()]`).  The map
  is represented as a sequence of (key, value) pairs.
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

    def test_readonly_reads_item_map(self):
        """`ReadOnlyItem` must read the same map that `Item` writes.

        `ReadOnlyItem` has a getter but no setter.  It shares storage with
        `Item`, so a value stored through `Item` must be visible through
        `ReadOnlyItem` for the same index.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.ReadOnlyItem[3], 42)

    def test_readonly_unset_returns_badindex_variant(self):
        """Reading an unassigned index through `ReadOnlyItem` reports DISP_E_BADINDEX.

        `ReadOnlyItem` shares the `Item` map, so an index that was never
        assigned must report DISP_E_BADINDEX just like `Item` does.
        """
        result = self.obj.ReadOnlyItem[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_readonly_setitem_raises_type_error(self):
        """Assigning to a read-only parameterized property must raise TypeError.

        `bound_named_property.__setitem__` must reject the assignment when the
        property has no setter, instead of silently doing nothing or crashing.
        """
        with self.assertRaises(TypeError):
            self.obj.ReadOnlyItem[3] = 1

    def test_writeonly_writes_item_map(self):
        """`WriteOnlyItem` must write the same map that `Item` reads.

        `WriteOnlyItem` has a setter but no getter.  It shares storage with
        `Item`, so a value stored through `WriteOnlyItem` must be visible
        through `Item` for the same index.
        """
        self.obj.WriteOnlyItem[3] = 42
        self.assertEqual(self.obj.Item[3], 42)

    def test_writeonly_getitem_raises_type_error(self):
        """Reading a write-only parameterized property must raise TypeError.

        `bound_named_property.__getitem__` must reject the read when the
        property has no getter.
        """
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem[3]

    def test_writeonly_call_raises_type_error(self):
        """Calling a write-only parameterized property must raise TypeError.

        `bound_named_property.__call__` must reject the call when the property
        has no getter.
        """
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem(3)

    def test_optional_getitem_all_slice_returns_pairs(self):
        """`obj.OptionalItem[:]` must return the whole map as (key, value) pairs.

        The all-slice spelling passes no index, which the server interprets as
        "the whole map".  The map is returned as a sequence of (key, value)
        tuples, sorted by key because the server stores it in a std::map.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20)))

    def test_optional_getitem_empty_tuple_returns_pairs(self):
        """`obj.OptionalItem[()]` must behave like the all-slice spelling.

        An empty tuple index is treated as "no index" by
        `bound_named_property.__getitem__`, so it must return the whole map
        just like `obj.OptionalItem[:]`.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[()], ((1, 10), (2, 20)))

    def test_optional_getitem_index_returns_value(self):
        """`obj.OptionalItem[index]` must return the value stored for that index.

        When an index is supplied, `OptionalItem` behaves like `Item` and
        returns the single stored value.
        """
        self.obj.OptionalItem[5] = 50
        self.assertEqual(self.obj.OptionalItem[5], 50)

    def test_optional_getitem_unset_index_returns_badindex_variant(self):
        """Reading an unassigned `OptionalItem` index reports DISP_E_BADINDEX.

        With an index, `OptionalItem` behaves like `Item`, so an index that was
        never assigned must report DISP_E_BADINDEX.
        """
        result = self.obj.OptionalItem[5]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_optional_setitem_all_slice_replaces_map(self):
        """`obj.OptionalItem[:] = pairs` must replace the whole map.

        The all-slice setter must not merge with the existing entries: the map
        is fully overwritten by the supplied (key, value) pairs.

        CHARACTERIZATION: this records the current (buggy) behavior, not the
        desired one.  On the dual vtable interface the setter is a bound method
        whose parameters are declared as `(index, value)`, and
        `bound_named_property.__setitem__` calls it as `fset(instance, value)`
        for the all-slice spelling.  That single positional argument binds to
        the `index` parameter, so the `value` parameter is left unfilled and
        Python raises `TypeError: required argument 'value' missing` before the
        COM call is made.  The dispinterface path is unaffected because its
        setter is a variadic wrapper around `IDispatch::Invoke`.  When the
        all-slice setter is fixed to pass the value in the right position, this
        test will start failing and should be updated to assert that the map is
        replaced.
        """
        self.obj.OptionalItem[1] = 10
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the all-slice setter "
                "is fixed, self.obj.OptionalItem[:] == ((3, 30), (4, 40)) should pass"
            ),
        ) as cm:
            self.obj.OptionalItem[:] = ((3, 30), (4, 40))
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_optional_setitem_all_slice_empty_clears_map(self):
        """`obj.OptionalItem[:] = ()` must clear the whole map.

        An empty sequence of pairs must leave the map empty, not keep the
        previous entries.

        CHARACTERIZATION: this records the current (buggy) behavior, not the
        desired one.  The all-slice spelling calls the dual vtable setter as
        `fset(instance, value)`, so the single positional argument binds to the
        `index` parameter and the `value` parameter is left unfilled; Python
        raises `TypeError: required argument 'value' missing` before the COM
        call is made.  When the all-slice setter is fixed, this test will start
        failing and should be updated to assert that the map is cleared.
        """
        self.obj.OptionalItem[1] = 10
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the all-slice setter "
                "is fixed, self.obj.OptionalItem[:] == () should pass"
            ),
        ) as cm:
            self.obj.OptionalItem[:] = ()
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_optional_setitem_empty_tuple_replaces_map(self):
        """`obj.OptionalItem[()] = pairs` must replace the whole map.

        An empty tuple index means "no index", just like the all-slice
        spelling, so the setter receives the index defaults followed by the
        value and the whole map is replaced.

        CHARACTERIZATION: this records the current (buggy) behavior, not the
        desired one.  `bound_named_property.__setitem__` expands the empty
        tuple index to `fset(instance, value)`, which is the same call as the
        all-slice spelling.  On the dual vtable interface that single
        positional argument binds to the `index` parameter, leaving the `value`
        parameter unfilled, so Python raises `TypeError: required argument
        'value' missing` before the COM call is made.  When the empty-tuple
        setter is fixed, this test will start failing and should be updated to
        assert that the map is replaced.
        """
        self.obj.OptionalItem[1] = 10
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the empty-tuple setter "
                "is fixed, self.obj.OptionalItem[:] == ((3, 30), (4, 40)) should pass"
            ),
        ) as cm:
            self.obj.OptionalItem[()] = ((3, 30), (4, 40))
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_setitem_all_slice_with_required_index_raises_type_error(self):
        """`obj.Item[:] = value` must raise TypeError when the index is required.

        `Item` has a required index, so the all-slice form cannot fill in a
        default for it.  The setter is called with a single argument, which
        binds to the index parameter, leaving the value parameter missing.
        """
        with self.assertRaises(TypeError):
            self.obj.Item[:] = 42


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

    def test_readonly_reads_item_map(self):
        """`ReadOnlyItem` must read the same map that `Item` writes.

        `ReadOnlyItem` has a getter but no setter.  It shares storage with
        `Item`, so a value stored through `Item` must be visible through
        `ReadOnlyItem` for the same index.
        """
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.ReadOnlyItem[3], 42)

    def test_readonly_unset_returns_badindex_variant(self):
        """Reading an unassigned index through `ReadOnlyItem` reports DISP_E_BADINDEX.

        `ReadOnlyItem` shares the `Item` map, so an index that was never
        assigned must report DISP_E_BADINDEX just like `Item` does.
        """
        result = self.obj.ReadOnlyItem[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_readonly_setitem_raises_type_error(self):
        """Assigning to a read-only parameterized property must raise TypeError.

        `bound_named_property.__setitem__` must reject the assignment when the
        property has no setter, instead of silently doing nothing or crashing.
        """
        with self.assertRaises(TypeError):
            self.obj.ReadOnlyItem[3] = 1

    def test_writeonly_writes_item_map(self):
        """`WriteOnlyItem` must write the same map that `Item` reads.

        `WriteOnlyItem` has a setter but no getter.  It shares storage with
        `Item`, so a value stored through `WriteOnlyItem` must be visible
        through `Item` for the same index.
        """
        self.obj.WriteOnlyItem[3] = 42
        self.assertEqual(self.obj.Item[3], 42)

    def test_writeonly_getitem_raises_type_error(self):
        """Reading a write-only parameterized property must raise TypeError.

        `bound_named_property.__getitem__` must reject the read when the
        property has no getter.
        """
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem[3]

    def test_writeonly_call_raises_type_error(self):
        """Calling a write-only parameterized property must raise TypeError.

        `bound_named_property.__call__` must reject the call when the property
        has no getter.
        """
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem(3)

    def test_optional_getitem_all_slice_returns_pairs(self):
        """`obj.OptionalItem[:]` must return the whole map as (key, value) pairs.

        The all-slice spelling passes no index, which the server interprets as
        "the whole map".  The map is returned as a sequence of (key, value)
        tuples, sorted by key because the server stores it in a std::map.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20)))

    def test_optional_getitem_empty_tuple_returns_pairs(self):
        """`obj.OptionalItem[()]` must behave like the all-slice spelling.

        An empty tuple index is treated as "no index" by
        `bound_named_property.__getitem__`, so it must return the whole map
        just like `obj.OptionalItem[:]`.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[()], ((1, 10), (2, 20)))

    def test_optional_getitem_index_returns_value(self):
        """`obj.OptionalItem[index]` must return the value stored for that index.

        When an index is supplied, `OptionalItem` behaves like `Item` and
        returns the single stored value.
        """
        self.obj.OptionalItem[5] = 50
        self.assertEqual(self.obj.OptionalItem[5], 50)

    def test_optional_getitem_unset_index_returns_badindex_variant(self):
        """Reading an unassigned `OptionalItem` index reports DISP_E_BADINDEX.

        With an index, `OptionalItem` behaves like `Item`, so an index that was
        never assigned must report DISP_E_BADINDEX.
        """
        result = self.obj.OptionalItem[5]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_optional_setitem_all_slice_replaces_map(self):
        """`obj.OptionalItem[:] = pairs` must replace the whole map.

        The all-slice setter must not merge with the existing entries: the map
        is fully overwritten by the supplied (key, value) pairs.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[:] = ((3, 30), (4, 40))
        self.assertEqual(self.obj.OptionalItem[:], ((3, 30), (4, 40)))

    def test_optional_setitem_all_slice_empty_clears_map(self):
        """`obj.OptionalItem[:] = ()` must clear the whole map.

        An empty sequence of pairs must leave the map empty, not keep the
        previous entries.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[:] = ()
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_setitem_empty_tuple_replaces_map(self):
        """`obj.OptionalItem[()] = pairs` must replace the whole map.

        The dispinterface setter is a variadic wrapper around
        `IDispatch::Invoke`, so a single argument is taken as the value and the
        optional index is left missing.  This makes the empty-tuple spelling
        work for the dispinterface, unlike the dual vtable setter.
        """
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[()] = ((3, 30), (4, 40))
        self.assertEqual(self.obj.OptionalItem[:], ((3, 30), (4, 40)))

    def test_setitem_all_slice_with_required_index_raises_com_error(self):
        """`obj.Item[:] = value` must raise COMError when the index is required.

        `Item` has a required index, so the all-slice form cannot fill in a
        default for it.  The dispinterface setter forwards the single argument
        as the value, leaving the index parameter missing.  Unlike the dual
        vtable setter (which raises TypeError before the call is made), the
        dispinterface setter reaches `IDispatch::Invoke`, so the server rejects
        the wrong argument count with DISP_E_BADPARAMCOUNT, surfaced as COMError.
        """
        with self.assertRaises(COMError) as cm:
            self.obj.Item[:] = 42
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)


if __name__ == "__main__":
    unittest.main()
