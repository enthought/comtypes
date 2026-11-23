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
* `OptionalRange` has a `start` argument defaulting to 1 and an optional `end`.
  With both bounds it behaves like `Range`; if `end` is omitted, reads extend
  through the last assigned index.  The all-slice and empty-tuple forms omit
  `start` and therefore use its default.  The dispinterface setter also accepts
  omitted bounds, while the dual-vtable setter currently raises `TypeError`
  for those forms because the value argument is not bound correctly.
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

    def test_optional_range_getitem_both_args_returns_range(self):
        """`obj.OptionalRange[start, end]` must return the range [start, end].

        When both arguments are given, `OptionalRange` behaves like `Range`:
        one element per index in [start, end], with an unassigned index filled
        with a VT_NULL VARIANT (converted to `None`).
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))

    def test_optional_range_getitem_start_only_reads_to_last_index(self):
        """`obj.OptionalRange[start]` must read from `start` to the last index.

        The second argument is optional.  When it is omitted, the getter reads
        from `start` to the last assigned index, so the whole stored range is
        returned.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2], (9, 8, 7))

    def test_optional_range_getitem_all_slice_uses_default_start(self):
        """`obj.OptionalRange[:]` must use the default start value 1.

        The first argument has the default value 1.  The all-slice spelling
        passes no argument, so the getter reads from index 1 to the last
        assigned index.  Index 1 was never assigned, so it is filled with
        VT_NULL (converted to `None`).
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[:], (None, 9, 8, 7))

    def test_optional_range_getitem_empty_tuple_uses_default_start(self):
        """`obj.OptionalRange[()]` must behave like the all-slice spelling.

        An empty tuple index is treated as "no argument", so the default start
        value 1 is used and the getter reads to the last assigned index.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[()], (None, 9, 8, 7))

    def test_optional_range_call_both_args(self):
        """`obj.OptionalRange(start, end)` must forward both arguments.

        This is the call spelling of tuple indexing for a property that mixes a
        default-valued argument with an optional one.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2, 4), (9, 8, 7))

    def test_optional_range_call_start_only(self):
        """`obj.OptionalRange(start)` must forward only `start`.

        The omitted `end` must fall back to the optional default, so the getter
        reads to the last assigned index.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2), (9, 8, 7))

    def test_optional_range_call_no_args(self):
        """`obj.OptionalRange()` must use the default start value 1.

        With no arguments, `start` falls back to its default 1 and `end` is
        omitted, so the getter reads from index 1 to the last assigned index.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(), (None, 9, 8, 7))

    def test_optional_range_getitem_unset_range_returns_nulls(self):
        """Reading an unassigned range must return VT_NULL elements.

        With an explicit `end`, `OptionalRange` behaves like `Range`: an index
        that has not been assigned is reported as a VT_NULL VARIANT, which
        comtypes converts to `None`.
        """
        self.assertEqual(self.obj.OptionalRange[2, 4], (None, None, None))

    def test_optional_range_getitem_empty_map_returns_empty(self):
        """Reading with an omitted `end` on an empty map must return an empty tuple.

        There is no last assigned index to read to, so the getter returns an
        empty array, which comtypes converts to an empty tuple.
        """
        self.assertEqual(self.obj.OptionalRange[:], ())

    def test_optional_range_getitem_start_beyond_data_returns_empty(self):
        """Reading from a start beyond the last index must return an empty tuple.

        When `end` is omitted and `start` is greater than the last assigned
        index, there is nothing to read, so the getter returns an empty array.
        """
        self.assertEqual(self.obj.OptionalRange[5], ())

    def test_optional_range_getitem_end_before_start_raises_com_error(self):
        """An explicit `end` before `start` must be rejected.

        With an explicit `end`, `OptionalRange` behaves like `Range`, which
        rejects `end < start` with E_INVALIDARG.
        """
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[5, 3]
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)

    def test_optional_range_setitem_both_args_round_trips(self):
        """`obj.OptionalRange[start, end] = values` must assign the range.

        With both arguments, `OptionalRange` behaves like `Range`:
        `values[i]` is stored under index `start + i`.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))

    def test_optional_range_setitem_wrong_length_raises_com_error(self):
        """Assigning a sequence whose length differs from the range must fail.

        `OptionalRange[2, 4]` covers three indices, so a two-element sequence
        cannot be assigned.  The server rejects the mismatch with
        DISP_E_BADPARAMCOUNT.
        """
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[2, 4] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)

    def test_optional_range_setitem_start_only_raises_type_error(self):
        """`obj.OptionalRange[start] = values` must raise TypeError.

        CHARACTERIZATION: this records the current (buggy) behavior, not the
        desired one.  The setter is a bound method whose parameters are
        declared as `(start, end, value)`, and
        `bound_named_property.__setitem__` calls it as `fset(instance, start,
        value)` for a single (non-tuple) index.  That binds `value` to the
        `end` parameter, so the `value` parameter is left unfilled and Python
        raises `TypeError: required argument 'value' missing` before the COM
        call is made.  The dispinterface path is unaffected because its setter
        is a variadic wrapper around `IDispatch::Invoke`.  When the
        single-index setter is fixed to pass the value in the right position,
        this test will start failing and should be updated to assert that the
        range is written.
        """
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the single-index "
                "setter is fixed, self.obj.OptionalRange[2] = (9, 8, 7) should "
                "write indices 2..4"
            ),
        ) as cm:
            self.obj.OptionalRange[2] = (9, 8, 7)
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_optional_range_setitem_all_slice_raises_type_error(self):
        """`obj.OptionalRange[:] = values` must raise TypeError.

        CHARACTERIZATION: this records the current (buggy) behavior, not the
        desired one.  The all-slice spelling calls the dual vtable setter as
        `fset(instance, value)`, so the single positional argument binds to the
        `start` parameter and the `value` parameter is left unfilled; Python
        raises `TypeError: required argument 'value' missing` before the COM
        call is made.  When the all-slice setter is fixed, this test will start
        failing and should be updated to assert that the range is written from
        the default start.
        """
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the all-slice "
                "setter is fixed, self.obj.OptionalRange[:] = (9, 8, 7) should "
                "write indices 1..3"
            ),
        ) as cm:
            self.obj.OptionalRange[:] = (9, 8, 7)
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_optional_range_setitem_empty_tuple_raises_type_error(self):
        """`obj.OptionalRange[()] = values` must raise TypeError.

        CHARACTERIZATION: this records the current (buggy) behavior, not the
        desired one.  `bound_named_property.__setitem__` expands the empty
        tuple index to `fset(instance, value)`, which is the same call as the
        all-slice spelling.  On the dual vtable interface that single
        positional argument binds to the `start` parameter, leaving the `value`
        parameter unfilled, so Python raises `TypeError: required argument
        'value' missing` before the COM call is made.  When the empty-tuple
        setter is fixed, this test will start failing and should be updated to
        assert that the range is written from the default start.
        """
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the empty-tuple "
                "setter is fixed, self.obj.OptionalRange[()] = (9, 8, 7) should "
                "write indices 1..3"
            ),
        ) as cm:
            self.obj.OptionalRange[()] = (9, 8, 7)
        self.assertIn("required argument 'value' missing", str(cm.exception))


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

    def test_optional_range_getitem_both_args_returns_range(self):
        """`obj.OptionalRange[start, end]` must return the range [start, end].

        When both arguments are given, `OptionalRange` behaves like `Range`:
        one element per index in [start, end], with an unassigned index filled
        with a VT_NULL VARIANT (converted to `None`).
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))

    def test_optional_range_getitem_start_only_reads_to_last_index(self):
        """`obj.OptionalRange[start]` must read from `start` to the last index.

        The second argument is optional.  When it is omitted, the getter reads
        from `start` to the last assigned index, so the whole stored range is
        returned.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2], (9, 8, 7))

    def test_optional_range_getitem_all_slice_uses_default_start(self):
        """`obj.OptionalRange[:]` must use the default start value 1.

        The first argument has the default value 1.  The all-slice spelling
        passes no argument, so the getter reads from index 1 to the last
        assigned index.  Index 1 was never assigned, so it is filled with
        VT_NULL (converted to `None`).
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[:], (None, 9, 8, 7))

    def test_optional_range_getitem_empty_tuple_uses_default_start(self):
        """`obj.OptionalRange[()]` must behave like the all-slice spelling.

        An empty tuple index is treated as "no argument", so the default start
        value 1 is used and the getter reads to the last assigned index.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[()], (None, 9, 8, 7))

    def test_optional_range_call_both_args(self):
        """`obj.OptionalRange(start, end)` must forward both arguments.

        This is the call spelling of tuple indexing for a property that mixes a
        default-valued argument with an optional one.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2, 4), (9, 8, 7))

    def test_optional_range_call_start_only(self):
        """`obj.OptionalRange(start)` must forward only `start`.

        The omitted `end` must fall back to the optional default, so the getter
        reads to the last assigned index.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2), (9, 8, 7))

    def test_optional_range_call_no_args(self):
        """`obj.OptionalRange()` must use the default start value 1.

        With no arguments, `start` falls back to its default 1 and `end` is
        omitted, so the getter reads from index 1 to the last assigned index.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(), (None, 9, 8, 7))

    def test_optional_range_getitem_unset_range_returns_nulls(self):
        """Reading an unassigned range must return VT_NULL elements.

        With an explicit `end`, `OptionalRange` behaves like `Range`: an index
        that has not been assigned is reported as a VT_NULL VARIANT, which
        comtypes converts to `None`.
        """
        self.assertEqual(self.obj.OptionalRange[2, 4], (None, None, None))

    def test_optional_range_getitem_empty_map_returns_empty(self):
        """Reading with an omitted `end` on an empty map must return an empty tuple.

        There is no last assigned index to read to, so the getter returns an
        empty array, which comtypes converts to an empty tuple.
        """
        self.assertEqual(self.obj.OptionalRange[:], ())

    def test_optional_range_getitem_start_beyond_data_returns_empty(self):
        """Reading from a start beyond the last index must return an empty tuple.

        When `end` is omitted and `start` is greater than the last assigned
        index, there is nothing to read, so the getter returns an empty array.
        """
        self.assertEqual(self.obj.OptionalRange[5], ())

    def test_optional_range_getitem_end_before_start_raises_com_error(self):
        """An explicit `end` before `start` must be rejected.

        With an explicit `end`, `OptionalRange` behaves like `Range`, which
        rejects `end < start` with E_INVALIDARG.  The dispinterface getter
        reaches `IDispatch::Invoke`, and `ITypeInfo::Invoke` wraps the failure
        HRESULT returned by the method into DISP_E_EXCEPTION, so that is what
        surfaces here (unlike the dual vtable getter, which reports the
        method's own HRESULT).
        """
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[5, 3]
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_optional_range_setitem_both_args_round_trips(self):
        """`obj.OptionalRange[start, end] = values` must assign the range.

        With both arguments, `OptionalRange` behaves like `Range`:
        `values[i]` is stored under index `start + i`.
        """
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))

    def test_optional_range_setitem_wrong_length_raises_com_error(self):
        """Assigning a sequence whose length differs from the range must fail.

        `OptionalRange[2, 4]` covers three indices, so a two-element sequence
        cannot be assigned.  The dispinterface setter reaches
        `IDispatch::Invoke`, and `ITypeInfo::Invoke` wraps the failure HRESULT
        returned by the method into DISP_E_EXCEPTION, so that is what surfaces
        here (unlike the dual vtable setter, which reports the method's own
        HRESULT).
        """
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[2, 4] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_optional_range_setitem_start_only_writes_to_end(self):
        """`obj.OptionalRange[start] = values` must write from `start` to the end.

        The dispinterface setter is a variadic wrapper around
        `IDispatch::Invoke`.  The last argument is passed as the PROPERTYPUT
        value and the preceding arguments are passed positionally, so a single
        index is taken as `start` and the omitted `end` is left at its optional
        default.  The server then writes the values from `start` to the end of
        the supplied sequence.
        """
        self.obj.OptionalRange[2] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))

    def test_optional_range_setitem_all_slice_uses_default_start(self):
        """`obj.OptionalRange[:] = values` must write from the default start 1.

        The all-slice spelling passes no positional argument, so `start` falls
        back to its default 1 and `end` is omitted.  The server writes the
        values from index 1 to the end of the supplied sequence.
        """
        self.obj.OptionalRange[:] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[1, 3], (9, 8, 7))

    def test_optional_range_setitem_empty_tuple_uses_default_start(self):
        """`obj.OptionalRange[()] = values` must behave like the all-slice spelling.

        An empty tuple index is treated as "no argument", so `start` falls back
        to its default 1 and `end` is omitted, just like the all-slice spelling.
        """
        self.obj.OptionalRange[()] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[1, 3], (9, 8, 7))


if __name__ == "__main__":
    unittest.main()
