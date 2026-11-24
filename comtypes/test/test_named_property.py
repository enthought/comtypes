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
* Every parameterized property (`Item`, `ReadOnlyItem`, `WriteOnlyItem`,
  `OptionalItem`, `Range`, `OptionalRange` and `Stride`) shares a single map,
  so a value written through any of them is visible through all of them.
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
* `Stride` combines a required argument (`start`), a default-valued argument
  (`step`, default 1) and an optional argument (`end`) in one signature, so the
  order of the arguments carries meaning: `start` says where to begin, `step`
  says how far to advance and `end` says where to stop.  Reading
  `obj.Stride[start, step, end]` returns one element per index in the
  arithmetic sequence `start, start+step, ..., <= end`, filling an index that
  has not been assigned with a VT_NULL VARIANT.  Writing
  `obj.Stride[start, step, end] = values` assigns `values[i]` to index
  `start + i*step`.  When `end` is omitted, reads extend through the last
  assigned index and writes cover as many values as were supplied.  Because
  `start` is required, the all-slice form (`obj.Stride[:]`) cannot supply it
  and is rejected.
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
        result = self.obj.Item[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_setitem_then_getitem_round_trips(self):
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item[3], 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_call_passes_index_to_getter(self):
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item(3), 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_getitem_range_unset_returns_nulls(self):
        self.assertEqual(self.obj.Range[1, 3], (None, None, None))

    def test_setitem_range_then_getitem_round_trips(self):
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 3], (9, 8, 7))
        self.assertEqual(self.obj.OptionalRange[:], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_call_with_multiple_args(self):
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range(1, 3), (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_getitem_range_fills_unset_indices_with_null(self):
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 5], (9, 8, 7, None, None))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_setitem_range_wrong_length_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Range[1, 3] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)

    def test_readonly_reads_item_map(self):
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.ReadOnlyItem[3], 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_readonly_unset_returns_badindex_variant(self):
        result = self.obj.ReadOnlyItem[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_readonly_setitem_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.ReadOnlyItem[3] = 1

    def test_writeonly_writes_item_map(self):
        self.obj.WriteOnlyItem[3] = 42
        self.assertEqual(self.obj.Item[3], 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_writeonly_getitem_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem[3]

    def test_writeonly_call_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem(3)

    def test_optional_getitem_all_slice_returns_pairs(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20)))

    def test_optional_getitem_empty_tuple_returns_pairs(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[()], ((1, 10), (2, 20)))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20)))

    def test_optional_getitem_index_returns_value(self):
        self.obj.OptionalItem[5] = 50
        self.assertEqual(self.obj.OptionalItem[5], 50)
        self.assertEqual(self.obj.OptionalItem[:], ((5, 50),))

    def test_optional_getitem_unset_index_returns_badindex_variant(self):
        result = self.obj.OptionalItem[5]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_optional_setitem_all_slice_replaces_map(self):
        # The vtable setter binds the supplied pairs as the index, leaving value unset.
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
        # The vtable setter binds the supplied pairs as the index, leaving value unset.
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
        # The empty-tuple form uses the same vtable call, binding pairs as the index.
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
        with self.assertRaises(TypeError):
            self.obj.Item[:] = 42

    def test_optional_range_getitem_both_args_returns_range(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_start_only_reads_to_last_index(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_all_slice_uses_default_start(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[:], (None, 9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_empty_tuple_uses_default_start(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[()], (None, 9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_call_both_args(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2, 4), (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_call_start_only(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2), (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_call_no_args(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(), (None, 9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_unset_range_returns_nulls(self):
        self.assertEqual(self.obj.OptionalRange[2, 4], (None, None, None))
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_range_getitem_empty_map_returns_empty(self):
        self.assertEqual(self.obj.OptionalRange[:], ())
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_range_getitem_start_beyond_data_returns_empty(self):
        self.assertEqual(self.obj.OptionalRange[5], ())
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_range_getitem_end_before_start_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[5, 3]
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)

    def test_optional_range_setitem_both_args_round_trips(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_setitem_wrong_length_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[2, 4] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)

    def test_optional_range_setitem_start_only_raises_type_error(self):
        # The vtable setter binds values as end, leaving the required value unset.
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
        # The vtable setter binds values as start, leaving the required value unset.
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
        # The empty-tuple form binds values as start, leaving the required value unset.
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

    def test_stride_getitem_all_args_returns_range(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2, 5], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_getitem_unset_returns_nulls(self):
        self.assertEqual(self.obj.Stride[1, 1, 3], (None, None, None))
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_stride_getitem_fills_unset_indices_with_null(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 1, 5], (10, None, 20, None, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_getitem_start_only_uses_default_step_and_reads_to_last(self):
        self.obj.Stride[2, 1, 4] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[2], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 10), (3, 20), (4, 30)))

    def test_stride_getitem_start_and_step_omits_end(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_getitem_empty_map_returns_empty(self):
        self.assertEqual(self.obj.Stride[1], ())
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_stride_getitem_start_beyond_data_returns_empty(self):
        self.obj.Stride[1, 1, 3] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[5], ())
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20), (3, 30)))

    def test_stride_getitem_end_before_start_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Stride[5, 1, 3]
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)

    def test_stride_getitem_zero_step_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Stride[1, 0, 3]
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)

    def test_stride_setitem_all_args_round_trips(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2, 5], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_setitem_start_only_uses_default_step_and_end(self):
        # The vtable setter binds the supplied values as step, leaving the
        # required value argument unset.
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the start-only "
                "setter is fixed, self.obj.Stride[1] = (10, 20, 30) should "
                "write indices 1, 2, 3 and self.obj.Stride[1, 1, 3] == (10, 20, 30)"
                "shoud pass"
            ),
        ) as cm:
            self.obj.Stride[1] = (10, 20, 30)
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_stride_setitem_start_and_step_omits_end(self):
        # The vtable setter binds the supplied values as end, leaving the
        # required value argument unset.
        with self.assertRaises(
            TypeError,
            msg=(
                "CHARACTERIZATION: failure is expected; once the start-and-step "
                "setter is fixed, self.obj.Stride[1, 2] = (10, 20, 30) should "
                "write indices 1, 3, 5 and self.obj.Stride[1, 2, 5] == (10, 20, 30) "
                "should pass"
            ),
        ) as cm:
            self.obj.Stride[1, 2] = (10, 20, 30)
        self.assertIn("required argument 'value' missing", str(cm.exception))

    def test_stride_setitem_wrong_length_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Stride[1, 1, 3] = (10, 20)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)

    def test_stride_setitem_all_slice_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.Stride[:] = (10, 20, 30)


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
        result = self.obj.Item[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_setitem_then_getitem_round_trips(self):
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item[3], 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_call_passes_index_to_getter(self):
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.Item(3), 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_getitem_range_unset_returns_nulls(self):
        self.assertEqual(self.obj.Range[1, 3], (None, None, None))

    def test_setitem_range_then_getitem_round_trips(self):
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 3], (9, 8, 7))
        self.assertEqual(self.obj.OptionalRange[:], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_call_with_multiple_args(self):
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range(1, 3), (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_getitem_range_fills_unset_indices_with_null(self):
        self.obj.Range[1, 3] = (9, 8, 7)
        self.assertEqual(self.obj.Range[1, 5], (9, 8, 7, None, None))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_setitem_range_wrong_length_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Range[1, 3] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_readonly_reads_item_map(self):
        self.obj.Item[3] = 42
        self.assertEqual(self.obj.ReadOnlyItem[3], 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_readonly_unset_returns_badindex_variant(self):
        result = self.obj.ReadOnlyItem[3]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_readonly_setitem_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.ReadOnlyItem[3] = 1

    def test_writeonly_writes_item_map(self):
        self.obj.WriteOnlyItem[3] = 42
        self.assertEqual(self.obj.Item[3], 42)
        self.assertEqual(self.obj.OptionalItem[:], ((3, 42),))

    def test_writeonly_getitem_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem[3]

    def test_writeonly_call_raises_type_error(self):
        with self.assertRaises(TypeError):
            self.obj.WriteOnlyItem(3)

    def test_optional_getitem_all_slice_returns_pairs(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20)))

    def test_optional_getitem_empty_tuple_returns_pairs(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[2] = 20
        self.assertEqual(self.obj.OptionalItem[()], ((1, 10), (2, 20)))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20)))

    def test_optional_getitem_index_returns_value(self):
        self.obj.OptionalItem[5] = 50
        self.assertEqual(self.obj.OptionalItem[5], 50)
        self.assertEqual(self.obj.OptionalItem[:], ((5, 50),))

    def test_optional_getitem_unset_index_returns_badindex_variant(self):
        result = self.obj.OptionalItem[5]
        self.assertEqual(result._.scode, hresult.DISP_E_BADINDEX)

    def test_optional_setitem_all_slice_replaces_map(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[:] = ((3, 30), (4, 40))
        self.assertEqual(self.obj.OptionalItem[:], ((3, 30), (4, 40)))

    def test_optional_setitem_all_slice_empty_clears_map(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[:] = ()
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_setitem_empty_tuple_replaces_map(self):
        self.obj.OptionalItem[1] = 10
        self.obj.OptionalItem[()] = ((3, 30), (4, 40))
        self.assertEqual(self.obj.OptionalItem[:], ((3, 30), (4, 40)))

    def test_setitem_all_slice_with_required_index_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Item[:] = 42
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_BADPARAMCOUNT)

    def test_optional_range_getitem_both_args_returns_range(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_start_only_reads_to_last_index(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_all_slice_uses_default_start(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[:], (None, 9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_empty_tuple_uses_default_start(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[()], (None, 9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_call_both_args(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2, 4), (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_call_start_only(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(2), (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_call_no_args(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange(), (None, 9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_getitem_unset_range_returns_nulls(self):
        self.assertEqual(self.obj.OptionalRange[2, 4], (None, None, None))
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_range_getitem_empty_map_returns_empty(self):
        self.assertEqual(self.obj.OptionalRange[:], ())
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_range_getitem_start_beyond_data_returns_empty(self):
        self.assertEqual(self.obj.OptionalRange[5], ())
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_optional_range_getitem_end_before_start_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[5, 3]
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_optional_range_setitem_both_args_round_trips(self):
        self.obj.OptionalRange[2, 4] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_setitem_wrong_length_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.OptionalRange[2, 4] = (9, 8)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_optional_range_setitem_start_only_writes_to_end(self):
        self.obj.OptionalRange[2] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[2, 4], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 9), (3, 8), (4, 7)))

    def test_optional_range_setitem_all_slice_uses_default_start(self):
        self.obj.OptionalRange[:] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[1, 3], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_optional_range_setitem_empty_tuple_uses_default_start(self):
        self.obj.OptionalRange[()] = (9, 8, 7)
        self.assertEqual(self.obj.OptionalRange[1, 3], (9, 8, 7))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 9), (2, 8), (3, 7)))

    def test_stride_getitem_all_args_returns_range(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2, 5], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_getitem_unset_returns_nulls(self):
        self.assertEqual(self.obj.Stride[1, 1, 3], (None, None, None))
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_stride_getitem_fills_unset_indices_with_null(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 1, 5], (10, None, 20, None, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_getitem_start_only_uses_default_step_and_reads_to_last(self):
        self.obj.Stride[2, 1, 4] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[2], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((2, 10), (3, 20), (4, 30)))

    def test_stride_getitem_start_and_step_omits_end(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_getitem_empty_map_returns_empty(self):
        self.assertEqual(self.obj.Stride[1], ())
        self.assertEqual(self.obj.OptionalItem[:], ())

    def test_stride_getitem_start_beyond_data_returns_empty(self):
        self.obj.Stride[1, 1, 3] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[5], ())
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (2, 20), (3, 30)))

    def test_stride_getitem_end_before_start_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Stride[5, 1, 3]
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_stride_getitem_zero_step_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Stride[1, 0, 3]
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_stride_setitem_all_args_round_trips(self):
        self.obj.Stride[1, 2, 5] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2, 5], (10, 20, 30))
        self.assertEqual(self.obj.OptionalItem[:], ((1, 10), (3, 20), (5, 30)))

    def test_stride_setitem_start_only_uses_default_step_and_end(self):
        self.obj.Stride[1] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 1, 3], (10, 20, 30))

    def test_stride_setitem_start_and_step_omits_end(self):
        self.obj.Stride[1, 2] = (10, 20, 30)
        self.assertEqual(self.obj.Stride[1, 2, 5], (10, 20, 30))

    def test_stride_setitem_wrong_length_raises_com_error(self):
        with self.assertRaises(COMError) as cm:
            self.obj.Stride[1, 1, 3] = (10, 20)
        self.assertEqual(cm.exception.hresult, hresult.DISP_E_EXCEPTION)

    def test_stride_setitem_all_slice_raises_com_error(self):
        with self.assertRaises(COMError):
            self.obj.Stride[:] = (10, 20, 30)


if __name__ == "__main__":
    unittest.main()
