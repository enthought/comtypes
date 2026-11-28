"""Tests for the lifetime of [out] record parameters."""

import ctypes
import gc
import unittest
from _ctypes import COMError
from ctypes import ArgumentError, addressof, byref, pointer

from comtypes import CLSCTX_LOCAL_SERVER, hresult
from comtypes.automation import (
    DISPATCH_PROPERTYPUT,
    DISPATCH_PROPERTYPUTREF,
    VARIANT,
)
from comtypes.client import CreateObject, GetModule
from comtypes.recordinfo import GetRecordInfoFromGuids, IRecordInfo
from comtypes.test.find_memleak import find_memleak

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    # Server type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


def _create_server() -> "ComtypesCppTestSrvLib.IOutRecordTest":
    return CreateObject(
        "Comtypes.OutRecordTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IOutRecordTest,
    )


def _create_record_info() -> IRecordInfo:
    return GetRecordInfoFromGuids(
        *ComtypesCppTestSrvLib.StructOutRecordTest._recordinfo_
    )


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestServerBehavior(unittest.TestCase):
    def test_get_record_returns_record_with_expected_fields(self):
        """`GetRecord` returns a `StructOutRecordTest` through an [out]
        parameter.

        This verifies that comtypes converts the [out] record into the
        generated Python record type and that its fields carry the values the
        server wrote.  If this fails, the [out] record is either not converted
        to the expected type or its fields are not populated correctly.
        """
        obj = _create_server()
        rec = obj.GetRecord(4)
        self.assertIsInstance(rec, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(rec.tag, 1)
        self.assertEqual(rec.text, "aaaa")

    def test_get_record_retval_returns_record_with_expected_fields(self):
        """`GetRecordRetval` returns a `StructOutRecordTest` as the [out,
        retval] value.

        This verifies that comtypes converts a record returned as the [out,
        retval] value into the generated Python record type and that its fields
        carry the values the server wrote.  If this fails, the [out, retval]
        record is either not converted to the expected type or its fields are
        not populated correctly.
        """
        obj = _create_server()
        retval = obj.GetRecordRetval()
        self.assertIsInstance(retval, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(retval.tag, 2)

    def test_get_records_around_returns_records_and_plain_value(self):
        """`GetRecordsAround` returns two [out] records and a plain LONG.

        This verifies that comtypes converts the two [out] records into the
        generated Python record type while leaving the plain LONG parameter as
        an int, and that each carries the value the server wrote.  If this
        fails, one of the [out] records is not converted to the expected type,
        the plain value is not an int, or a field is not populated correctly.
        """
        obj = _create_server()
        first, middle, last = obj.GetRecordsAround()
        self.assertIsInstance(first, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(first.tag, 10)
        self.assertEqual(middle, 42)
        self.assertIsInstance(last, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(last.tag, 20)

    def test_modify_record_returns_modified_record(self):
        """`ModifyRecord` returns the [in, out] record with the modified field.

        This verifies that comtypes passes a record as an [in, out] parameter
        and converts the returned record into the generated Python record type,
        and that the server's modification is visible in the returned field.
        If this fails, the [in, out] record is not converted to the expected
        type or the modification is not reflected in the returned record.
        """
        obj = _create_server()
        inout = ComtypesCppTestSrvLib.StructOutRecordTest()
        inout.tag = 5
        modified = obj.ModifyRecord(inout)
        self.assertIsInstance(modified, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(modified.tag, 6)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestReleaseAutomatically(unittest.TestCase):
    def test_out_record_does_not_release_sentinel_on_del(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # `GetRecord` fills the [out] record with a sentinel `IUnknown` created
        # with a reference count of 1 and transfers that reference to the
        # caller, so one sentinel is alive.
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # The record owns the sentinel, so deleting the record is expected to
        # release it.
        del rec
        gc.collect()
        # CHARACTERIZATION: The sentinel is NOT released, count stays at 1.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            1,
            msg="CHARACTERIZATION: Expected sentinel to be released on del, "
            "but current behavior leaves it alive. This test records the "
            "current (buggy) behavior; it will fail when the bug is fixed.",
        )

    def test_retval_record_does_not_release_sentinel_on_del(self):
        obj = _create_server()
        rec = obj.GetRecordRetval()
        # Same ownership transfer as `test_out_record_does_not_release_sentinel_on_del`,
        # but the record is returned as the [out, retval] value: the sentinel
        # is created with a reference count of 1 and handed to the caller.
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # The record owns the sentinel, so deleting the record is expected to
        # release it.
        del rec
        gc.collect()
        # CHARACTERIZATION: The sentinel is NOT released, count stays at 1.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            1,
            msg="CHARACTERIZATION: Expected sentinel to be released on del, "
            "but current behavior leaves it alive. This test records the "
            "current (buggy) behavior; it will fail when the bug is fixed.",
        )

    def test_multiple_out_records_do_not_release_sentinels_on_del(self):
        obj = _create_server()
        first, middle, last = obj.GetRecordsAround()
        self.assertEqual(first.tag, 10)
        self.assertEqual(middle, 42)
        self.assertEqual(last.tag, 20)
        # `GetRecordsAround` fills two [out] records (`first` and `last`),
        # each with its own sentinel, so two sentinels are alive.  `middle` is
        # a plain LONG and owns no sentinel.
        self.assertEqual(obj.GetLiveSentinelCount(), 2)
        # Each record owns its sentinel, so deleting the records is expected to
        # release them.
        del first, last
        gc.collect()
        # CHARACTERIZATION: The sentinels are NOT released, count stays at 2.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            2,
            msg="CHARACTERIZATION: Expected sentinels to be released on del, "
            "but current behavior leaves them alive. This test records the "
            "current (buggy) behavior; it will fail when the bug is fixed.",
        )

    def test_inout_record_is_not_released_on_del(self):
        obj = _create_server()
        inout = ComtypesCppTestSrvLib.StructOutRecordTest()
        inout.tag = 5
        modified = obj.ModifyRecord(inout)
        self.assertEqual(modified.tag, 6)
        # `ModifyRecord` takes the record as [in, out].  The caller allocated the
        # record, so the caller owns it and is responsible for releasing its
        # contents.  The callee creates a new sentinel (reference count 1) and
        # stores it in the record, so one sentinel is alive.
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # Deleting the Python record does not release the sentinel, so the
        # count stays at 1.
        del modified
        gc.collect()
        self.assertEqual(obj.GetLiveSentinelCount(), 1)

    def test_fail_get_record_raises_and_creates_no_sentinel(self):
        obj = _create_server()
        # `FailGetRecord` returns `E_FAIL` without touching the [out] record, so
        # comtypes raises `COMError` and no sentinel is ever created.
        with self.assertRaises(COMError) as cm:
            obj.FailGetRecord()
        self.assertEqual(cm.exception.hresult, hresult.E_FAIL)
        self.assertEqual(obj.GetLiveSentinelCount(), 0)

    def test_from_buffer_copy_does_not_release_sentinel(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # The [out] record owns one sentinel (reference count 1).
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # `from_buffer_copy` makes a plain ctypes copy that does not own the
        # record's resources, so deleting the copy does not release the
        # sentinel.
        copy = ComtypesCppTestSrvLib.StructOutRecordTest.from_buffer_copy(rec)
        del copy
        gc.collect()
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # The original [out] record owns the sentinel, so deleting it is
        # expected to release the sentinel.
        del rec
        gc.collect()
        # CHARACTERIZATION: The sentinel is NOT released, count stays at 1.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            1,
            msg="CHARACTERIZATION: Expected sentinel to be released on del, "
            "but current behavior leaves it alive. This test records the "
            "current (buggy) behavior; it will fail when the bug is fixed.",
        )

    def test_reading_fields_then_del_releases_sentinel(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        self.assertEqual(rec.tag, 1)
        self.assertEqual(rec.text, "aaaa")
        unk = rec.unk
        self.assertTrue(unk)
        # Reading fields returns values without transferring ownership, so the
        # sentinel is still alive.
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        del unk
        del rec
        gc.collect()
        # CHARACTERIZATION: The sentinel IS released, count becomes 0.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            0,
            msg="CHARACTERIZATION: Expected sentinel to be released on del, "
            "and current behavior does release it. This test records the "
            "current behavior; it will fail if the behavior changes.",
        )

    def test_huge_bstr_is_released_without_leak(self):
        # Each `GetRecord` allocates a large `BSTR` that the record owns, so
        # deleting the record is expected to free it.  The current behavior
        # does not, so the process leaks memory.
        obj = _create_server()
        n = 100
        size = 2 * 1024 * 1024  # wchars -> 4 MiB per `BSTR`
        expected = n * size * 2

        def get_and_drop_record() -> None:
            rec = obj.GetRecord(size)
            del rec

        leaked = find_memleak(get_and_drop_record, loops=(1, n), target="private_usage")
        self.assertGreaterEqual(
            leaked,
            expected * 0.5,
            msg="CHARACTERIZATION: Expected large leak due to BSTR not being "
            "freed on record deletion. This test records the current (buggy) "
            "behavior; it will fail when the bug is fixed and the leak is small.",
        )

    def test_record_does_not_release_resources_on_del(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # The [out] record owns one sentinel (reference count 1).
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # The record owns its resources, so deleting the record is expected to
        # release them.
        del rec
        gc.collect()
        # CHARACTERIZATION: The sentinel is NOT released, count stays at 1.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            1,
            msg="CHARACTERIZATION: Expected sentinel to be released on del, "
            "but current behavior leaves it alive. This test records the "
            "current (buggy) behavior; it will fail when the bug is fixed.",
        )


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestRecordClear(unittest.TestCase):
    def test_explicit_record_clear_then_del_is_idempotent(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # The [out] record owns one sentinel (reference count 1).
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # `RecordClear` releases the sentinel, so it is destroyed.
        ri = _create_record_info()
        ri.RecordClear(addressof(rec))
        self.assertEqual(obj.GetLiveSentinelCount(), 0)
        # Deleting the record after an explicit `RecordClear` must not release
        # the sentinel a second time; `RecordClear` is idempotent (it zeroes
        # the pointer fields), so the count stays 0.
        del rec
        gc.collect()
        self.assertEqual(obj.GetLiveSentinelCount(), 0)

    def test_explicit_record_clear_releases_sentinel(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # The [out] record owns one sentinel (reference count 1).
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # Calling `RecordClear` directly releases the sentinel, so it is
        # destroyed.
        ri = _create_record_info()
        ri.RecordClear(pointer(rec))
        self.assertEqual(obj.GetLiveSentinelCount(), 0)

    def test_record_clear_zeroes_pointer_fields(self):
        # `RecordClear` releases the managed fields and zeroes their pointers,
        # which is what makes a second `RecordClear` a no-op.
        obj = _create_server()
        rec = obj.GetRecord(4)
        ri = _create_record_info()
        ri.RecordClear(byref(rec))
        raw = ctypes.string_at(addressof(rec), ctypes.sizeof(rec))
        wordsize = ctypes.sizeof(ctypes.c_void_p)
        for field in ("text", "unk"):
            offset = getattr(ComtypesCppTestSrvLib.StructOutRecordTest, field).offset
            word = int.from_bytes(raw[offset : offset + wordsize], "little")
            self.assertEqual(word, 0)

    def test_record_clear_raw_struct_raises_argumenterror(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        ri = _create_record_info()
        with self.assertRaises(
            ArgumentError,
            msg="CHARACTERIZATION: failure is expected; `RecordClear`'s "
            "`pvExisting` parameter is declared as `c_void_p`, and ctypes cannot "
            "convert a `StructOutRecordTest` instance to `c_void_p`, so passing "
            "the record object itself raises `ArgumentError` instead of clearing "
            "the record; the caller must pass the record's address "
            "(`addressof(rec)`, `byref(rec)` or `pointer(rec)`); once "
            "`RecordClear` accepts a record instance and takes its address "
            "automatically, this test will fail.",
        ) as cm:
            ri.RecordClear(rec)
        self.assertIn(
            "'StructOutRecordTest' object cannot be interpreted", str(cm.exception)
        )
        del rec
        gc.collect()


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestPutField(unittest.TestCase):
    def test_putref_unk_releases_old_sentinel(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # `GetRecord` gives the record one sentinel (reference count 1).
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # `GetSentinel` creates a second, independent sentinel (reference count
        # 1), so two sentinels are alive.
        s2 = obj.GetSentinel()
        self.assertEqual(obj.GetLiveSentinelCount(), 2)
        # `PutField` with `DISPATCH_PROPERTYPUTREF` releases the record's old
        # `unk` value before storing the new one, so the old sentinel is
        # destroyed and only `s2` remains alive.
        ri = _create_record_info()
        ri.PutField(DISPATCH_PROPERTYPUTREF, addressof(rec), "unk", s2)
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        del s2
        del rec
        gc.collect()

    def test_putref_copies_unk(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # `GetRecord` gives the record one sentinel; `GetSentinel` creates a
        # second, so two sentinels are alive.
        s2 = obj.GetSentinel()
        self.assertEqual(obj.GetLiveSentinelCount(), 2)
        # `PutField` with `DISPATCH_PROPERTYPUTREF` releases the record's old
        # sentinel and stores a reference to `s2`, so only `s2` remains alive.
        ri = _create_record_info()
        ri.PutField(DISPATCH_PROPERTYPUTREF, addressof(rec), "unk", s2)
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # Deleting the Python `s2` does not destroy the sentinel, because the
        # record still holds a reference to it.
        del s2
        gc.collect()
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        del rec
        gc.collect()

    def test_putref_new_unk_not_released_with_record(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        # `GetRecord` gives the record one sentinel; `GetSentinel` creates a
        # second, so two sentinels are alive.
        s2 = obj.GetSentinel()
        self.assertEqual(obj.GetLiveSentinelCount(), 2)
        # `PutField` with `DISPATCH_PROPERTYPUTREF` releases the record's old
        # sentinel and stores a reference to `s2`, so only `s2` remains alive.
        ri = _create_record_info()
        ri.PutField(DISPATCH_PROPERTYPUTREF, byref(rec), "unk", s2)
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # Deleting the Python `s2` leaves the sentinel alive because the record
        # still references it.
        del s2
        gc.collect()
        self.assertEqual(obj.GetLiveSentinelCount(), 1)
        # The record owns the stored sentinel, so deleting the record is
        # expected to release it.
        del rec
        gc.collect()
        # CHARACTERIZATION: The sentinel is NOT released, count stays at 1.
        self.assertEqual(
            obj.GetLiveSentinelCount(),
            1,
            msg="CHARACTERIZATION: Expected sentinel to be released on del, "
            "but current behavior leaves it alive. This test records the "
            "current (buggy) behavior; it will fail when the bug is fixed.",
        )

    def test_put_bstr_replaces_text(self):
        # `PutField` with `DISPATCH_PROPERTYPUT` replaces the `BSTR` field.  No
        # sentinel is involved, so the live count is not checked here.
        obj = _create_server()
        rec = obj.GetRecord(4)
        self.assertEqual(rec.text, "aaaa")
        ri = _create_record_info()
        ri.PutField(DISPATCH_PROPERTYPUT, pointer(rec), "text", "new text")
        self.assertEqual(rec.text, "new text")
        del rec
        gc.collect()

    def test_put_bstr_repeated_does_not_leak(self):
        # Repeatedly replacing the `BSTR` field must free the previous `BSTR`
        # each time; otherwise the process leaks memory.
        obj = _create_server()
        ri = _create_record_info()
        rec = obj.GetRecord(4)
        # Take the address once: the nested function below must write into the
        # same record on every call, and referring to `rec` there would be
        # flagged as an undefined name because `rec` is deleted afterwards.
        rec_addr = addressof(rec)
        n = 100
        size = 2 * 1024 * 1024  # wchars -> 4 MiB per `BSTR`
        expected = n * size * 2

        def put_repeatedly() -> None:
            v = VARIANT("x" * size)
            ri.PutField(DISPATCH_PROPERTYPUT, rec_addr, "text", v)
            del v

        leaked = find_memleak(put_repeatedly, loops=(1, n), target="private_usage")
        self.assertLessEqual(leaked, expected * 0.1)
        del rec
        gc.collect()

    def test_putref_bstr_raises_invalidarg(self):
        # `DISPATCH_PROPERTYPUTREF` is only valid for interface pointer fields;
        # using it on the `BSTR` field fails with `E_INVALIDARG`.
        obj = _create_server()
        rec = obj.GetRecord(4)
        ri = _create_record_info()
        with self.assertRaises(COMError) as cm:
            ri.PutField(DISPATCH_PROPERTYPUTREF, addressof(rec), "text", "new text")
        self.assertEqual(cm.exception.hresult, hresult.E_INVALIDARG)
        del rec
        gc.collect()

    # def test_direct_assignment_leaks_old_sentinel(self):
    #     """Assigning the `unk` field directly (instead of calling
    #     `IRecordInfo::PutField`) does not release the old sentinel and does
    #     not AddRef the new one.
    #
    #     `PutField` with `DISPATCH_PROPERTYPUTREF` releases the record's old
    #     `unk` value and AddRefs the new one, so the record owns the new
    #     sentinel.  A plain Python assignment (`rec.unk = s2`) is only a pointer
    #     store: it leaves the old sentinel alive (a leak) and stores the new
    #     pointer without AddRef, so the record holds a borrowed reference.
    #     """
    #     obj = _create_server()
    #     rec = obj.GetRecord(4)
    #     # `GetRecord` gives the record one sentinel (reference count 1).
    #     self.assertEqual(obj.GetLiveSentinelCount(), 1)
    #     s2 = obj.GetSentinel()
    #     self.assertEqual(obj.GetLiveSentinelCount(), 2)
    #     # Direct assignment does not release the old sentinel: the count stays
    #     # at 2 instead of dropping to 1 as it does with `PutField`.
    #     rec.unk = s2
    #     self.assertEqual(obj.GetLiveSentinelCount(), 2)
    #     # Releasing the record releases the *new* sentinel (the value stored in
    #     # the field), not the old one.  The old sentinel leaks.
    #     ri = _create_record_info()
    #     ri.RecordClear(addressof(rec))
    #     self.assertEqual(obj.GetLiveSentinelCount(), 1)
    #     # `s2` still points to the sentinel that `RecordClear` just released,
    #     # so deleting `s2` releases it a second time.  The count stays at 1
    #     # because the old sentinel is still leaked.
    #     del s2
    #     gc.collect()
    #     self.assertEqual(obj.GetLiveSentinelCount(), 1)
    #     del rec
    #     gc.collect()
    #     self.assertEqual(obj.GetLiveSentinelCount(), 1)

    def test_putref_raw_struct_raises_argumenterror(self):
        obj = _create_server()
        rec = obj.GetRecord(4)
        ri = _create_record_info()
        with self.assertRaises(
            ArgumentError,
            msg="CHARACTERIZATION: failure is expected; `PutField`'s `pvData` "
            "parameter is declared as `c_void_p`, and ctypes cannot convert a "
            "`StructOutRecordTest` instance to `c_void_p`, so passing the record "
            "object itself raises `ArgumentError` instead of writing the field; "
            "the caller must pass the record's address (`addressof(rec)`, "
            "`byref(rec)` or `pointer(rec)`); once `PutField` accepts a record "
            "instance and takes its address automatically, this test will fail.",
        ) as cm:
            ri.PutField(DISPATCH_PROPERTYPUT, rec, "text", "new text")
        self.assertIn(
            "'StructOutRecordTest' object cannot be interpreted", str(cm.exception)
        )
        del rec
        gc.collect()


if __name__ == "__main__":
    unittest.main()
