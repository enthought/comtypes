"""Record SAFEARRAY Tests (BSTR Records)

Tests record marshaling of BSTR records (structures owning heap-allocated
strings) between Python and a COM server.
"""

import faulthandler
import gc
import io
import multiprocessing
import re
import sys
import tempfile
import unittest
from _ctypes import COMError
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any, ClassVar, Optional, TypeVar
from typing import Union as _UnionT

from comtypes import CLSCTX_LOCAL_SERVER, IUnknown, _safearray
from comtypes.client import CreateObject, GetModule
from comtypes.hresult import DISP_E_EXCEPTION, DISP_E_TYPEMISMATCH, E_INVALIDARG
from comtypes.safearray import _midlSAFEARRAY

# _TCBase is unittest.TestCase during type checking (so mixin methods can use
# self.assertEqual etc.) but object at runtime. This prevents the mixin class
# itself from being collected as a test case by unittest.
if TYPE_CHECKING:
    _TCBase = unittest.TestCase
else:
    _TCBase = object

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


def _create_record(
    id: int, seq: int, text: str, checksum: int
) -> "ComtypesCppTestSrvLib.StructBstrRecordSafearrayTest":
    return ComtypesCppTestSrvLib.StructBstrRecordSafearrayTest(
        id=id, seq=seq, text=text, checksum=checksum
    )


_T_IUnknown = TypeVar("_T_IUnknown", bound=IUnknown)


def _create_server(interface: type[_T_IUnknown]) -> _T_IUnknown:
    return CreateObject(
        "Comtypes.RecordSafeArrayRoundTripTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=interface,
    )


def _get_bstr_record_sa_type():
    """Return the SAFEARRAY type for BSTR records"""
    return _midlSAFEARRAY(ComtypesCppTestSrvLib.StructBstrRecordSafearrayTest)


class _BstrRecordScenariosMixin(_TCBase):
    """Mixin providing test scenarios for BSTR records
    (StructBstrRecordSafearrayTest).
    """

    interface: ClassVar[
        _UnionT[
            type["ComtypesCppTestSrvLib.IDualRecordSafearrayTest"],
            type["ComtypesCppTestSrvLib.IDispRecordSafearrayTest"],
        ]
    ]

    def test_first_record_length_1_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=400, seq=0, text="RecSafeArray-400-0", checksum=40708),
        ]
        lp = _get_bstr_record_sa_type().from_param(records)

        result = ifc.FirstBstrRecord(lp)
        self.assertEqual(result.id, 400)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-400-0")
        self.assertEqual(result.checksum, 40708)

        del records, lp
        gc.collect()

    def test_first_record_length_1_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=400, seq=0, text="RecSafeArray-400-0", checksum=40708),
        ]
        lp = _get_bstr_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        result = ifc.FirstBstrRecord(lp)
        self.assertEqual(result.id, 400)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-400-0")
        self.assertEqual(result.checksum, 40708)

        del records
        gc.collect()

    def test_first_record_length_3_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
            _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
            _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
        ]
        lp = _get_bstr_record_sa_type().from_param(records)

        result = ifc.FirstBstrRecord(lp)
        self.assertEqual(result.id, 401)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-401-0")
        self.assertEqual(result.checksum, 40808)

        del records, lp
        gc.collect()

    def test_first_record_length_3_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
            _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
            _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
        ]
        lp = _get_bstr_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        result = ifc.FirstBstrRecord(lp)
        self.assertEqual(result.id, 401)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-401-0")
        self.assertEqual(result.checksum, 40808)

        del records
        gc.collect()

    def test_first_record_survives_input_release(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=402, seq=0, text="RecSafeArray-402-0", checksum=40908),
            _create_record(id=403, seq=1, text="RecSafeArray-403-1", checksum=41018),
            _create_record(id=404, seq=2, text="RecSafeArray-404-2", checksum=41128),
        ]
        lp = _get_bstr_record_sa_type().from_param(records)

        result = ifc.FirstBstrRecord(lp)

        del lp  # Release the input array
        gc.collect()

        # Result should still be valid after the input SAFEARRAY is released
        self.assertEqual(result.id, 402)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-402-0")
        self.assertEqual(result.checksum, 40908)

        del records
        gc.collect()


# ======================================================================
# Vtable pathway tests (IDualRecordSafearrayTest)
# ======================================================================


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestVtableBstrRecordSafeArray(_BstrRecordScenariosMixin, unittest.TestCase):
    """IDispatch dual vtable tests, BSTR records."""

    @classmethod
    def setUpClass(cls):
        cls.interface = ComtypesCppTestSrvLib.IDualRecordSafearrayTest

    def test_wrap_record(self):
        ifc = _create_server(self.interface)
        record = _create_record(id=500, seq=0, text="WrapBstr-500-0", checksum=50708)

        result = ifc.WrapBstrRecord(record)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].id, 500)
        self.assertEqual(result[0].seq, 0)
        self.assertEqual(str(result[0].text), "WrapBstr-500-0")
        self.assertEqual(result[0].checksum, 50708)

        del record
        gc.collect()

    def test_first_record_length_1_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=400, seq=0, text="RecSafeArray-400-0", checksum=40708),
        ]

        result = ifc.FirstBstrRecord(records)
        self.assertEqual(result.id, 400)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-400-0")
        self.assertEqual(result.checksum, 40708)

        del records
        gc.collect()

    def test_first_record_length_3_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
            _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
            _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
        ]

        result = ifc.FirstBstrRecord(records)
        self.assertEqual(result.id, 401)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.text, "RecSafeArray-401-0")
        self.assertEqual(result.checksum, 40808)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_bstr_record_sa_type().from_param(records)

        with self.assertRaises(COMError) as cm:
            ifc.FirstBstrRecord(lp)
        self.assertEqual(cm.exception.hresult, E_INVALIDARG)

        del records, lp
        gc.collect()

    def test_first_record_empty_list(self):
        ifc = _create_server(self.interface)
        records = []

        with self.assertRaises(COMError) as cm:
            ifc.FirstBstrRecord(records)
        self.assertEqual(cm.exception.hresult, E_INVALIDARG)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_bstr_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        with self.assertRaises(COMError) as cm:
            ifc.FirstBstrRecord(lp)
        self.assertEqual(cm.exception.hresult, E_INVALIDARG)

        del records
        gc.collect()

    def test_reverse_records_length_3(self):
        ifc = _create_server(self.interface)

        records = [
            _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
            _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
            _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
        ]
        lp = _get_bstr_record_sa_type().from_param(records)

        result = ifc.ReverseBstrRecords(lp)
        self.assertEqual(len(result), 3)
        self.assertEqual(
            [(r.id, r.seq, r.text, r.checksum) for r in result],
            [(s.id, s.seq, s.text, s.checksum) for s in reversed(records)],
        )

        del records
        gc.collect()


# ======================================================================
# Dispatch pathway tests (IDispRecordSafearrayTest)
# ======================================================================


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDispatchBstrRecordSafeArray(_BstrRecordScenariosMixin, unittest.TestCase):
    """IDispatch dispinterface Invoke tests, BSTR records."""

    @classmethod
    def setUpClass(cls):
        cls.interface = ComtypesCppTestSrvLib.IDispRecordSafearrayTest

    def test_wrap_record(self):
        ifc = _create_server(self.interface)
        record = _create_record(id=500, seq=0, text="WrapBstr-500-0", checksum=50708)

        result = ifc.WrapBstrRecord(record)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].id, 500)
        self.assertEqual(result[0].seq, 0)
        self.assertEqual(str(result[0].text), "WrapBstr-500-0")
        self.assertEqual(result[0].checksum, 50708)

        del record
        gc.collect()

    def test_first_record_length_1_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=400, seq=0, text="RecSafeArray-400-0", checksum=40708),
        ]

        with self.assertRaises(
            COMError,
            msg="CHARACTERIZATION(#63): "
            "If https://github.com/enthought/comtypes/issues/63 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ) as cm:
            ifc.FirstBstrRecord(records)
        self.assertEqual(cm.exception.hresult, DISP_E_TYPEMISMATCH)

        del records
        gc.collect()

    def test_first_record_length_3_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
            _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
            _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
        ]

        with self.assertRaises(
            COMError,
            msg="CHARACTERIZATION(#63): "
            "If https://github.com/enthought/comtypes/issues/63 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ) as cm:
            ifc.FirstBstrRecord(records)
        self.assertEqual(cm.exception.hresult, DISP_E_TYPEMISMATCH)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_bstr_record_sa_type().from_param(records)

        with self.assertRaises(COMError) as cm:
            ifc.FirstBstrRecord(lp)
        # wrapped in DISP_E_EXCEPTION with scode=E_INVALIDARG.
        self.assertEqual(cm.exception.hresult, DISP_E_EXCEPTION)
        self.assertEqual(cm.exception.details[4], E_INVALIDARG)

        del records, lp
        gc.collect()

    def test_first_record_empty_safearray_list(self):
        ifc = _create_server(self.interface)
        records = []

        with self.assertRaises(COMError) as cm:
            ifc.FirstBstrRecord(records)
        self.assertEqual(cm.exception.hresult, DISP_E_TYPEMISMATCH)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_bstr_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        with self.assertRaises(COMError) as cm:
            ifc.FirstBstrRecord(lp)
        # wrapped in DISP_E_EXCEPTION with scode=E_INVALIDARG.
        self.assertEqual(cm.exception.hresult, DISP_E_EXCEPTION)
        self.assertEqual(cm.exception.details[4], E_INVALIDARG)

        del records
        gc.collect()

    @unittest.skipIf(sys.version_info >= (3, 14), "CRASH! CHARACTERIZATION(#63)")
    def test_reverse_records_length_3(self):
        ifc = _create_server(self.interface)

        records = [
            _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
            _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
            _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
        ]
        lp = _get_bstr_record_sa_type().from_param(records)

        result = ifc.ReverseBstrRecords(lp)
        self.assertEqual(len(result), 3)
        self.assertEqual(
            [(r.id, r.seq, r.text, r.checksum) for r in result],
            [(s.id, s.seq, s.text, s.checksum) for s in reversed(records)],
        )

        del records
        gc.collect()


def enable_fault_log(tmpfile: Path) -> io.TextIOWrapper:
    """Send faulthandler output to `tmpfile` until the process exits."""
    f = tmpfile.open(mode="w", encoding="utf-8")
    # We keep `faulthandler` enabled and the file open until the process
    # exits so that crash information can be captured.
    # It is only enabled in the child process and does not affect the main
    # process, so there is no need to worry about forgetting to close and
    # disable.
    faulthandler.enable(file=f)
    return f


def run_in_mp(fn: Callable[[Path], Any]) -> tuple[Optional[int], str]:
    """Run a function in a separate process and return exit code and output.

    Using spawn isolates the process so crashes like segfaults do not affect
    the main test process.
    """
    with tempfile.TemporaryDirectory() as td:
        tmpdir = Path(td)
        tmpfile = tmpdir / "fault.txt"
        ctx = multiprocessing.get_context("spawn")
        process = ctx.Process(target=fn, args=(tmpfile,))
        process.start()
        process.join()
        output = tmpfile.read_text()
    return process.exitcode, output


def call_reverse_bstr_records(tmpfile: Path) -> None:
    enable_fault_log(tmpfile)
    ifc = _create_server(ComtypesCppTestSrvLib.IDispRecordSafearrayTest)

    records = [
        _create_record(id=401, seq=0, text="RecSafeArray-401-0", checksum=40808),
        _create_record(id=402, seq=1, text="RecSafeArray-402-1", checksum=40918),
        _create_record(id=403, seq=2, text="RecSafeArray-403-2", checksum=41028),
    ]
    lp = _get_bstr_record_sa_type().from_param(records)

    result = ifc.ReverseBstrRecords(lp)

    del records
    del result
    del lp
    gc.collect()


# Windows NTSTATUS values (unsigned 32-bit).
# Python's process.exitcode may return these as unsigned integers
# (e.g., 3221226356 = 0xC0000374) depending on the platform/Python version.
# We mask with 0xFFFFFFFF to ensure consistent comparison.
STATUS_HEAP_CORRUPTION = 0xC0000374
STATUS_ACCESS_VIOLATION = 0xC0000005
STATUS_STACK_BUFFER_OVERRUN = 0xC0000409


class TestOwnership(unittest.TestCase):
    @unittest.skipIf(sys.version_info < (3, 14), "NOT CRASH! CHARACTERIZATION(#63)")
    def test_bstr_record_safearray_ownership(self):
        exitcode, output = run_in_mp(call_reverse_bstr_records)
        self.assertIsNotNone(exitcode)
        # Python's process.exitcode on Windows may return the exit code as
        # either a signed or unsigned 32-bit integer. Masking with 0xFFFFFFFF
        # normalizes it to the unsigned 32-bit representation for comparison
        # with the NTSTATUS constants defined above.
        self.assertEqual(
            exitcode & 0xFFFFFFFF,  # type: ignore
            STATUS_HEAP_CORRUPTION,
            msg="CHARACTERIZATION(#63): "
            "If https://github.com/enthought/comtypes/issues/63 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        )
        ptn = re.compile(
            r'test_record_safearray_bstr\.py", line \d+ in call_reverse_bstr_records'
        )
        self.assertTrue(bool(ptn.search(output)))


if __name__ == "__main__":
    unittest.main()
