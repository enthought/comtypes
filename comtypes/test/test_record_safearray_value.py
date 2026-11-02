"""Record SAFEARRAY Tests (Value Records)

Tests record marshaling of value records (integer-only structures) between
Python and a COM server.
"""

import gc
import unittest
from _ctypes import COMError
from typing import TYPE_CHECKING, ClassVar, TypeVar
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
    # Server type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


def _create_record(
    id: int, seq: int, checksum: int
) -> "ComtypesCppTestSrvLib.StructValueRecordSafearrayTest":
    return ComtypesCppTestSrvLib.StructValueRecordSafearrayTest(
        id=id, seq=seq, checksum=checksum
    )


_T_IUnknown = TypeVar("_T_IUnknown", bound=IUnknown)


def _create_server(interface: type[_T_IUnknown]) -> _T_IUnknown:
    """Create an out-of-process COM server instance for the given interface.

    Uses CLSCTX_LOCAL_SERVER to launch CoRecordSafearrayTest as a separate
    process, ensuring proper marshaling of SAFEARRAY across process boundaries.
    """
    return CreateObject(
        "Comtypes.RecordSafeArrayRoundTripTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=interface,
    )


def _get_value_record_sa_type():
    """Return the SAFEARRAY type for value records (the _midlSAFEARRAY specialization).

    This is a function (not a module-level constant) to defer evaluation until
    the type library has been imported. This allows the module to be imported
    even when the test server is not registered.
    """
    return _midlSAFEARRAY(ComtypesCppTestSrvLib.StructValueRecordSafearrayTest)


class _ValueRecordScenariosMixin(_TCBase):
    """Mixin providing test scenarios for value records
    (StructValueRecordSafearrayTest).
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
            _create_record(id=200, seq=0, checksum=20042),
        ]
        lp = _get_value_record_sa_type().from_param(records)

        result = ifc.FirstValueRecord(lp)
        self.assertEqual(result.id, 200)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20042)

        del records, lp
        gc.collect()

    def test_first_record_length_1_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=200, seq=0, checksum=20042),
        ]
        lp = _get_value_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        result = ifc.FirstValueRecord(lp)
        self.assertEqual(result.id, 200)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20042)

        del records
        gc.collect()

    def test_first_record_length_3_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=201, seq=0, checksum=20142),
            _create_record(id=202, seq=1, checksum=20252),
            _create_record(id=203, seq=2, checksum=20362),
        ]
        lp = _get_value_record_sa_type().from_param(records)

        result = ifc.FirstValueRecord(lp)
        self.assertEqual(result.id, 201)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20142)

        del records, lp
        gc.collect()

    # Server returns the first record from a length-3 array (create LP).
    def test_first_record_length_3_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=201, seq=0, checksum=20142),
            _create_record(id=202, seq=1, checksum=20252),
            _create_record(id=203, seq=2, checksum=20362),
        ]
        lp = _get_value_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        result = ifc.FirstValueRecord(lp)
        self.assertEqual(result.id, 201)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20142)

        del records
        gc.collect()

    def test_first_record_survives_input_release(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=202, seq=0, checksum=20242),
            _create_record(id=203, seq=1, checksum=20352),
            _create_record(id=204, seq=2, checksum=20462),
        ]
        lp = _get_value_record_sa_type().from_param(records)

        result = ifc.FirstValueRecord(lp)

        # Release the input array
        del lp
        gc.collect()

        # Result should still be valid
        self.assertEqual(result.id, 202)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20242)

        del records
        gc.collect()


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestVtableValueRecordSafeArray(_ValueRecordScenariosMixin, unittest.TestCase):
    """IDispatch dual vtable tests, value records."""

    @classmethod
    def setUpClass(cls):
        cls.interface = ComtypesCppTestSrvLib.IDualRecordSafearrayTest

    def test_wrap_record(self):
        ifc = _create_server(self.interface)
        record = _create_record(id=300, seq=0, checksum=30042)

        result = ifc.WrapValueRecord(record)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].id, 300)
        self.assertEqual(result[0].seq, 0)
        self.assertEqual(result[0].checksum, 30042)

        del record
        gc.collect()

    def test_first_record_length_1_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=200, seq=0, checksum=20042),
        ]

        result = ifc.FirstValueRecord(records)
        self.assertEqual(result.id, 200)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20042)

        del records
        gc.collect()

    def test_first_record_length_3_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=201, seq=0, checksum=20142),
            _create_record(id=202, seq=1, checksum=20252),
            _create_record(id=203, seq=2, checksum=20362),
        ]

        result = ifc.FirstValueRecord(records)
        self.assertEqual(result.id, 201)
        self.assertEqual(result.seq, 0)
        self.assertEqual(result.checksum, 20142)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_value_record_sa_type().from_param(records)

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(lp)
        self.assertEqual(cm.exception.hresult, E_INVALIDARG)

        del records, lp
        gc.collect()

    def test_first_record_empty_list(self):
        ifc = _create_server(self.interface)
        records = []

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(records)
        self.assertEqual(cm.exception.hresult, E_INVALIDARG)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_value_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(lp)
        self.assertEqual(cm.exception.hresult, E_INVALIDARG)

        del records
        gc.collect()

    def test_reverse_records_length_3(self):
        ifc = _create_server(self.interface)

        records = [
            _create_record(id=201, seq=0, checksum=20142),
            _create_record(id=202, seq=1, checksum=20252),
            _create_record(id=203, seq=2, checksum=20362),
        ]
        lp = _get_value_record_sa_type().from_param(records)

        result = ifc.ReverseValueRecords(lp)
        self.assertEqual(len(result), 3)
        self.assertEqual(
            [(r.id, r.seq, r.checksum) for r in result],
            [(s.id, s.seq, s.checksum) for s in reversed(records)],
        )

        del records
        gc.collect()


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestDispatchValueRecordSafeArray(_ValueRecordScenariosMixin, unittest.TestCase):
    """IDispatch dispinterface Invoke tests, value records."""

    @classmethod
    def setUpClass(cls):
        cls.interface = ComtypesCppTestSrvLib.IDispRecordSafearrayTest

    def test_wrap_record(self):
        ifc = _create_server(self.interface)
        record = _create_record(id=300, seq=0, checksum=30042)

        with self.assertRaises(KeyError) as cm:
            ifc.WrapValueRecord(record)
        self.assertEqual(str(cm.exception), "36")

        del record
        gc.collect()

    def test_first_record_length_1_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=200, seq=0, checksum=20042),
        ]

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(records)
        self.assertEqual(cm.exception.hresult, DISP_E_TYPEMISMATCH)

        del records
        gc.collect()

    def test_first_record_length_3_list(self):
        ifc = _create_server(self.interface)
        records = [
            _create_record(id=201, seq=0, checksum=20142),
            _create_record(id=202, seq=1, checksum=20252),
            _create_record(id=203, seq=2, checksum=20362),
        ]
        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(records)
        self.assertEqual(cm.exception.hresult, DISP_E_TYPEMISMATCH)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_from_param(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_value_record_sa_type().from_param(records)

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(lp)
        # wrapped in DISP_E_EXCEPTION with scode=E_INVALIDARG.
        self.assertEqual(cm.exception.hresult, DISP_E_EXCEPTION)
        self.assertEqual(cm.exception.details[4], E_INVALIDARG)

        del records, lp
        gc.collect()

    def test_first_record_empty_list(self):
        ifc = _create_server(self.interface)
        records = []

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(records)
        self.assertEqual(cm.exception.hresult, DISP_E_TYPEMISMATCH)

        del records
        gc.collect()

    def test_first_record_empty_safearray_by_create(self):
        ifc = _create_server(self.interface)
        records = []
        lp = _get_value_record_sa_type().create(records)
        self.addCleanup(_safearray.SafeArrayDestroy, lp)

        with self.assertRaises(COMError) as cm:
            ifc.FirstValueRecord(lp)
        # wrapped in DISP_E_EXCEPTION with scode=E_INVALIDARG.
        self.assertEqual(cm.exception.hresult, DISP_E_EXCEPTION)
        self.assertEqual(cm.exception.details[4], E_INVALIDARG)

        del records
        gc.collect()

    def test_reverse_records_length_3(self):
        ifc = _create_server(self.interface)

        records = [
            _create_record(id=201, seq=0, checksum=20142),
            _create_record(id=202, seq=1, checksum=20252),
            _create_record(id=203, seq=2, checksum=20362),
        ]
        lp = _get_value_record_sa_type().from_param(records)

        with self.assertRaises(KeyError) as cm:
            ifc.ReverseValueRecords(lp)
        self.assertEqual(str(cm.exception), "36")

        del records
        gc.collect()


if __name__ == "__main__":
    unittest.main()
