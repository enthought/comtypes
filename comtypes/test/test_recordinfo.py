import unittest
from ctypes import Structure, Union, byref, pointer, sizeof

from comtypes import GUID, typeinfo
from comtypes.client import GetModule
from comtypes.recordinfo import SupportsRecordInfo

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import StructRecordParamTest

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


class SupportsRecordInfoTest(unittest.TestCase):
    class Valid(Structure):
        _recordinfo_ = (str(GUID.create_new()), 1, 0, 0, str(GUID.create_new()))

    class Missing(Structure): ...

    class WithNone(Structure):
        _recordinfo_ = None

    class WithEmptyTuple(Structure):
        _recordinfo_ = ()

    class WithFourElements(Structure):
        _recordinfo_ = (str(GUID.create_new()), 1, 0, 0)

    class WithSixElements(Structure):
        _recordinfo_ = (str(GUID.create_new()), 1, 0, 0, str(GUID.create_new()), 0)

    class WithFiveButWrongElement(Structure):
        _recordinfo_ = (str(GUID.create_new()), 1, 0, 0, 0)

    class NotAStructureButUnion(Union):
        _recordinfo_ = (str(GUID.create_new()), 1, 0, 0, str(GUID.create_new()))

    class NotAStructureButPlain:
        _recordinfo_ = (str(GUID.create_new()), 1, 0, 0, str(GUID.create_new()))

    def test_valid(self):
        self.assertTrue(issubclass(self.Valid, SupportsRecordInfo))
        self.assertIsInstance(self.Valid(), SupportsRecordInfo)

    def test_invalid_structures(self):
        for typ in (
            self.Missing,
            self.WithNone,
            self.WithEmptyTuple,
            self.WithFourElements,
            self.WithSixElements,
            self.WithFiveButWrongElement,
        ):
            with self.subTest(typ.__name__):
                self.assertNotIsSubclass(typ, SupportsRecordInfo)
                self.assertNotIsInstance(typ(), SupportsRecordInfo)

    def test_not_a_structure(self):
        for typ in (self.NotAStructureButUnion, self.NotAStructureButPlain):
            with self.subTest(typ.__name__):
                self.assertNotIsSubclass(typ, SupportsRecordInfo)
                self.assertNotIsInstance(typ(), SupportsRecordInfo)

    def test_not_a_class(self):
        self.assertNotIsSubclass(42, SupportsRecordInfo)  # type: ignore

    def test_instance_attribute_is_ignored(self):
        obj = self.Missing()
        obj._recordinfo_ = (str(GUID.create_new()), 1, 0, 0, str(GUID.create_new()))
        self.assertNotIsInstance(obj, SupportsRecordInfo)

    def test_cannot_subclass(self):
        with self.assertRaises(TypeError):

            class BadSub(SupportsRecordInfo): ...  # type: ignore

        with self.assertRaises(TypeError):

            class BadMeta(type(SupportsRecordInfo)): ...

    def test_cannot_instantiated(self):
        with self.assertRaises(TypeError):
            SupportsRecordInfo()


def _create_recordinfo() -> typeinfo.IRecordInfo:
    return typeinfo.GetRecordInfoFromGuids(*StructRecordParamTest._recordinfo_)


def _create_record(
    question: str, answer: int, needs_clarification: bool
) -> "StructRecordParamTest":
    record = StructRecordParamTest()
    record.question = question
    record.answer = answer
    record.needs_clarification = needs_clarification
    return record


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class Test_IRecordInfo(unittest.TestCase):
    def test_RecordCopy(self):
        dst_rec = StructRecordParamTest()
        ri = _create_recordinfo()
        ri.RecordCopy(pointer(_create_record("foo", 3, True)), byref(dst_rec))
        self.assertEqual(dst_rec.question, "foo")
        self.assertEqual(dst_rec.answer, 3)
        self.assertEqual(dst_rec.needs_clarification, True)

    def test_GetGuid(self):
        *_, expected_guid = StructRecordParamTest._recordinfo_
        self.assertEqual(str(_create_recordinfo().GetGuid()), expected_guid)

    def test_GetName(self):
        self.assertEqual(
            _create_recordinfo().GetName(),
            StructRecordParamTest.__qualname__,
        )

    def test_GetSize(self):
        self.assertEqual(
            _create_recordinfo().GetSize(),
            sizeof(StructRecordParamTest),
        )

    def test_GetTypeInfo(self):
        ta = _create_recordinfo().GetTypeInfo().GetTypeAttr()
        *_, expected_guid = StructRecordParamTest._recordinfo_
        self.assertEqual(str(ta.guid), expected_guid)

    def test_IsMatchingType(self):
        ri = _create_recordinfo()
        ti = ri.GetTypeInfo()
        self.assertTrue(typeinfo.GetRecordInfoFromTypeInfo(ti).IsMatchingType(ri))

    def test_RecordCreateCopy(self):
        ri = _create_recordinfo()
        actual = ri.RecordCreateCopy(byref(_create_record("foo", 3, True)))
        self.assertIsInstance(actual, int)
        dst_rec = StructRecordParamTest()
        ri.RecordCopy(actual, byref(dst_rec))
        ri.RecordDestroy(actual)
        self.assertEqual(dst_rec.question, "foo")
        self.assertEqual(dst_rec.answer, 3)
        self.assertEqual(dst_rec.needs_clarification, True)
