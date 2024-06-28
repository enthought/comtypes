# coding: utf-8

from ctypes import byref, c_int, pointer, POINTER
import unittest

import comtypes
import comtypes.safearray
from comtypes import CLSCTX_INPROC_SERVER
from comtypes.client import CreateObject, GetModule
import comtypes.typeinfo

GetModule("UIAutomationCore.dll")
from comtypes.gen.UIAutomationClient import CUIAutomation, IUIAutomation

GetModule("scrrun.dll")
from comtypes.gen.Scripting import Dictionary, IDictionary

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import StructRecordParamTest

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


class Test_midlSAFEARRAY_create(unittest.TestCase):
    def test_iunk(self):
        extra = pointer(IUIAutomation._iid_)
        iuia = CreateObject(
            CUIAutomation().IPersist_GetClassID(),
            interface=IUIAutomation,
            clsctx=CLSCTX_INPROC_SERVER,
        )
        sa_type = comtypes.safearray._midlSAFEARRAY(POINTER(IUIAutomation))
        for ptn, sa in [
            ("with extra", sa_type.create([iuia], extra=extra)),
            ("without extra", sa_type.create([iuia])),
        ]:
            with self.subTest(ptn=ptn):
                (unpacked,) = sa.unpack()
                self.assertIsInstance(unpacked, POINTER(IUIAutomation))

    def test_idisp(self):
        extra = pointer(IDictionary._iid_)
        idic = CreateObject(Dictionary, interface=IDictionary)
        idic["foo"] = "bar"
        sa_type = comtypes.safearray._midlSAFEARRAY(POINTER(IDictionary))
        for ptn, sa in [
            ("with extra", sa_type.create([idic], extra=extra)),
            ("without extra", sa_type.create([idic])),
        ]:
            with self.subTest(ptn=ptn):
                (unpacked,) = sa.unpack()
                self.assertIsInstance(unpacked, POINTER(IDictionary))
                self.assertEqual(unpacked["foo"], "bar")

    @unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
    def test_record(self):
        extra = comtypes.typeinfo.GetRecordInfoFromGuids(
            *StructRecordParamTest._recordinfo_
        )
        record = StructRecordParamTest()
        record.answer = 42
        sa_type = comtypes.safearray._midlSAFEARRAY(StructRecordParamTest)
        for ptn, sa in [
            ("with extra", sa_type.create([record], extra=extra)),
            ("without extra", sa_type.create([record])),
        ]:
            with self.subTest(ptn=ptn):
                (unpacked,) = sa.unpack()
                self.assertIsInstance(unpacked, StructRecordParamTest)
                self.assertEqual(unpacked.answer, 42)

    def test_ctype(self):
        extra = None
        cdata = c_int(1)
        sa_type = comtypes.safearray._midlSAFEARRAY(c_int)
        for ptn, sa in [
            ("with extra", sa_type.create([cdata], extra=extra)),
            ("without extra", sa_type.create([cdata])),
        ]:
            with self.subTest(ptn=ptn):
                (unpacked,) = sa.unpack()
                self.assertIsInstance(unpacked, int)
                self.assertEqual(unpacked, 1)


if __name__ == "__main__":
    unittest.main()
