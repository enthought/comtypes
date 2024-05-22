# coding: utf-8

import unittest

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from ctypes import byref, pointer

ComtypesTestLib_GUID = "07D2AEE5-1DF8-4D2C-953A-554ADFD25F99"
ProgID = "ComtypesTest.COM.Server"

try:
    GetModule([f"{{{ComtypesTestLib_GUID}}}", 1, 0, 0])
    import comtypes.gen.ComtypesTestLib as ComtypesTestLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class Test(unittest.TestCase):
    """Test dispmethods with record pointer parameters."""

    def test(self):
        # Explicitely ask for the dispinterface of the COM-server.
        dispifc = CreateObject(
            ProgID, clsctx=CLSCTX_LOCAL_SERVER, interface=ComtypesTestLib.IComtypesTest
        )

        # Test passing a record by reference.
        test_record = ComtypesTestLib.T_TEST_RECORD()
        self.assertEqual(test_record.question, None)
        self.assertEqual(test_record.answer, 0)
        self.assertEqual(test_record.needs_clarification, False)
        dispifc.InitRecord(byref(test_record))
        self.assertEqual(
            test_record.question, "The meaning of life, the universe and everything?"
        )
        self.assertEqual(test_record.answer, 42)
        self.assertEqual(test_record.needs_clarification, True)

        # Test passing a record pointer.
        test_record = ComtypesTestLib.T_TEST_RECORD()
        self.assertEqual(test_record.question, None)
        self.assertEqual(test_record.answer, 0)
        self.assertEqual(test_record.needs_clarification, False)
        test_record_pointer = pointer(test_record)
        dispifc.InitRecord(test_record_pointer)
        self.assertEqual(
            test_record.question, "The meaning of life, the universe and everything?"
        )
        self.assertEqual(test_record.answer, 42)
        self.assertEqual(test_record.needs_clarification, True)


if __name__ == "__main__":
    unittest.main()
