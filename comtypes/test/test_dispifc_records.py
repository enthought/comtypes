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
        # Explicitely ask for the dispinterface of the component.
        dispifc = CreateObject(
            ProgID, clsctx=CLSCTX_LOCAL_SERVER, interface=ComtypesTestLib.IComtypesTest
        )

        # Passing a record by reference to a method that has declared the parameter
        # as [in, out] we expect modifications of the record on the server side to
        # also change the record on the client side.
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

        # Passing a record pointer to a method that has declared the parameter
        # as [in, out] we expect modifications of the record on the server side to
        # also change the record on the client side.
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

        # Passing a record to a method that has declared the parameter just as [in]
        # we expect modifications of the record on the server side NOT to change
        # the record on the client side.
        # We also need to test if the record gets properly passed to the method on
        # the server side. For this, the 'VerifyRecord' method returns 'True' if
        # all record fields have the initialization values provided by 'InitRecord'.
        self.assertTrue(dispifc.VerifyRecord(test_record))
        # Check if the 'answer' field is unchanged although the method modifies this
        # field on the server side.
        self.assertEqual(test_record.answer, 42)
        # Also perform the inverted test.
        # For this, first create a blank record.
        test_record = ComtypesTestLib.T_TEST_RECORD()
        self.assertEqual(test_record.question, None)
        self.assertEqual(test_record.answer, 0)
        self.assertEqual(test_record.needs_clarification, False)
        # Perform the check on initialization values.
        self.assertFalse(dispifc.VerifyRecord(test_record))
        # The record on the client side should be unchanged.
        self.assertEqual(test_record.answer, 0)


if __name__ == "__main__":
    unittest.main()
