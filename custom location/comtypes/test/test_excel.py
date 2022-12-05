# -*- coding: latin-1 -*-
from __future__ import print_function

import datetime
import unittest

from comtypes.client import CreateObject, GetModule

################################################################
#
# TODO:
#
# It seems bad that only external test like this
# can verify the behavior of `comtypes` implementation.
# Find a different built-in win32 API to use.
#
################################################################

try:
    GetModule(("{00020813-0000-0000-C000-000000000046}",))  # Excel libUUID
    from comtypes.gen.Excel import xlRangeValueDefault
    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


class BaseBindTest(object):
    # `dynamic = True/False` must be defined in subclasses!

    def setUp(self):
        self.xl = CreateObject("Excel.Application", dynamic=self.dynamic)

    def tearDown(self):
        # Close all open workbooks without saving, then quit excel.
        for wb in self.xl.Workbooks:
            wb.Close(0)
        self.xl.Quit()
        del self.xl

    def test(self):
        xl = self.xl
        xl.Visible = 0
        self.assertEqual(xl.Visible, False)  # type: ignore
        xl.Visible = 1
        self.assertEqual(xl.Visible, True)  # type: ignore

        wb = xl.Workbooks.Add()

        # Test with empty-tuple argument
        xl.Range["A1", "C1"].Value[()] = (10,"20",31.4)  # XXX: in Python >= 3.8.x, cannot set values to A1:C1
        xl.Range["A2:C2"].Value[()] = ('x','y','z')
        # Test with empty slice argument
        xl.Range["A3:C3"].Value[:] = ('3','2','1')
        # not implemented:
        #     xl.Range["A4:C4"].Value = ("3", "2" ,"1")

        # call property to retrieve value
        expected_values = ((10.0, 20.0, 31.4),
                           ("x", "y", "z"),
                           (3.0, 2.0, 1.0))
        # XXX: in Python >= 3.8.x, fails below
        self.assertEqual(xl.Range["A1:C3"].Value(),
                         expected_values)
        # index with empty tuple
        self.assertEqual(xl.Range["A1:C3"].Value[()],
                         expected_values)
        # index with empty slice
        self.assertEqual(xl.Range["A1:C3"].Value[:],
                         expected_values)
        self.assertEqual(xl.Range["A1:C3"].Value[xlRangeValueDefault],
                         expected_values)
        self.assertEqual(xl.Range["A1", "C3"].Value[()],
                         expected_values)

        # Test for iteration support in "Range" interface
        iter(xl.Range["A1:C3"])
        self.assertEqual([c.Value() for c in xl.Range["A1:C3"]],
                             [10.0, 20.0, 31.4,
                              "x", "y", "z",
                              3.0, 2.0, 1.0])

        # With pywin32, one could write xl.Cells(a, b)
        # With comtypes, one must write xl.Cells.Item(1, b)

        for i in range(20):
            val = "Hi %d" % i
            xl.Cells.Item[i+1,i+1].Value[()] = val
            self.assertEqual(xl.Cells.Item[i+1, i+1].Value[()], val)

        for i in range(20):
            val = "Hi %d" % i
            xl.Cells(i+1,i+1).Value[()] = val
            self.assertEqual(xl.Cells(i+1, i+1).Value[()], val)

        # test dates out with Excel
        xl.Range["A5"].Value[()] = "Excel time"
        xl.Range["B5"].Formula = "=Now()"
        self.assertEqual(xl.Cells.Item[5,2].Formula, "=NOW()")

        xl.Range["A6"].Calculate()
        excel_time = xl.Range["B5"].Value[()]
        self.assertEqual(type(excel_time), datetime.datetime)
        python_time = datetime.datetime.now()

        self.assertTrue(python_time >= excel_time)
        self.assertTrue(python_time - excel_time < datetime.timedelta(seconds=1))

        # some random code, grabbed from c.l.p
        sh = wb.Worksheets[1]

        sh.Cells.Item[1,1].Value[()] = "Hello World!"
        sh.Cells.Item[3,3].Value[()] = "Hello World!"
        sh.Range[sh.Cells.Item[1,1],sh.Cells.Item[3,3]].Copy(sh.Cells.Item[4,1])
        sh.Range[sh.Cells.Item[4,1],sh.Cells.Item[6,3]].Select()


@unittest.skipIf(IMPORT_FAILED, "This depends on Excel.")
@unittest.skip("There is difference of `Range.Value` behavior "
    "between Python >= 3.8.x and Python <= 3.7.x.")
class Test_EarlyBind(BaseBindTest, unittest.TestCase):
    dynamic = False


@unittest.skipIf(IMPORT_FAILED, "This depends on Excel.")
class Test_LateBind(BaseBindTest, unittest.TestCase):
    dynamic = True


if __name__ == "__main__":
    unittest.main()
