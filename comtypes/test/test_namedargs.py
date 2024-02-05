from _ctypes import COMError
from pathlib import Path
import tempfile
import unittest as ut

from comtypes import client
from comtypes.automation import VARIANT

# TODO: Add TestCase using non-env-specific typelib.


class Test_Excel(ut.TestCase):
    """for DispMethods"""

    def setUp(self):
        try:
            client.GetModule(("{00020813-0000-0000-C000-000000000046}",))
            from comtypes.gen import Excel

            self.Excel = Excel
            self.xl = client.CreateObject(
                Excel.Application, interface=Excel._Application
            )
        except (ImportError, OSError):
            self.skipTest("This depends on Excel.")

    def tearDown(self):
        # Close all open workbooks without saving, then quit excel.
        for wb in self.xl.Workbooks:
            wb.Close(0)
        self.xl.Quit()
        del self.xl

    def test_range_value(self):
        xl = self.xl
        xl.Workbooks.Add()
        xl.Range["A1:C1"].Value[()] = (10, "20", 31.4)
        xl.Range["A2:C2"].Value[()] = ("x", "y", "z")
        xl.Range["A3:C3"].Value[:] = ("3", "2", "1")
        rng = xl.Range("A1:C3")
        expected_values = ((10.0, 20.0, 31.4), ("x", "y", "z"), (3.0, 2.0, 1.0))
        self.assertTrue(
            rng.Value[self.Excel.xlRangeValueDefault]
            == rng.Value(RangeValueDataType=self.Excel.xlRangeValueDefault)
            == rng.Value(VARIANT.missing)
            == rng.Value()
            == expected_values
        )
        with self.assertRaises(TypeError):
            rng.Value(
                self.Excel.xlRangeValueDefault,
                RangeValueDataType=self.Excel.xlRangeValueDefault,
            )
        with self.assertRaises(TypeError):
            rng.Value(self.Excel.xlRangeValueDefault, Foo="Ham")
        with self.assertRaises(TypeError):
            rng.Value(Foo="Ham")

    def test_range_address(self):
        Excel = self.Excel
        self.xl.Workbooks.Add()
        rng = self.xl.Range("A1:C3")
        self.assertTrue(
            rng.Address()
            == rng.Address(None)
            == rng.Address[None, None]
            == rng.Address(RowAbsolute=True)
            == rng.Address(ColumnAbsolute=True)
            == rng.Address(RowAbsolute=True, ColumnAbsolute=True)
            == rng.Address(ColumnAbsolute=True, RowAbsolute=True)
            == "$A$1:$C$3"
        )
        self.assertTrue(
            rng.Address(RowAbsolute=False)
            == rng.Address[False]
            == rng.Address(ColumnAbsolute=True, RowAbsolute=False)
            == "$A1:$C3"
        )
        self.assertTrue(
            rng.Address[None, False]
            == rng.Address(ColumnAbsolute=False)
            == rng.Address(ColumnAbsolute=False, RowAbsolute=True)
            == "A$1:C$3"
        )
        self.assertTrue(
            rng.Address[None, None, self.Excel.xlR1C1]
            == rng.Address(None, ReferenceStyle=self.Excel.xlR1C1)
            == rng.Address(ReferenceStyle=self.Excel.xlR1C1)
            == "R1C1:R3C3"
        )
        with self.assertRaises(TypeError):
            rng.Address(Foo="Ham")
        with self.assertRaises(COMError):
            rng.Address(False, False, Excel.xlR1C1, False, self.xl.Range("D4"), False)
        with self.assertRaises(TypeError):
            rng.Address(
                False, False, Excel.xlR1C1, False, self.xl.Range("D4"), Foo="Ham"
            )
        with self.assertRaises(TypeError):
            rng.Address(
                RowAbsolute=False,
                ColumnAbsolute=False,
                ReferenceStyle=Excel.xlR1C1,
                External=False,
                RelativeTo=self.xl.Range("D4"),
                Foo="Ham",
            )
        with self.assertRaises(TypeError):
            rng.Address(
                RowAbsolute=False,
                ColumnAbsolute=False,
                ReferenceStyle=Excel.xlR1C1,
                External=False,
                Foo="Ham",
            )

    def test_range_autofill(self):
        Excel = self.Excel
        xl = self.xl
        xl.Workbooks.Add()
        xl.Range["A1:E1"].Value[()] = (1, 1, 1, 1, 1)
        xl.Range["A2:E2"].Value[()] = (2, 2, 2, 2, 2)
        xl.Range("A1").AutoFill(xl.Range("A1:A4"))
        xl.Range("B1").AutoFill(xl.Range("B1:B4"), Excel.xlFillSeries)
        xl.Range("C1:C2").AutoFill(xl.Range("C1:C4"), Excel.xlFillCopy)
        xl.Range("C1:C2").AutoFill(xl.Range("C1:C4"), Type=Excel.xlFillCopy)
        xl.Range("D1:D2").AutoFill(Destination=xl.Range("D1:D4"), Type=Excel.xlFillCopy)
        xl.Range("E1:E3").AutoFill(Type=Excel.xlFillCopy, Destination=xl.Range("E1:E4"))
        self.assertEqual(
            xl.Range("A1:E4").Value(),
            (
                (1.0, 1.0, 1.0, 1.0, 1.0),
                (1.0, 2.0, 2.0, 2.0, 2.0),
                (1.0, 3.0, 1.0, 1.0, None),
                (1.0, 4.0, 2.0, 2.0, 1.0),
            ),
        )
        with self.assertRaises(COMError):
            xl.Range("A1").AutoFill()
        with self.assertRaises(TypeError):
            xl.Range("B1").AutoFill(xl.Range("B1:B4"), Destination=xl.Range("B1:B4"))
        with self.assertRaises(TypeError):
            xl.Range("B1").AutoFill(Type=Excel.xlFillCopy)
        with self.assertRaises(TypeError):
            xl.Range("B1").AutoFill(
                xl.Range("B1:B4"), Type=Excel.xlFillCopy, Foo="spam"
            )


class Test_IDictionary(ut.TestCase):
    """for ComMethods"""

    def setUp(self):
        client.GetModule("scrrun.dll")
        from comtypes.gen import Scripting

        self.dic = client.CreateObject(
            Scripting.Dictionary, interface=Scripting.IDictionary
        )

    def tearDown(self):
        del self.dic

    def test_takes_valid_args(self):
        self.dic.Add(Key="foo", Item="spam")
        self.dic.Add("bar", Item="ham")
        self.dic.Add(Item="bacon", Key="baz")
        self.dic.Add("qux", "egg")
        self.assertEqual(set(self.dic.Keys()), {"foo", "bar", "baz", "qux"})
        self.assertEqual(set(self.dic.Items()), {"spam", "ham", "bacon", "egg"})

    def test_takes_invalid_args(self):
        with self.assertRaises(TypeError):
            self.dic.Add(Key="foo", Item="spam", Eric="Idle")
        with self.assertRaises(TypeError):
            self.dic.Add("foo", "spam", Eric="Idle")
        with self.assertRaises(TypeError):
            self.dic.Add("foo")


class Test_FSO(ut.TestCase):
    """for ComMethods"""

    def setUp(self):
        client.GetModule("scrrun.dll")
        from comtypes.gen import Scripting

        self.fso = client.CreateObject(
            Scripting.FileSystemObject, interface=Scripting.IFileSystem
        )

    def tearDown(self):
        del self.fso

    def test_takes_valid_args(self):
        with tempfile.TemporaryDirectory() as t:
            tmp_dir = Path(t)
            tmp_file = tmp_dir / "tmp.txt"
            for args, kwargs in [
                ((tmp_file.__fspath__(),), {}),
                ((tmp_file.__fspath__(), True), {}),
                ((tmp_file.__fspath__(),), {"Force": True}),
                ((), {"FileSpec": tmp_file.__fspath__(), "Force": True}),
                ((), {"Force": True, "FileSpec": tmp_file.__fspath__()}),
            ]:
                tmp_file.touch()
                with self.subTest(args=args, kwargs=kwargs):
                    self.fso.DeleteFile(*args, **kwargs)
                    self.assertFalse(tmp_file.exists())

    def test_takes_invalid_args(self):
        with tempfile.TemporaryDirectory() as t:
            tmp_dir = Path(t)
            tmp_file = tmp_dir / "tmp.txt"
            tmp_file.touch()
            with self.assertRaises(TypeError):
                self.fso.DeleteFile()
            with self.assertRaises(TypeError):
                self.fso.DeleteFile(Force=True)
            with self.assertRaises(TypeError):
                self.fso.DeleteFile(
                    tmp_file.__fspath__(), FileSpec=tmp_file.__fspath__()
                )


if __name__ == "__main__":
    ut.main()
