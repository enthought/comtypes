import sys
import unittest


class NumpySupportTestCase(unittest.TestCase):
    def test_not_imported_imported(self):
        self.assertFalse("numpy" in sys.modules)

        import comtypes.npsupport
        self.assertFalse("numpy" in sys.modules)
        self.assertRaises(ImportError, comtypes.npsupport.get_numpy)
        comtypes.npsupport.enable_numpy_interop()
        self.assertTrue("numpy" in sys.modules)
        import numpy
        self.assertEqual(numpy, comtypes.npsupport.get_numpy())


if __name__ == '__main__':
    unittest.main()
