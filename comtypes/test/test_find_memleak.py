import gc
import unittest

from comtypes.malloc import CoGetMalloc
from comtypes.test.find_memleak import find_memleak


class Test_find_memleak(unittest.TestCase):
    def setUp(self):
        gc.collect()
        gc.collect()
        gc.collect()

    def tearDown(self):
        gc.collect()
        gc.collect()
        gc.collect()

    def test_find_memleak_working_set_no_leak(self):
        def no_leak():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            malloc.Free(ptr)

        leak_amount = find_memleak(no_leak)
        self.assertEqual(leak_amount, 0)


if __name__ == "__main__":
    unittest.main()
