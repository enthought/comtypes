import gc
import unittest

from comtypes.malloc import CoGetMalloc
from comtypes.test.find_memleak import find_memleak, wss


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

        # `find_memleak()` takes its baseline with one `wss()` call and its
        # final reading with another. The first `wss()` call in a process
        # incurs one-time costs (e.g. first-time code paths faulting in
        # pages) that land between those two readings, so a cold probe can
        # make an otherwise leak-free function look like it grew by a page.
        # Calling `wss()` once here pays that cost beforehand.
        #
        # Note that this only warms up the probe itself. The warmup of
        # `no_leak()` is already done by `find_memleak()`.
        #
        # This is a workaround that makes the measurement stable in practice.
        # It does not address the fundamental limitation of Working Set
        # based measurement (see https://github.com/enthought/comtypes/issues/971):
        # the Working Set is the amount of resident pages, not the amount of
        # unreleased memory, so a stable result does not prove the absence of
        # a leak.
        _ = wss()
        # Actual measurement. `find_memleak` already warms up `no_leak`.
        leak_amount = find_memleak(no_leak)
        self.assertEqual(leak_amount, 0)


if __name__ == "__main__":
    unittest.main()
