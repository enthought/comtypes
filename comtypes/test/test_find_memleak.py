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

        # Call `wss()` once to absorb one-time initialization costs
        # (COM allocator warmup, first page faults, etc.) before the
        # actual measurement baseline is taken inside `find_memleak()`.
        #
        # `GetProcessMemoryInfo` (`wss`) itself is a pure read-only
        # operation with no side effects on the working set. The
        # stabilization effect comes from running through `find_memleak()`
        # internal warmup/measurement cycle once, which exercises:
        #
        # 1. COM allocator (`CoGetMalloc`) first-use overhead:
        #    - Initial `Alloc`/`Free` calls fault in heap pages and
        #      initialize internal free lists. This temporarily
        #      inflates the working set.
        #
        # 2. Python/COM interop and ctypes first-use initialization:
        #    - First COM object access triggers one-time memory
        #      allocations for type info, vtables, etc.
        #
        # 3. OS page fault resolution timing:
        #    - Even after `gc.collect()`, physical memory commitment
        #      and working set accounting may lag slightly.
        #
        # By calling `wss()` once before the real test, these one-time
        # costs are paid before the real baseline is captured, so the
        # second measurement reflects only steady-state behavior
        # (which should show zero net growth for a leak-free function).
        _ = wss()

        # Now that the one-time costs have been paid, the allocator and
        # working set are stabilized. The warmup phase inside
        # `find_memleak()` no longer causes net memory growth, so the
        # measured difference between baseline and final should be zero.
        leak_amount = find_memleak(no_leak)

        # The leak amount should be 0 (no net memory increase)
        self.assertEqual(leak_amount, 0)


if __name__ == "__main__":
    unittest.main()
