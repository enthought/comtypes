import gc
import unittest
from collections.abc import Callable
from typing import Any

from comtypes.malloc import CoGetMalloc, _CoRegisterMallocSpy, _CoRevokeMallocSpy
from comtypes.test.find_memleak import (
    BytesTrackingMallocSpy,
    create_memdiff_detector,
    find_memleak,
    get_working_set_size,
)


class Test_BytesTrackingMallocSpy(unittest.TestCase):
    """Test the BytesTrackingMallocSpy COMObject implementation."""

    def _run_with_spy(self, test_func: Callable[[BytesTrackingMallocSpy], Any]) -> None:
        """Helper to run a test with a registered spy, ensuring cleanup."""
        self.addCleanup(_CoRevokeMallocSpy)
        spy = BytesTrackingMallocSpy()
        _CoRegisterMallocSpy(spy)
        test_func(spy)

    def test_tracks_allocations(self):
        def _test(spy: BytesTrackingMallocSpy) -> None:
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            ptr2 = malloc.Alloc(200)
            malloc.Free(ptr1)
            malloc.Free(ptr2)
            # Check that allocations were tracked in bytes
            allocated = spy.get_allocated_bytes()
            freed = spy.get_freed_bytes()
            # We should have at least 300 bytes allocated and 300 bytes freed
            self.assertEqual(allocated, 300)
            self.assertEqual(freed, 300)
            # Net should be 0 since we freed everything
            net = spy.get_net_bytes()
            self.assertEqual(net, 0)

        self._run_with_spy(_test)

    def test_detects_leak(self):
        def _test(spy: BytesTrackingMallocSpy) -> None:
            # Do some COM allocations but don't free them (leak)
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            ptr2 = malloc.Alloc(200)
            # Intentionally not freeing - this is a leak
            # Check that allocations were tracked
            allocated = spy.get_allocated_bytes()
            freed = spy.get_freed_bytes()
            # We should have 300 bytes allocated and 0 bytes freed
            self.assertEqual(allocated, 300)
            self.assertEqual(freed, 0)
            # Net should be positive (leak detected)
            net = spy.get_net_bytes()
            self.assertEqual(net, 300)
            # Free the memory so we can revoke the spy
            malloc.Free(ptr1)
            malloc.Free(ptr2)

        self._run_with_spy(_test)

    def test_realloc(self):
        def _test(spy: BytesTrackingMallocSpy) -> None:
            malloc = CoGetMalloc()
            # Allocate initial memory (100 bytes)
            ptr1 = malloc.Alloc(100)
            # Realloc to larger size (200 bytes) - frees 100, allocates 200
            ptr2 = malloc.Realloc(ptr1, 200)
            # Realloc to smaller size (50 bytes) - frees 200, allocates 50
            ptr3 = malloc.Realloc(ptr2, 50)
            malloc.Free(ptr3)
            # Allocations: 100 + 200 + 50 = 350
            # Frees: 100 + 200 + 50 = 350
            # Net: 0
            self.assertEqual(spy.get_allocated_bytes(), 350)
            self.assertEqual(spy.get_freed_bytes(), 350)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)

    def test_no_changing_byte_counts(self):
        def _test(spy: BytesTrackingMallocSpy) -> None:
            malloc = CoGetMalloc()
            self.assertEqual(spy.get_net_bytes(), 0)
            # Allocate initial memory
            ptr1 = malloc.Alloc(100)
            self.assertEqual(spy.get_net_bytes(), 100)
            # This should trigger `PreGetSize`/`PostGetSize`
            self.assertEqual(malloc.GetSize(ptr1), 100)
            self.assertEqual(spy.get_net_bytes(), 100)
            # This should trigger `PreDidAlloc`/`PostDidAlloc`
            self.assertEqual(malloc.DidAlloc(ptr1), 1)
            self.assertEqual(spy.get_net_bytes(), 100)
            # This should trigger `PreHeapMinimize`/`PostHeapMinimize`
            self.assertEqual(malloc.HeapMinimize(), None)
            self.assertEqual(spy.get_net_bytes(), 100)
            # Free the memory
            malloc.Free(ptr1)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)


class Test_create_memdiff_detector(unittest.TestCase):
    def test_working_set_detects_leak_and_cleanup(self):
        with create_memdiff_detector("working_set") as detector:
            malloc = CoGetMalloc()
            # Allocate 2MB and intentionally leak it
            ptr1 = malloc.Alloc(1024 * 1024)
            ptr2 = malloc.Alloc(1024 * 1024)
            # Leak detected: WorkingSetSize measures resident pages.
            # The increase may be less than 2MB if pages haven't been faulted
            # in yet, but must be non-negative. We verify it's at least 0
            # (no false negative).
            self.assertGreaterEqual(detector.get_diff(), 0)
            # Cleanup: free the leaked memory
            malloc.Free(ptr1)
            malloc.Free(ptr2)
            # After cleanup, working set may not immediately shrink due to OS
            # keeping pages for reuse. Allow up to 1MB residual.
            self.assertLessEqual(detector.get_diff(), 1024 * 1024)

    def test_private_usage_detects_leak_and_cleanup(self):
        with create_memdiff_detector("private_usage") as detector:
            malloc = CoGetMalloc()
            # Allocate 2MB and intentionally leak it
            ptr1 = malloc.Alloc(1024 * 1024)
            ptr2 = malloc.Alloc(1024 * 1024)
            # Leak detected: PrivateUsage (commit charge) must increase by at least
            # the allocation size because committed pages are tracked regardless of
            # whether they are resident. Allocator overhead may add some bytes.
            self.assertGreaterEqual(detector.get_diff(), 2 * 1024 * 1024)
            # Cleanup: free the leaked memory
            malloc.Free(ptr1)
            malloc.Free(ptr2)
            # After cleanup, commit charge should return close to baseline.
            # Allow up to 1MB residual due to allocator caching.
            self.assertLessEqual(detector.get_diff(), 1024 * 1024)

    def test_mallocspy_detects_leak_and_cleanup(self):
        with create_memdiff_detector("mallocspy") as detector:
            malloc = CoGetMalloc()
            # Allocate 2MB and intentionally leak it
            ptr1 = malloc.Alloc(1024 * 1024)
            ptr2 = malloc.Alloc(1024 * 1024)
            # Leak detected: exact 2MB
            self.assertEqual(detector.get_diff(), 2 * 1024 * 1024)
            # Cleanup: free the leaked memory
            malloc.Free(ptr1)
            malloc.Free(ptr2)
            # After cleanup, leak should be gone (net 0)
            self.assertEqual(detector.get_diff(), 0)


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

        # `find_memleak()` takes its baseline with one `get_working_set_size()`
        # call and its final reading with another.
        # The first `get_working_set_size()` call in a process incurs one-time
        # costs (e.g. first-time code paths faulting in pages) that land
        # between those two readings, so a cold probe can make an otherwise
        # leak-free function look like it grew by a page.
        # Calling `get_working_set_size()` once here pays that cost beforehand.
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
        _ = get_working_set_size()
        # Actual measurement. `find_memleak` already warms up `no_leak`.
        leak_amount = find_memleak(no_leak)
        self.assertEqual(leak_amount, 0)

    def test_find_memleak_private_usage_no_leak(self):
        def no_leak():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            malloc.Free(ptr)

        # Call find_memleak - it should not detect a leak
        leak_amount = find_memleak(no_leak, target="private_usage")
        # The leak amount should be 0 (no net memory increase)
        self.assertEqual(leak_amount, 0)

    def test_find_memleak_mallocspy_no_leak(self):
        def no_leak():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            malloc.Free(ptr)

        # Call find_memleak - it should not detect a leak
        leak_amount = find_memleak(no_leak, target="mallocspy")
        # The leak amount should be 0 (no net allocations)
        self.assertEqual(leak_amount, 0)

    def test_find_memleak_invalid_taget(self):
        with self.assertRaises(TypeError):
            find_memleak(lambda: None, target="invalid target")


if __name__ == "__main__":
    unittest.main()
