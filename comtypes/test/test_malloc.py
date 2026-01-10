import unittest as ut
from ctypes import POINTER, byref

from comtypes.malloc import IMalloc, _CoGetMalloc


def _get_malloc() -> IMalloc:
    malloc = POINTER(IMalloc)()
    _CoGetMalloc(1, byref(malloc))
    assert bool(malloc)
    return malloc  # type: ignore


class Test(ut.TestCase):
    def test_Realloc(self):
        malloc = _get_malloc()
        size1 = 4
        ptr1 = malloc.Alloc(size1)
        self.assertEqual(malloc.DidAlloc(ptr1), 1)
        self.assertEqual(malloc.GetSize(ptr1), size1)
        size2 = size1 - 1
        ptr2 = malloc.Realloc(ptr1, size2)
        self.assertEqual(malloc.DidAlloc(ptr1), 0)
        self.assertEqual(malloc.DidAlloc(ptr2), 1)
        self.assertEqual(malloc.GetSize(ptr2), size2)
        size3 = size1 + 1
        ptr3 = malloc.Realloc(ptr2, size3)
        self.assertEqual(malloc.DidAlloc(ptr2), 0)
        self.assertEqual(malloc.DidAlloc(ptr3), 1)
        self.assertEqual(malloc.GetSize(ptr3), size3)
        malloc.Free(ptr3)
        self.assertEqual(malloc.DidAlloc(ptr3), 0)
        malloc.HeapMinimize()
        del ptr3
