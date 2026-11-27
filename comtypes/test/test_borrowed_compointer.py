import ctypes
import gc
import unittest

from comtypes import COMObject, IUnknown
from comtypes.automation import VARIANT, _midlSAFEARRAY


class Owner(COMObject):
    _com_interfaces_ = [IUnknown]


class PointerField(ctypes.Structure):
    _fields_ = [("pointer", ctypes.POINTER(IUnknown))]


class BorrowedCOMPointerTest(unittest.TestCase):
    def test_destroying_view_does_not_release_the_com_reference(self):
        for make_holder in (ctypes.pointer, PointerField):
            with self.subTest(make_holder=make_holder):
                owner = Owner()
                pointer = owner.QueryInterface(IUnknown)
                # Keep the COM object alive on the failing path so cleanup is safe.
                pointer.AddRef()
                holder = make_holder(pointer)
                if isinstance(holder, PointerField):
                    view = holder.pointer
                else:
                    view = holder.contents
                self.assertIs(view._b_base_, holder)
                self.assertEqual(owner._refcnt.value, 2)

                del view
                gc.collect()
                remaining = owner._refcnt.value
                if remaining == 2:
                    pointer.Release()
                del holder, pointer
                gc.collect()

                self.assertEqual(remaining, 2)
                self.assertEqual(owner._refcnt.value, 0)

    def test_variant_and_returned_pointer_own_separate_references(self):
        owner = Owner()
        pointer = owner.QueryInterface(IUnknown)
        variant = VARIANT(pointer)
        self.assertEqual(owner._refcnt.value, 2)
        returned = variant.value
        self.assertEqual(owner._refcnt.value, 3)
        del returned
        gc.collect()
        self.assertEqual(owner._refcnt.value, 2)
        variant.value = None
        self.assertEqual(owner._refcnt.value, 1)

    def test_safearray_and_unpacked_pointers_own_separate_references(self):
        owner = Owner()
        pointer = owner.QueryInterface(IUnknown)
        array = _midlSAFEARRAY(type(pointer)).from_param([pointer, pointer])
        gc.collect()
        self.assertEqual(owner._refcnt.value, 3)
        returned = array[0]
        gc.collect()
        self.assertEqual(owner._refcnt.value, 5)
        del returned
        gc.collect()
        self.assertEqual(owner._refcnt.value, 3)
        del array
        gc.collect()
        self.assertEqual(owner._refcnt.value, 1)
