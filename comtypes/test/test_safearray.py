import array
import datetime
import unittest
from ctypes import POINTER, c_long
from decimal import Decimal

from comtypes import BSTR, IUnknown
from comtypes._safearray import (
    SAFEARRAY,
    SafeArrayCopy,
    SafeArrayDestroy,
    SafeArrayGetVartype,
)
from comtypes.automation import (
    VARIANT,
    VARIANT_BOOL,
    VT_ARRAY,
    VT_BSTR,
    VT_I4,
    VT_R4,
    VT_R8,
    VT_VARIANT,
    _midlSAFEARRAY,
)
from comtypes.safearray import safearray_as_ndarray
from comtypes.test.find_memleak import find_memleak


def get_array(sa):
    """Get an array from a safe array type"""
    with safearray_as_ndarray:
        return sa[0]


def com_refcnt(o):
    """Return the COM refcount of an interface pointer"""
    import gc

    gc.collect()
    gc.collect()
    o.AddRef()
    return o.Release()


class VariantTestCase(unittest.TestCase):
    def test_VARIANT_array(self):
        v = VARIANT()
        v.value = ((1, 2, 3), ("foo", "bar", None))
        self.assertEqual(v.vt, VT_ARRAY | VT_VARIANT)
        self.assertEqual(v.value, ((1, 2, 3), ("foo", "bar", None)))

        def func():
            VARIANT((1, 2, 3), ("foo", "bar", None))

        # The VARIANT owns a SAFEARRAY of VARIANTs (with BSTR elements).
        # `mallocspy` counts these COM task allocations exactly, so a per-call
        # leak is detected deterministically. `working_set` was unreliable here
        # because it measures resident pages, and the first call in a process
        # faults in pages that look like a leak
        # (see https://github.com/enthought/comtypes/issues/971).
        bytes = find_memleak(func, target="mallocspy")
        self.assertFalse(bytes, "Leaks %d bytes" % bytes)

    def test_double_array(self):
        a = array.array("d", (3.14, 2.78))
        v = VARIANT(a)
        self.assertEqual(v.vt, VT_ARRAY | VT_R8)
        self.assertEqual(tuple(a.tolist()), v.value)

        def func():
            VARIANT(array.array("d", (3.14, 2.78)))

        # The VARIANT owns a SAFEARRAY of doubles. `mallocspy` counts this COM
        # task allocation exactly, unlike `working_set`, which measures
        # resident pages and reports first-call page faults as leaks
        # (see https://github.com/enthought/comtypes/issues/971).
        bytes = find_memleak(func, target="mallocspy")
        self.assertFalse(bytes, "Leaks %d bytes" % bytes)

    def test_float_array(self):
        a = array.array("f", (3.14, 2.78))
        v = VARIANT(a)
        self.assertEqual(v.vt, VT_ARRAY | VT_R4)
        self.assertEqual(tuple(a.tolist()), v.value)

    def test_2dim_array(self):
        data = ((1, 2, 3, 4), (5, 6, 7, 8), (9, 10, 11, 12))
        v = VARIANT(data)
        self.assertEqual(v.value, data)


class SafeArrayTestCase(unittest.TestCase):
    def test_equality(self):
        a = _midlSAFEARRAY(c_long)
        b = _midlSAFEARRAY(c_long)
        self.assertTrue(a is b)

        c = _midlSAFEARRAY(BSTR)
        d = _midlSAFEARRAY(BSTR)
        self.assertTrue(c is d)

        self.assertNotEqual(a, c)

        # XXX remove:
        self.assertEqual((a._itemtype_, a._vartype_), (c_long, VT_I4))
        self.assertEqual((c._itemtype_, c._vartype_), (BSTR, VT_BSTR))

    def test_VT_BSTR(self):
        t = _midlSAFEARRAY(BSTR)

        sa = t.from_param(["a", "b", "c"])
        self.assertEqual(sa[0], ("a", "b", "c"))
        self.assertEqual(SafeArrayGetVartype(sa), VT_BSTR)

    def test_VT_BSTR_leaks(self):
        sb = _midlSAFEARRAY(BSTR)

        def doit():
            sb.from_param(["foo", "bar"])

        # `from_param` creates a SAFEARRAY of BSTRs. `mallocspy` counts the
        # SAFEARRAY descriptor/data and the BSTR elements exactly, so a
        # per-call leak is detected deterministically. `working_set` was
        # unreliable here because it measures resident pages, and the first
        # call in a process faults in pages that look like a leak
        # (see https://github.com/enthought/comtypes/issues/971).
        bytes = find_memleak(doit, target="mallocspy")
        self.assertFalse(bytes, "Leaks %d bytes" % bytes)

    def test_VT_I4_leaks(self):
        sa = _midlSAFEARRAY(c_long)

        def doit():
            sa.from_param([1, 2, 3, 4, 5, 6])

        # `from_param` creates a SAFEARRAY of longs. `mallocspy` counts this
        # COM task allocation exactly, unlike `working_set`, which measures
        # resident pages and reports first-call page faults as leaks
        # (see https://github.com/enthought/comtypes/issues/971).
        bytes = find_memleak(doit, target="mallocspy")
        self.assertFalse(bytes, "Leaks %d bytes" % bytes)

    def test_VT_I4(self):
        t = _midlSAFEARRAY(c_long)

        sa = t.from_param([11, 22, 33])

        self.assertEqual(sa[0], (11, 22, 33))

        self.assertEqual(SafeArrayGetVartype(sa), VT_I4)

        # TypeError: len() of unsized object
        with self.assertRaises(TypeError):
            t.from_param(object())

    def test_VT_VARIANT(self):
        t = _midlSAFEARRAY(VARIANT)

        now = datetime.datetime.now()
        sa = t.from_param([11, "22", None, True, now, Decimal("3.14")])
        self.assertEqual(sa[0], (11, "22", None, True, now, Decimal("3.14")))

        self.assertEqual(SafeArrayGetVartype(sa), VT_VARIANT)

    def test_VT_BOOL(self):
        t = _midlSAFEARRAY(VARIANT_BOOL)

        sa = t.from_param([True, False, True, False])
        self.assertEqual(sa[0], (True, False, True, False))

    def test_VT_UNKNOWN_1(self):
        a = _midlSAFEARRAY(POINTER(IUnknown))
        t = _midlSAFEARRAY(POINTER(IUnknown))
        self.assertTrue(a is t)

        from comtypes.typeinfo import CreateTypeLib

        # will never be saved to disk
        punk = CreateTypeLib("spam").QueryInterface(IUnknown)

        # initial refcount
        initial = com_refcnt(punk)

        # This should increase the refcount by 1
        sa = t.from_param([punk])
        self.assertEqual(initial + 1, com_refcnt(punk))

        # Unpacking the array must not change the refcount, and must
        # return an equal object.
        self.assertEqual((punk,), sa[0])
        self.assertEqual(initial + 1, com_refcnt(punk))

        del sa
        self.assertEqual(initial, com_refcnt(punk))

        sa = t.from_param([None])
        self.assertEqual((POINTER(IUnknown)(),), sa[0])

    def test_VT_UNKNOWN_multi(self):
        a = _midlSAFEARRAY(POINTER(IUnknown))
        t = _midlSAFEARRAY(POINTER(IUnknown))
        self.assertTrue(a is t)

        from comtypes.typeinfo import CreateTypeLib

        # will never be saved to disk
        punk = CreateTypeLib("spam").QueryInterface(IUnknown)

        # initial refcount
        initial = com_refcnt(punk)

        # This should increase the refcount by 4
        sa = t.from_param((punk,) * 4)
        self.assertEqual(initial + 4, com_refcnt(punk))

        # Unpacking the array must not change the refcount, and must
        # return an equal object.
        self.assertEqual((punk,) * 4, sa[0])
        self.assertEqual(initial + 4, com_refcnt(punk))

        del sa
        self.assertEqual(initial, com_refcnt(punk))

        # This should increase the refcount by 2
        sa = t.from_param((punk, None, punk, None))
        self.assertEqual(initial + 2, com_refcnt(punk))

        null = POINTER(IUnknown)()
        self.assertEqual((punk, null, punk, null), sa[0])

        del sa
        self.assertEqual(initial, com_refcnt(punk))

        # repeat same test, with 2 different com pointers

        plib = CreateTypeLib("foo")
        a, b = com_refcnt(plib), com_refcnt(punk)
        sa = t.from_param([plib, punk, plib])

        ####        self.failUnlessEqual((plib, punk, plib), sa[0])
        self.assertEqual((a + 2, b + 1), (com_refcnt(plib), com_refcnt(punk)))

        del sa
        self.assertEqual((a, b), (com_refcnt(plib), com_refcnt(punk)))


class SafeArrayCopyTestCase(unittest.TestCase):
    def test(self):
        sa_type = _midlSAFEARRAY(c_long)
        src = sa_type.from_param([10, 20, 30])
        copied = SafeArrayCopy(src)
        try:
            # `SafeArrayCopy` returns a plain `POINTER(SAFEARRAY)`, not the
            # type created by `_midlSAFEARRAY`.
            self.assertIsInstance(copied, POINTER(SAFEARRAY))
            self.assertNotIsInstance(copied, sa_type)
        finally:
            # The type from `_midlSAFEARRAY` has a `__del__` method that calls
            # `SafeArrayDestroy` automatically.
            # But the plain pointer does not have this, so we must call
            # `SafeArrayDestroy` ourselves to free the memory.
            SafeArrayDestroy(copied)


if __name__ == "__main__":
    unittest.main()
