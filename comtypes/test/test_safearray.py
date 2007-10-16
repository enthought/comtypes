import unittest
from decimal import Decimal
import datetime
from ctypes import *
from comtypes.test.find_memleak import find_memleak
from comtypes import BSTR, IUnknown

from comtypes.automation import VARIANT, IDispatch, VT_ARRAY, VT_VARIANT, \
     VT_I4, VT_R4, VT_R8, VT_BSTR, VARIANT_BOOL, VT_DATE, VT_CY
from comtypes.automation import _midlSAFEARRAY

from comtypes._safearray import SafeArrayGetVartype

class VariantTestCase(unittest.TestCase):
    def test_VARIANT_array(self):
        v = VARIANT()
        v.value = ((1, 2, 3), ("foo", "bar", None))
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_VARIANT)
        self.failUnlessEqual(v.value, ((1, 2, 3), ("foo", "bar", None)))

        def func():
            v = VARIANT((1, 2, 3), ("foo", "bar", None))

        bytes = find_memleak(func)
        self.failIf(bytes, "Leaks %d bytes" % bytes)


    def test_object(self):
        self.assertRaises(TypeError, lambda: VARIANT(object()))

    def test_double_array(self):
        import array
        a = array.array("d", (3.14, 2.78))
        v = VARIANT(a)
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_R8)
        self.failUnlessEqual(tuple(a.tolist()), v.value)

        def func():
            v = VARIANT(array.array("d", (3.14, 2.78)))

        bytes = find_memleak(func)
        self.failIf(bytes, "Leaks %d bytes" % bytes)


    def test_float_array(self):
        import array
        a = array.array("f", (3.14, 2.78))
        v = VARIANT(a)
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_R4)
        self.failUnlessEqual(tuple(a.tolist()), v.value)

##    def test_2dim_array(self):
##        data = ((1, 2, 3, 4),
##                (5, 6, 7, 8),
##                (9, 10, 11, 12))
##        from comtypes.safearray import SafeArray_FromSequence, UnpackSafeArray
##        a = SafeArray_FromSequence(data)
##        self.failUnlessEqual(UnpackSafeArray(a), data)

    def test_datetime(self):
        now = datetime.datetime.now()

        v = VARIANT()
        v.value = now
        self.failUnlessEqual(v.value, now)
        self.failUnlessEqual(v.vt, VT_DATE)

    def test_decimal(self):
        pi = Decimal("3.13")

        v = VARIANT()
        v.value = pi
        self.failUnlessEqual(v.vt, VT_CY)
        self.failUnlessEqual(v.value, pi)

    def test_UDT(self):
        from comtypes.gen.TestComServerLib import MYCOLOR
        v = VARIANT(MYCOLOR(red=1.0, green=2.0, blue=3.0))
        value = v.value
        self.failUnlessEqual((1.0, 2.0, 3.0),
                             (value.red, value.green, value.blue))

        def func():
            v = VARIANT(MYCOLOR(red=1.0, green=2.0, blue=3.0))
            return v.value

        bytes = find_memleak(func)
        self.failIf(bytes, "Leaks %d bytes" % bytes)


class SafeArrayTestCase(unittest.TestCase):

    def test_equality(self):
        a = _midlSAFEARRAY(c_long)
        b = _midlSAFEARRAY(c_long)
        self.failUnless(a is b)

        c = _midlSAFEARRAY(BSTR)
        d = _midlSAFEARRAY(BSTR)
        self.failUnless(c is d)

        self.failIfEqual(a, c)

        # XXX remove:
        self.failUnlessEqual((a._itemtype_, a._vartype_),
                             (c_long, VT_I4))
        self.failUnlessEqual((c._itemtype_, c._vartype_),
                             (BSTR, VT_BSTR))

    def test_VT_BSTR(self):
        t = _midlSAFEARRAY(BSTR)

        sa = t.from_param(["a" ,"b", "c"])
        self.failUnlessEqual(sa[0], ("a", "b", "c"))
        self.failUnlessEqual(SafeArrayGetVartype(sa), VT_BSTR)

    def test_VT_BSTR_leaks(self):
        sb = _midlSAFEARRAY(BSTR)
        def doit():
            sb.from_param(["foo", "bar"])

        bytes = find_memleak(doit)
        self.failIf(bytes, "Leaks %d bytes" % bytes)

    def test_VT_I4_leaks(self):
        sa = _midlSAFEARRAY(c_long)
        def doit():
            sa.from_param([1, 2, 3, 4, 5, 6])

        bytes = find_memleak(doit)
        self.failIf(bytes, "Leaks %d bytes" % bytes)

    def test_VT_I4(self):
        t = _midlSAFEARRAY(c_long)

        sa = t.from_param([11, 22, 33])

        self.failUnlessEqual(sa[0], (11, 22, 33))

        self.failUnlessEqual(SafeArrayGetVartype(sa), VT_I4)

        # TypeError: len() of unsized object
        self.assertRaises(TypeError, lambda: t.from_param(object()))

    def test_VT_VARIANT(self):
        t = _midlSAFEARRAY(VARIANT)

        now = datetime.datetime.now()
        sa = t.from_param([11, "22", None, True, now, Decimal("3.14")])
        self.failUnlessEqual(sa[0], (11, "22", None, True, now, Decimal("3.14")))

        self.failUnlessEqual(SafeArrayGetVartype(sa), VT_VARIANT)

    def test_VT_BOOL(self):
        t = _midlSAFEARRAY(VARIANT_BOOL)

        sa = t.from_param([True, False, True, False])
        self.failUnlessEqual(sa[0], (True, False, True, False))

    def test_UDT(self):
        from comtypes.gen.TestComServerLib import MYCOLOR

        t = _midlSAFEARRAY(MYCOLOR)
        self.failUnless(t is _midlSAFEARRAY(MYCOLOR))

        sa = t.from_param([MYCOLOR(0, 0, 0), MYCOLOR(1, 2, 3)])

        self.failUnlessEqual([(x.red, x.green, x.blue) for x in sa[0]],
                             [(0.0, 0.0, 0.0), (1.0, 2.0, 3.0)])

        def doit():
            t.from_param([MYCOLOR(0, 0, 0), MYCOLOR(1, 2, 3)])
        bytes = find_memleak(doit)
        self.failIf(bytes, "Leaks %d bytes" % bytes)

if __name__ == "__main__":
    unittest.main()
