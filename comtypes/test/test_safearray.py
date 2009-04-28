import unittest
from decimal import Decimal
import datetime
from ctypes import *
from ctypes.wintypes import BOOL
from comtypes.test.find_memleak import find_memleak
from comtypes import BSTR, IUnknown
from comtypes.test import is_resource_enabled
import array

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
        a = array.array("d", (3.14, 2.78))
        v = VARIANT(a)
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_R8)
        self.failUnlessEqual(tuple(a.tolist()), v.value)

        def func():
            v = VARIANT(array.array("d", (3.14, 2.78)))

        bytes = find_memleak(func)
        self.failIf(bytes, "Leaks %d bytes" % bytes)


    def test_float_array(self):
        a = array.array("f", (3.14, 2.78))
        v = VARIANT(a)
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_R4)
        self.failUnlessEqual(tuple(a.tolist()), v.value)

    def test_2dim_array(self):
        data = ((1, 2, 3, 4),
                (5, 6, 7, 8),
                (9, 10, 11, 12))
        v = VARIANT(data)
        self.failUnlessEqual(v.value, data)

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

    def test_array(self):
        t = _midlSAFEARRAY(c_double)
        pat = pointer(t())

        try:
            import numpy
        except ImportError:
            pass # numpy not available
        else:
            pat[0] = numpy.zeros(32, dtype=numpy.float)
            self.failUnlessEqual(tuple(pat[0][0]),
                                 (0.0,) * 32)

            data = ((1.0, 2.0, 3.0),
                    (4.0, 5.0, 6.0),
                    (7.0, 8.0, 9.0))
            a = numpy.array(data,
                            dtype=numpy.double)
            pat[0] = a
            self.failUnlessEqual(pat[0][0],
                                 data)

            data = ((1.0, 2.0), (3.0, 4.0), (5.0, 6.0))
            a = numpy.array(data,
                            dtype=numpy.double,
                            order="F")
            pat[0] = a
            self.failUnlessEqual(pat[0][0],
                                 data)

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

    def test_VT_UNKNOWN_1(self):
        a = _midlSAFEARRAY(POINTER(IUnknown))
        t = _midlSAFEARRAY(POINTER(IUnknown))
        self.failUnless(a is t)

        def com_refcnt(o):
            "Return the COM refcount of an interface pointer"
            import gc; gc.collect(); gc.collect()
            o.AddRef()
            return o.Release()

        from comtypes.typeinfo import CreateTypeLib, ICreateTypeLib
        punk = CreateTypeLib("spam").QueryInterface(IUnknown) # will never be saved to disk

        # initial refcount
        initial = com_refcnt(punk)

        # This should increase the refcount by 1
        sa = t.from_param([punk])
        self.failUnlessEqual(initial + 1, com_refcnt(punk))

        # Unpacking the array must not change the refcount, and must
        # return an equal object.
        self.failUnlessEqual((punk,), sa[0])
        self.failUnlessEqual(initial + 1, com_refcnt(punk))

        del sa
        self.failUnlessEqual(initial, com_refcnt(punk))

        sa = t.from_param([None])
        self.failUnlessEqual((POINTER(IUnknown)(),), sa[0])


    def test_VT_UNKNOWN_multi(self):
        a = _midlSAFEARRAY(POINTER(IUnknown))
        t = _midlSAFEARRAY(POINTER(IUnknown))
        self.failUnless(a is t)

        def com_refcnt(o):
            "Return the COM refcount of an interface pointer"
            import gc; gc.collect(); gc.collect()
            o.AddRef()
            return o.Release()

        from comtypes.typeinfo import CreateTypeLib, ICreateTypeLib
        punk = CreateTypeLib("spam").QueryInterface(IUnknown) # will never be saved to disk

        # initial refcount
        initial = com_refcnt(punk)

        # This should increase the refcount by 4
        sa = t.from_param((punk,) * 4)
        self.failUnlessEqual(initial + 4, com_refcnt(punk))

        # Unpacking the array must not change the refcount, and must
        # return an equal object.
        self.failUnlessEqual((punk,)*4, sa[0])
        self.failUnlessEqual(initial + 4, com_refcnt(punk))

        del sa
        self.failUnlessEqual(initial, com_refcnt(punk))

        # This should increase the refcount by 2
        sa = t.from_param((punk, None, punk, None))
        self.failUnlessEqual(initial + 2, com_refcnt(punk))

        null = POINTER(IUnknown)()
        self.failUnlessEqual((punk, null, punk, null), sa[0])

        del sa
        self.failUnlessEqual(initial, com_refcnt(punk))

        # repeat same test, with 2 different com pointers

        plib = CreateTypeLib("foo")
        a, b = com_refcnt(plib), com_refcnt(punk)
        sa = t.from_param([plib, punk, plib])

####        self.failUnlessEqual((plib, punk, plib), sa[0])
        self.failUnlessEqual((a+2, b+1), (com_refcnt(plib), com_refcnt(punk)))

        del sa
        self.failUnlessEqual((a, b), (com_refcnt(plib), com_refcnt(punk)))

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

if is_resource_enabled("pythoncom"):
    try:
        import pythoncom
    except ImportError:
        # pywin32 not installed...
        pass
    else:
        # pywin32 is available.  The pythoncom dll contains two handy
        # exported functions that allow to create a VARIANT from a Python
        # object, also a function that unpacks a VARIANT into a Python
        # object.
        #
        # This allows us to create und unpack SAFEARRAY instances
        # contained in VARIANTs, and check for consistency with the
        # comtypes code.

        _dll = PyDLL(pythoncom.__file__)

        # c:/sf/pywin32/com/win32com/src/oleargs.cpp 213
        # PyObject *PyCom_PyObjectFromVariant(const VARIANT *var)
        unpack = _dll.PyCom_PyObjectFromVariant
        unpack.restype = py_object
        unpack.argtypes = POINTER(VARIANT),

        # c:/sf/pywin32/com/win32com/src/oleargs.cpp 54
        # BOOL PyCom_VariantFromPyObject(PyObject *obj, VARIANT *var)
        _pack = _dll.PyCom_VariantFromPyObject
        _pack.argtypes = py_object, POINTER(VARIANT)
        _pack.restype = BOOL
        def pack(obj):
            var = VARIANT()
            result = _pack(obj, byref(var))
            return var

        class PyWinTest(unittest.TestCase):
            def test_1dim(self):
                data = (1, 2, 3)
                variant = pack(data)
                self.failUnlessEqual(variant.value, data)
                self.failUnlessEqual(unpack(variant), data)

            def test_2dim(self):
                data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
                variant = pack(data)
                self.failUnlessEqual(variant.value, data)
                self.failUnlessEqual(unpack(variant), data)

            def test_3dim(self):
                data = ( ( (1, 2), (3, 4), (5, 6) ),
                         ( (7, 8), (9, 10), (11, 12) ) )
                variant = pack(data)
                self.failUnlessEqual(variant.value, data)
                self.failUnlessEqual(unpack(variant), data)

            def test_4dim(self):
                data = ( ( ( ( 1,  2), ( 3,  4) ),
                           ( ( 5,  6), ( 7,  8) ) ),
                         ( ( ( 9, 10), (11, 12) ),
                           ( (13, 14), (15, 16) ) ) )
                variant = pack(data)
                self.failUnlessEqual(variant.value, data)
                self.failUnlessEqual(unpack(variant), data)

if __name__ == "__main__":
    unittest.main()
