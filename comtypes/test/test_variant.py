import unittest, os, sys
from ctypes import *
from comtypes import IUnknown, GUID
from comtypes.automation import VARIANT, DISPPARAMS
from comtypes.automation import VT_NULL, VT_EMPTY, VT_ERROR
from comtypes.automation import VT_I1, VT_I2, VT_I4, VT_I8
from comtypes.automation import VT_UI1, VT_UI2, VT_UI4, VT_UI8
from comtypes.automation import VT_R4, VT_R8
from comtypes.automation import BSTR, VT_BSTR, VT_DATE
from comtypes.typeinfo import LoadTypeLibEx, LoadRegTypeLib
from comtypes.test import is_resource_enabled

def get_refcnt(comptr):
    # return the COM reference count of a COM interface pointer
    if not comptr:
        return 0
    comptr.AddRef()
    return comptr.Release()

class VariantTestCase(unittest.TestCase):

    def test_constants(self):
        empty = VARIANT.empty
        self.failUnlessEqual(empty.vt, VT_EMPTY)
        self.failUnless(empty.value is None)

        null = VARIANT.null
        self.failUnlessEqual(null.vt, VT_NULL)
        self.failUnless(null.value is None)

        missing = VARIANT.missing
        self.failUnlessEqual(missing.vt, VT_ERROR)
        self.assertRaises(NotImplementedError, lambda: missing.value)

    def test_com_refcounts(self):
        # typelib for oleaut32
        tlb = LoadRegTypeLib(GUID("{00020430-0000-0000-C000-000000000046}"), 2, 0, 0)
        rc = get_refcnt(tlb)

        p = tlb.QueryInterface(IUnknown)
        self.failUnlessEqual(get_refcnt(tlb), rc+1)

        del p
        self.failUnlessEqual(get_refcnt(tlb), rc)

    def test_com_pointers(self):
        # Storing a COM interface pointer in a VARIANT increments the refcount,
        # changing the variant to contain something else decrements it
        tlb = LoadRegTypeLib(GUID("{00020430-0000-0000-C000-000000000046}"), 2, 0, 0)
        rc = get_refcnt(tlb)

        v = VARIANT(tlb)
        self.failUnlessEqual(get_refcnt(tlb), rc+1)

        p = v.value
        self.failUnlessEqual(get_refcnt(tlb), rc+2)
        del p
        self.failUnlessEqual(get_refcnt(tlb), rc+1)

        v.value = None
        self.failUnlessEqual(get_refcnt(tlb), rc)

    def test_null_com_pointers(self):
        p = POINTER(IUnknown)()
        self.failUnlessEqual(get_refcnt(p), 0)

        v = VARIANT(p)
        self.failUnlessEqual(get_refcnt(p), 0)

    def test_dispparams(self):
        # DISPPARAMS is a complex structure, well worth testing.
        d = DISPPARAMS()
        d.rgvarg = (VARIANT * 3)()
        values = [1, 5, 7]
        for i, v in enumerate(values):
            d.rgvarg[i].value = v
        result = [d.rgvarg[i].value for i in range(3)]
        self.failUnlessEqual(result, values)

    def test_pythonobjects(self):
        objects = [None, 42, 3.14, True, False, "abc", u"abc", 7L]
        for x in objects:
            v = VARIANT(x)
            self.failUnlessEqual(x, v.value)

    def test_integers(self):
        v = VARIANT()

        if (hasattr(sys, "maxint")):
            # this test doesn't work in Python 3000
            v.value = sys.maxint
            self.failUnlessEqual(v.value, sys.maxint)
            self.failUnlessEqual(type(v.value), int)

            v.value += 1
            self.failUnlessEqual(v.value, sys.maxint+1)
            self.failUnlessEqual(type(v.value), long)

        v.value = 1L
        self.failUnlessEqual(v.value, 1)
        self.failUnlessEqual(type(v.value), int)

    def test_datetime(self):
        import datetime
        now = datetime.datetime.now()

        v = VARIANT()
        v.value = now
        self.failUnlessEqual(v.vt, VT_DATE)
        self.failUnlessEqual(v.value, now)

    def test_BSTR(self):
        v = VARIANT()
        v.value = u"abc\x00123\x00"
        self.failUnlessEqual(v.value, "abc\x00123\x00")

        v.value = None
        # manually clear the variant
        v._.VT_I4 = 0

        # NULL pointer BSTR should be handled as empty string
        v.vt = VT_BSTR
        self.failUnless(v.value in ("", None))

    def test_ctypes_in_variant(self):
        v = VARIANT()
        objs = [(c_ubyte(3), VT_UI1),
                (c_char("x"), VT_UI1),
                (c_byte(3), VT_I1),
                (c_ushort(3), VT_UI2),
                (c_short(3), VT_I2),
                (c_uint(3), VT_UI4),
                (c_int(3), VT_I4),
                (c_double(3.14), VT_R8),
                (c_float(3.14), VT_R4),
                ]
        for value, vt in objs:
            v.value = value
            self.failUnlessEqual(v.vt, vt)

class ArrayTest(unittest.TestCase):
    def test_double(self):
        import array
        for typecode in "df":
            # because of FLOAT rounding errors, whi will only work for
            # certain values!
            a = array.array(typecode, (1.0, 2.0, 3.0, 4.5))
            v = VARIANT()
            v.value = a
            self.failUnlessEqual(v.value, (1.0, 2.0, 3.0, 4.5))

    def test_int(self):
        import array
        for typecode in "bhiBHIlL":
            a = array.array(typecode, (1, 1, 1, 1))
            v = VARIANT()
            v.value = a
            self.failUnlessEqual(v.value, (1, 1, 1, 1))

################################################################
def run_test(rep, msg, func=None, previous={}, results={}):
##    items = [None] * rep
    if func is None:
        locals = sys._getframe(1).f_locals
        func = eval("lambda: %s" % msg, locals)
    items = xrange(rep)
    from time import clock
    start = clock()
    for i in items:
        func(); func(); func(); func(); func()
    stop = clock()
    duration = (stop-start)*1e6/5/rep
    try:
        prev = previous[msg]
    except KeyError:
        print >> sys.stderr, "%40s: %7.1f us" % (msg, duration)
        delta = 0.0
    else:
        delta = duration / prev * 100.0
        print >> sys.stderr, "%40s: %7.1f us, time = %5.1f%%" % (msg, duration, delta)
    results[msg] = duration
    return delta

def check_perf(rep=20000):
    from ctypes import c_int
    from comtypes.automation import VARIANT

    import cPickle
    try:
        previous = cPickle.load(open("result.pickle", "rb"))
    except IOError:
        previous = {}

    results = {}

    d = 0.0
    d += run_test(rep, "VARIANT()", previous=previous, results=results)
    d += run_test(rep, "VARIANT().value", previous=previous, results=results)
    d += run_test(rep, "VARIANT(None).value", previous=previous, results=results)
    d += run_test(rep, "VARIANT(42).value", previous=previous, results=results)
    d += run_test(rep, "VARIANT(42L).value", previous=previous, results=results)
    d += run_test(rep, "VARIANT(3.14).value", previous=previous, results=results)
    d += run_test(rep, "VARIANT(u'Str').value", previous=previous, results=results)
    d += run_test(rep, "VARIANT('Str').value", previous=previous, results=results)
    d += run_test(rep, "VARIANT((42,)).value", previous=previous, results=results)
    d += run_test(rep, "VARIANT([42,]).value", previous=previous, results=results)

    print "Average duration %.1f%%" % (d / 10)
##    cPickle.dump(results, open("result.pickle", "wb"))

if __name__ == '__main__':
    try:
        unittest.main()
    except SystemExit:
        pass
    import comtypes
    print "Running benchmark with comtypes %s/Python %s ..." % (comtypes.__version__, sys.version.split()[0],)
    check_perf()
