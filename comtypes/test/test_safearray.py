import unittest as ut
from comtypes.automation import VARIANT, VT_ARRAY, VT_VARIANT, VT_I4, VT_R4, VT_R8

class TestCase(ut.TestCase):
    def test_1(self):
        v = VARIANT()
        v.value = ((1, 2, 3), ("foo", "bar", None))
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_VARIANT)
        self.failUnlessEqual(v.value, ((1, 2, 3), ("foo", "bar", None)))

    def test_double_array(self):
        import array
        a = array.array("d", (3.14, 2.78))
        v = VARIANT(a)
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_R8)
        self.failUnlessEqual(tuple(a.tolist()), v.value)

    def test_float_array(self):
        import array
        a = array.array("f", (3.14, 2.78))
        v = VARIANT(a)
        self.failUnlessEqual(v.vt, VT_ARRAY | VT_R4)
        self.failUnlessEqual(tuple(a.tolist()), v.value)

if __name__ == "__main__":
    unittest.main()
