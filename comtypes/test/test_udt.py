import unittest
from ctypes import *
from comtypes.automation import VT_RECORD
from comtypes.typeinfo import IRecordInfo, GetRecordInfoFromGuids
##from comtypes.safearray import SafeArrayCreateVectorEx, SafeArrayPutElement
from comtypes.safearray import UnpackSafeArray

class UDTTest(unittest.TestCase):
    def test(self):
        # Test unpacking safearrays that contain UDT

        from comtypes.gen.TestComServerLib import MYCOLOR
        oleaut32 = WinDLL("oleaut32")

        ri = GetRecordInfoFromGuids(*MYCOLOR._recordinfo_)
        refcount_before = (ri.AddRef(), ri.Release())

        sa = oleaut32.SafeArrayCreateVectorEx(VT_RECORD, 0, 10, ri)
        for i in range(10):
            oleaut32.SafeArrayPutElement(sa, byref(c_long(i)), byref(MYCOLOR(i, i*2, i*3)))

        for i, item in enumerate(UnpackSafeArray(sa)):
            self.failUnlessEqual((item.red, item.green, item.blue),
                                 (i, i*2, i*3))

        oleaut32.SafeArrayDestroy(sa)
        refcount_after = (ri.AddRef(), ri.Release())

        self.failUnlessEqual(refcount_before, refcount_after)


if __name__ == "__main__":
    unittest.main()
