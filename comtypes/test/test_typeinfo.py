import os
import unittest
from ctypes import POINTER, byref
from comtypes import GUID, COMError
from comtypes.automation import DISPATCH_METHOD
from comtypes.typeinfo import LoadTypeLibEx, LoadRegTypeLib, \
     QueryPathOfRegTypeLib, TKIND_INTERFACE, TKIND_DISPATCH, TKIND_ENUM

# We should add other test cases for Windows CE.
if os.name == "nt":
    class Test(unittest.TestCase):
        # No LoadTypeLibEx on windows ce
        def test_LoadTypeLibEx(self):
            dllname = "scrrun.dll"
            self.assertRaises(WindowsError, lambda: LoadTypeLibEx("<xxx.xx>"))
            tlib = LoadTypeLibEx(dllname)
            self.assertTrue(tlib.GetTypeInfoCount())
            tlib.GetDocumentation(-1)
            self.assertEqual(tlib.IsName("idictionary"), "IDictionary")
            self.assertEqual(tlib.IsName("IDICTIONARY"), "IDictionary")
            self.assertTrue(tlib.FindName("IDictionary"))
            self.assertEqual(tlib.IsName("Spam"), None)
            tlib.GetTypeComp()

            attr = tlib.GetLibAttr()
            info = attr.guid, attr.wMajorVerNum, attr.wMinorVerNum
            other_tlib = LoadRegTypeLib(*info)
            self.assertEqual(tlib, other_tlib)

            for i in range(tlib.GetTypeInfoCount()):
                ti = tlib.GetTypeInfo(i)
                ti.GetTypeAttr()
                tlib.GetDocumentation(i)
                tlib.GetTypeInfoType(i)

                c_tlib, index = ti.GetContainingTypeLib()
                self.assertEqual(c_tlib, tlib)
                self.assertEqual(index, i)

            guid_null = GUID()
            self.assertRaises(COMError, lambda: tlib.GetTypeInfoOfGuid(guid_null))

            self.assertTrue(tlib.GetTypeInfoOfGuid(GUID("{42C642C1-97E1-11CF-978F-00A02463E06F}")))

            path = QueryPathOfRegTypeLib(*info)
            path = path.split("\0")[0]
            self.assertTrue(path.lower().endswith(dllname))

        def test_TypeInfo(self):
            tlib = LoadTypeLibEx("scrrun.dll")
            for index in range(tlib.GetTypeInfoCount()):
                ti = tlib.GetTypeInfo(index)
                ta = ti.GetTypeAttr()
                ti.GetDocumentation(-1)
                if ta.typekind in (TKIND_INTERFACE, TKIND_DISPATCH):
                    if ta.cImplTypes:
                        href = ti.GetRefTypeOfImplType(0)
                        base = ti.GetRefTypeInfo(href)
                        base.GetDocumentation(-1)
                        ti.GetImplTypeFlags(0)
                for f in range(ta.cFuncs):
                    fd = ti.GetFuncDesc(f)
                    names = ti.GetNames(fd.memid, 32)
                    ti.GetIDsOfNames(*names)
                    ti.GetMops(fd.memid)

                for v in range(ta.cVars):
                    ti.GetVarDesc(v)

if __name__ == "__main__":
    unittest.main()
