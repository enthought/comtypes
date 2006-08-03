import unittest as ut
from ctypes import POINTER
from comtypes.client import CoGetObject

# WMI has dual interfaces.
# Some methods/properties have "[out] POINTER(VARIANT)" parameters.
# This test checks that these parameters are returned as strings:
# that's what VARIANT.__ctypes_from_outparam__ does.
class Test(ut.TestCase):
    def test_wmi(self):
        wmi = CoGetObject("winmgmts:")
        disks = wmi.InstancesOf("Win32_LogicalDisk")

        # There are different typelibs installed for WMI on win2k and winXP.
        # WbemScripting refers to their guid:
        #   Win2k:
        #     import comtypes.gen._565783C6_CB41_11D1_8B02_00600806D9B6_0_1_1 as mod
        #   WinXP:
        #     import comtypes.gen._565783C6_CB41_11D1_8B02_00600806D9B6_0_1_2 as mod
        # So, the one that's referenced onm WbemScripting will be used, whether the actual
        # typelib is available or not.  XXX
        from comtypes.gen import WbemScripting
        WbemScripting.wbemPrivilegeCreateToken

        for item in disks:
            # obj[index] is forwarded to obj.Item(index)
            # .Value is a property with "[out] POINTER(VARIANT)" parameter.
            a = item.Properties_["Caption"].Value
            b = item.Properties_.Item("Caption").Value
            self.failUnlessEqual(a, b)
            self.failUnless(isinstance(a, basestring))
            self.failUnless(isinstance(b, basestring))
##            for prop in item.Properties_:
##                    print "\t", (prop.Name, prop.Value)
            self.failUnless(isinstance(item.Properties_["Description"].Value, unicode))
        # len(obj) is forwared to obj.Count
        c1 = len(disks)
        c2 = disks.Count
        self.failUnlessEqual(c1, c2)

if __name__ == "__main__":
    unittest.main()
