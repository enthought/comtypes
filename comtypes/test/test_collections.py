import unittest
from comtypes.client import CreateObject, GetModule #, Constants
from ctypes import ArgumentError

from comtypes.test.find_memleak import find_memleak

class Test(unittest.TestCase):

    def test_IEnumVARIANT(self):
        # The XP firewall manager.
        fwmgr = CreateObject('HNetCfg.FwMgr')
        # apps has a _NewEnum property that implements IEnumVARIANT
        apps = fwmgr.LocalPolicy.CurrentProfile.AuthorizedApplications

        self.failUnlessEqual(apps.Count, len(apps))

        cv = iter(apps)

        names = [p.ProcessImageFileName for p in cv]
        self.failUnlessEqual(len(apps), len(names))

        # The iterator is consumed now:
        self.failUnlessEqual([p.ProcessImageFileName for p in cv], [])

        # But we can reset it:
        cv.Reset()
        self.failUnlessEqual([p.ProcessImageFileName for p in cv], names)

        # Reset, then skip:
        cv.Reset()
        cv.Skip(3)
        self.failUnlessEqual([p.ProcessImageFileName for p in cv], names[3:])

        # Reset, then skip:
        cv.Reset()
        cv.Skip(300)
        self.failUnlessEqual([p.ProcessImageFileName for p in cv], names[300:])

        # Hm, do we want to allow random access to the iterator?
        # Should the iterator support __getitem__ ???
        self.failUnlessEqual(cv[0].ProcessImageFileName, names[0])
        self.failUnlessEqual(cv[0].ProcessImageFileName, names[0])
        self.failUnlessEqual(cv[0].ProcessImageFileName, names[0])

        if len(names) > 1:
            self.failUnlessEqual(cv[1].ProcessImageFileName, names[1])
            self.failUnlessEqual(cv[1].ProcessImageFileName, names[1])
            self.failUnlessEqual(cv[1].ProcessImageFileName, names[1])

        # We can now call Next(celt) with celt != 1, the call always returns a list:
        cv.Reset()
        self.failUnlessEqual(names[:3], [p.ProcessImageFileName for p in cv.Next(3)])

        # calling Next(0) makes no sense, but should work anyway:
        self.failUnlessEqual(cv.Next(0), [])

        cv.Reset()
        self.failUnlessEqual(len(cv.Next(len(names) * 2)), len(names))

        # slicing is not (yet?) supported
        cv.Reset()
        self.failUnlessRaises(ArgumentError, lambda: cv[:])

    def test_leaks_1(self):
        # The XP firewall manager.
        fwmgr = CreateObject('HNetCfg.FwMgr')
        # apps has a _NewEnum property that implements IEnumVARIANT
        apps = fwmgr.LocalPolicy.CurrentProfile.AuthorizedApplications

        def doit():
            for item in iter(apps):
                item.ProcessImageFileName
        bytes = find_memleak(doit, (2, 20))
        self.failIf(bytes, "Leaks %d bytes" % bytes)

    def test_leaks_2(self):
        # The XP firewall manager.
        fwmgr = CreateObject('HNetCfg.FwMgr')
        # apps has a _NewEnum property that implements IEnumVARIANT
        apps = fwmgr.LocalPolicy.CurrentProfile.AuthorizedApplications

        def doit():
            iter(apps).Next(99)
        bytes = find_memleak(doit, (2, 20))
        self.failIf(bytes, "Leaks %d bytes" % bytes)

    def test_leaks_3(self):
        # The XP firewall manager.
        fwmgr = CreateObject('HNetCfg.FwMgr')
        # apps has a _NewEnum property that implements IEnumVARIANT
        apps = fwmgr.LocalPolicy.CurrentProfile.AuthorizedApplications

        def doit():
            for i in range(2):
                for what in iter(apps):
                    pass
        bytes = find_memleak(doit, (2, 20))
        self.failIf(bytes, "Leaks %d bytes" % bytes)

if __name__ == "__main__":
    unittest.main()
