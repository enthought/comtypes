import unittest

from comtypes.client import CreateObject, GetModule
from comtypes.test.find_memleak import find_memleak

# LIBID from AvmcIfc.idl: library AVMCIFCLib
AVMCIFCLib_LIBID = "{70577167-ED71-4977-B719-2C40C6DD8E1D}"

try:
    GetModule((AVMCIFCLib_LIBID, 1, 0, 0))
    IMPORT_FAILED = False
except (ImportError, OSError):
    # AvmcIfc type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


@unittest.skipIf(
    IMPORT_FAILED, "This depends on the AvmcIfc COM server being registered."
)
class Test(unittest.TestCase):
    """Test COM records"""

    def test(self):
        # The ATL COM dll
        avmc = CreateObject("AvmcIfc.Avmc.1")

        # This returns an array (a list) of DeviceInfo records.
        devs = avmc.FindAllAvmc()

        self.assertEqual(devs[0].Flags, 12)
        self.assertEqual(devs[0].ID, 13)
        self.assertEqual(devs[0].LocId, 14)
        self.assertEqual(devs[0].Description, "Avmc")
        self.assertEqual(devs[0].SerialNumber, "1234")

        self.assertEqual(devs[1].Flags, 22)
        self.assertEqual(devs[1].ID, 23)
        self.assertEqual(devs[1].LocId, 24)
        self.assertEqual(devs[1].Description, "Avmc2")
        self.assertEqual(devs[1].SerialNumber, "5678")

        # # Leaks... where?
        # def doit():
        #     avmc.FindAllAvmc()
        # self.check_leaks(doit)

    def check_leaks(self, func, limit=0):
        bytes = find_memleak(func)
        self.assertFalse(bytes > limit, f"Leaks {bytes} bytes")


if __name__ == "__main__":
    unittest.main()
