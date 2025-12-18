import unittest as ut
import winreg

import comtypes.client
from comtypes import typeinfo
from comtypes.automation import IDispatch

MSI_TLIB = typeinfo.LoadTypeLibEx("msi.dll")
comtypes.client.GetModule(MSI_TLIB)
import comtypes.gen.WindowsInstaller as msi

HKCR = 0


class Test_Installer(ut.TestCase):
    def test_hkcr_registry_value(self):
        # `WindowsInstaller.Installer` provides access to Windows configuration.
        inst = comtypes.client.CreateObject(
            "WindowsInstaller.Installer", interface=msi.Installer
        )
        IID_Installer = msi.Installer._iid_
        # This confirms that the Installer is a pure dispatch interface.
        self.assertIsInstance(inst, IDispatch)
        ti = MSI_TLIB.GetTypeInfoOfGuid(IID_Installer)
        ta = ti.GetTypeAttr()
        self.assertEqual(IID_Installer, ta.guid)
        self.assertFalse(ta.wTypeFlags & typeinfo.TYPEFLAG_FDUAL)
        # Both methods below get the "Programmatic Identifier" used to handle
        # ".txt" files.
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ".txt") as key:
            progid, _ = winreg.QueryValueEx(key, "")
        # This confirms that the Installer can correctly read system information.
        self.assertEqual(progid, inst.RegistryValue(HKCR, ".txt", ""))
