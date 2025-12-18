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
        installer = comtypes.client.CreateObject(
            "WindowsInstaller.Installer", interface=msi.Installer
        )
        IID_Installer = msi.Installer._iid_
        # This confirms that the Installer is a pure dispatch interface.
        self.assertIsInstance(installer, IDispatch)
        ti = MSI_TLIB.GetTypeInfoOfGuid(IID_Installer)
        ta = ti.GetTypeAttr()
        self.assertEqual(IID_Installer, ta.guid)
        self.assertFalse(ta.wTypeFlags & typeinfo.TYPEFLAG_FDUAL)
        # Both methods below get the "Programmatic Identifier" used to handle
        # ".txt" files.
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ".txt") as key:
            winreg_val, _ = winreg.QueryValueEx(key, "")
        msi_val = installer.RegistryValue(HKCR, ".txt", "")
        # This confirms that the Installer can correctly read system information.
        self.assertEqual(winreg_val, msi_val)
