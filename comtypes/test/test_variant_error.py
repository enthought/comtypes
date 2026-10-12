"""VT_ERROR VARIANT round-trip tests via the native C++ COM test server."""

import unittest
from _ctypes import COMError

from comtypes import CLSCTX_LOCAL_SERVER, hresult
from comtypes.automation import VARIANT
from comtypes.client import CreateObject, GetModule

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import (
        CoVariantErrorTest,
        IVariantErrorTest,
    )

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True


def _create_server() -> "IVariantErrorTest":
    return CreateObject(
        CoVariantErrorTest,
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=IVariantErrorTest,
    )


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestCreateVariantError(unittest.TestCase):
    def setUp(self):
        self.server = _create_server()

    def test_returns_error_code(self):
        for errcode in [
            hresult.DISP_E_TYPEMISMATCH,
            hresult.DISP_E_BADINDEX,
            hresult.DISP_E_UNKNOWNNAME,
            hresult.DISP_E_OVERFLOW,
        ]:
            with self.subTest(errcode=errcode):
                errvar = self.server.CreateVariantError(errcode)
                self.assertEqual(errcode, errvar.value)


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class TestGetErrorCode(unittest.TestCase):
    def setUp(self):
        self.server = _create_server()

    def test_raises_type_mismatch(self):
        """Passing a NON VT_ERROR VARIANT raises COMError(DISP_E_TYPEMISMATCH)."""
        for value in [VARIANT(123), VARIANT("hello"), VARIANT()]:
            with self.subTest(value=value):
                with self.assertRaises(COMError) as cm:
                    self.server.GetErrorCode(value)
                self.assertEqual(cm.exception.hresult, hresult.DISP_E_TYPEMISMATCH)

    def test_returns_errorcode(self):
        self.assertEqual(
            self.server.GetErrorCode(VARIANT.missing), hresult.DISP_E_PARAMNOTFOUND
        )
        self.assertEqual(VARIANT.missing.value, hresult.DISP_E_PARAMNOTFOUND)


if __name__ == "__main__":
    unittest.main()
