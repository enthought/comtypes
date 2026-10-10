"""Byte and string round trips through generated interfaces and native COM."""

import unittest
from ctypes import byref, c_char_p, c_void_p, cast

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.malloc import _CoTaskMemFree

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import CoByteEchoTest, IByteEchoTest

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class ByteEchoTest(unittest.TestCase):
    def setUp(self):
        self.server = CreateObject(
            CoByteEchoTest,
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=IByteEchoTest,
        )

    def test_negative_int8(self):
        for value in (-128, -7, -1):
            with self.subTest(value=value):
                self.assertEqual(self.server.EchoInt8(value), value)

    def test_nonnegative_int8(self):
        for value in (0, 1, 127):
            with self.subTest(value=value):
                self.assertEqual(self.server.EchoInt8(value), value)

    def test_uint8(self):
        for value in (0, 1, 42, 128, 255):
            with self.subTest(value=value):
                self.assertEqual(self.server.EchoUint8(value), value)

    def test_lpstr(self):
        for value in (b"", b"hello", b"\x80\xff"):
            with self.subTest(value=value):
                result = c_char_p()
                try:
                    # Use the generated raw method to retain the allocated address.
                    # The high-level result is bytes, which loses that address.
                    self.server._IByteEchoTest__com_EchoLpStr(value, byref(result))
                    self.assertEqual(result.value, value)
                finally:
                    _CoTaskMemFree(cast(result, c_void_p))

    @unittest.expectedFailure
    def test_legacy_char_pointer_bytes(self):
        # Plain char* was previously inferred as STRING. Changing VT_I1 to
        # signed char loses this implicit bytes input, unlike explicit LPSTR.
        # Keep the compatibility regression visible for the review decision.
        self.assertEqual(self.server.ReadCharPointer(b"h"), ord("h"))


if __name__ == "__main__":
    unittest.main()
