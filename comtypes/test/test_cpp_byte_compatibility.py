"""Byte and string compatibility through generated interfaces and native COM."""

import unittest
from ctypes import (
    byref,
    c_byte,
    c_char_p,
    c_void_p,
    cast,
    create_string_buffer,
    pointer,
)

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.malloc import _CoTaskMemFree

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen.ComtypesCppTestSrvLib import (
        CoByteCompatibilityTest,
        IByteCompatibilityTest,
    )

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class ByteCompatibilityTest(unittest.TestCase):
    def setUp(self):
        self.server = CreateObject(
            CoByteCompatibilityTest,
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=IByteCompatibilityTest,
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
                    self.server._IByteCompatibilityTest__com_EchoLpStr(
                        value, byref(result)
                    )
                    self.assertEqual(result.value, value)
                finally:
                    _CoTaskMemFree(cast(result, c_void_p))

    @unittest.expectedFailure
    def test_legacy_char_pointer_bytes(self):
        # Plain char* was previously inferred as STRING. Changing VT_I1 to
        # signed char loses this implicit bytes input, unlike explicit LPSTR.
        # Keep the compatibility regression visible for the review decision.
        self.assertEqual(self.server.ReadCharPointer(b"h"), ord("h"))

    def test_char_pointer_ctypes_inputs(self):
        for value in (
            pointer(c_byte(ord("h"))),
            byref(c_byte(ord("h"))),
            (c_byte * 1)(ord("h")),
        ):
            with self.subTest(value=value):
                self.assertEqual(self.server.ReadCharPointer(value), ord("h"))

    def test_character_pointer_output(self):
        for value in (0, 1, ord("h"), 128, 255):
            with self.subTest(value=value):
                # Plain char* output uses the signed numeric mapping as well.
                self.assertEqual(
                    self.server.GetCharacterBytes(value),
                    value if value < 128 else value - 256,
                )

    def test_character_ctypes_output_buffers(self):
        for value in (0, ord("h"), 128, 255):
            for result in (c_byte(), (c_byte * 1)()):
                with self.subTest(value=value, result=result):
                    argument = byref(result) if isinstance(result, c_byte) else result
                    self.server._IByteCompatibilityTest__com_GetCharacterBytes(
                        value, argument
                    )
                    actual = result.value if isinstance(result, c_byte) else result[0]
                    self.assertEqual(actual, value if value < 128 else value - 256)

    @unittest.expectedFailure
    def test_legacy_character_bytes_output(self):
        # Old STRING output accepts a char buffer. Reserve the terminator
        # locally because a plain char* marshals only one character.
        result = create_string_buffer(2)
        self.server._IByteCompatibilityTest__com_GetCharacterBytes(ord("h"), result)
        self.assertEqual(result.value, b"h")


if __name__ == "__main__":
    unittest.main()
