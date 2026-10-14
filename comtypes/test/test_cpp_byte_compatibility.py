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
class _NativeByteTest(unittest.TestCase):
    def setUp(self):
        self.server = CreateObject(
            CoByteCompatibilityTest,
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=IByteCompatibilityTest,
        )


class ByteCompatibilityTest(_NativeByteTest):
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


class LegacyMappingCompatibilityTest(_NativeByteTest):
    """Calls accepted by the old mapping that regress with the proposed mapping."""

    @unittest.expectedFailure
    def test_char_pointer_bytes_input(self):
        # Plain char* was previously inferred as STRING. Changing VT_I1 to
        # signed char loses this implicit bytes input, unlike explicit LPSTR.
        # Keep the compatibility regression visible for the review decision.
        self.assertEqual(self.server.ReadCharPointer(b"h"), 104)

    @unittest.expectedFailure
    def test_character_bytes_output_buffer(self):
        # Old STRING output accepts a char buffer. Reserve the terminator
        # locally because a plain char* marshals only one character.
        result = create_string_buffer(2)
        self.server._IByteCompatibilityTest__com_GetCharacterBytes(104, result)
        self.assertEqual(result.value, b"h")


class ProposedMappingCompatibilityTest(_NativeByteTest):
    """Numeric arguments and results supported by the proposed signed mapping."""

    def test_char_pointer_input(self):
        for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
            with self.subTest(value=value):
                self.assertEqual(
                    self.server.ReadCharPointer(pointer(c_byte(value))), expected
                )

    def test_char_reference_input(self):
        for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
            with self.subTest(value=value):
                self.assertEqual(
                    self.server.ReadCharPointer(byref(c_byte(value))), expected
                )

    def test_char_array_input(self):
        for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
            with self.subTest(value=value):
                self.assertEqual(
                    self.server.ReadCharPointer((c_byte * 1)(value)), expected
                )

    def test_character_pointer_output(self):
        for value, expected in ((0, 0), (1, 1), (104, 104), (128, -128), (255, -1)):
            with self.subTest(value=value):
                # Plain char* output uses the signed numeric mapping as well.
                self.assertEqual(self.server.GetCharacterBytes(value), expected)

    def test_character_pointer_output_buffer(self):
        for value, expected in ((0, 0), (104, 104), (128, -128), (255, -1)):
            with self.subTest(value=value):
                result = pointer(c_byte())
                self.server._IByteCompatibilityTest__com_GetCharacterBytes(
                    value, result
                )
                self.assertEqual(result.contents.value, expected)

    def test_character_reference_output_buffer(self):
        for value, expected in ((0, 0), (104, 104), (128, -128), (255, -1)):
            with self.subTest(value=value):
                result = c_byte()
                self.server._IByteCompatibilityTest__com_GetCharacterBytes(
                    value, byref(result)
                )
                self.assertEqual(result.value, expected)

    def test_character_array_output_buffer(self):
        for value, expected in ((0, 0), (104, 104), (128, -128), (255, -1)):
            with self.subTest(value=value):
                result = (c_byte * 1)()
                self.server._IByteCompatibilityTest__com_GetCharacterBytes(
                    value, result
                )
                self.assertEqual(result[0], expected)


if __name__ == "__main__":
    unittest.main()
