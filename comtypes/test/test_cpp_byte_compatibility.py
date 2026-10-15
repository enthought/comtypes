"""Compare old and proposed byte mappings across the native COM boundary."""

import unittest
from ctypes import (
    HRESULT,
    POINTER,
    ArgumentError,
    byref,
    c_byte,
    c_char,
    c_char_p,
    c_ubyte,
    c_void_p,
    cast,
    create_string_buffer,
    pointer,
)
from typing import Any, Protocol

from comtypes import CLSCTX_LOCAL_SERVER, COMMETHOD, GUID, IUnknown
from comtypes.client import CreateObject, GetModule
from comtypes.malloc import _CoTaskMemFree

try:
    GetModule(("{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}", 1, 0, 0))
    from comtypes.gen import ComtypesCppTestSrvLib

    IMPORT_SERVER_FAILED = False
except (ImportError, OSError):
    IMPORT_SERVER_FAILED = True


STRING = c_char_p

IID_IByteCompatibilityTest = GUID("{047E34C1-33A8-4DAA-A1E5-B941756A8319}")
PROPOSED_COMMETHODS = [
    COMMETHOD(
        [],
        HRESULT,
        "EchoInt8",
        (["in"], c_byte, "value"),
        (["out", "retval"], POINTER(c_byte), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "EchoUint8",
        (["in"], c_ubyte, "value"),
        (["out", "retval"], POINTER(c_ubyte), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "EchoLpStr",
        (["in"], STRING, "value"),
        (["out", "retval"], POINTER(STRING), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "ReadCharPointer",
        (["in"], POINTER(c_byte), "value"),
        (["out", "retval"], POINTER(c_ubyte), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetCharacterBytes",
        (["in"], c_ubyte, "value"),
        (["out", "retval"], POINTER(c_byte), "result"),
    ),
]
LEGACY_COMMETHODS = [
    COMMETHOD(
        [],
        HRESULT,
        "EchoInt8",
        (["in"], c_char, "value"),
        (["out", "retval"], STRING, "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "EchoUint8",
        (["in"], c_ubyte, "value"),
        (["out", "retval"], POINTER(c_ubyte), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "EchoLpStr",
        (["in"], STRING, "value"),
        (["out", "retval"], POINTER(STRING), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "ReadCharPointer",
        (["in"], STRING, "value"),
        (["out", "retval"], POINTER(c_ubyte), "result"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetCharacterBytes",
        (["in"], c_ubyte, "value"),
        (["out", "retval"], STRING, "result"),
    ),
]

# The generated definitions are checked in ByteCompatibilityTest below.


class _PByteCompatibilityTest(Protocol):
    def EchoInt8(self, value: Any) -> Any: ...
    def EchoUint8(self, value: Any) -> Any: ...
    def EchoLpStr(self, value: Any) -> Any: ...
    def ReadCharPointer(self, value: Any) -> Any: ...
    def GetCharacterBytes(self, value: Any) -> Any: ...


def create_legacy_itf() -> type[_PByteCompatibilityTest]:
    class IByteCompatibilityTest(IUnknown):
        _case_insensitive_ = True
        _iid_ = IID_IByteCompatibilityTest
        _idlflags_ = ["oleautomation"]
        _methods_ = LEGACY_COMMETHODS

        def EchoLpStr(self, value):
            result = c_char_p()
            try:
                self.__com_EchoLpStr(value, byref(result))
                return result.value
            finally:
                _CoTaskMemFree(cast(result, c_void_p))

        def GetCharacterBytes(self, value):
            result = create_string_buffer(2)
            self.__com_GetCharacterBytes(value, result)
            return result.value

    return IByteCompatibilityTest


def create_proposed_itf() -> type[_PByteCompatibilityTest]:
    class IByteCompatibilityTest(IUnknown):
        _case_insensitive_ = True
        _iid_ = IID_IByteCompatibilityTest
        _idlflags_ = ["oleautomation"]
        _methods_ = PROPOSED_COMMETHODS

        def EchoLpStr(self, value):
            result = c_char_p()
            try:
                self.__com_EchoLpStr(value, byref(result))
                return result.value
            finally:
                _CoTaskMemFree(cast(result, c_void_p))

    return IByteCompatibilityTest


def create_proposed_char_ptr_output_buf() -> type[_PByteCompatibilityTest]:
    class IByteCompatibilityTest(IUnknown):
        _case_insensitive_ = True
        _iid_ = IID_IByteCompatibilityTest
        _idlflags_ = ["oleautomation"]
        _methods_ = PROPOSED_COMMETHODS

        def EchoLpStr(self, value):
            result = c_char_p()
            try:
                self.__com_EchoLpStr(value, byref(result))
                return result.value
            finally:
                _CoTaskMemFree(cast(result, c_void_p))

        def GetCharacterBytes(self, value):
            result = pointer(c_byte())
            self.__com_GetCharacterBytes(value, result)
            return result.contents.value

    return IByteCompatibilityTest


def create_proposed_char_ref_output_buf() -> type[_PByteCompatibilityTest]:
    class IByteCompatibilityTest(IUnknown):
        _case_insensitive_ = True
        _iid_ = IID_IByteCompatibilityTest
        _idlflags_ = ["oleautomation"]
        _methods_ = PROPOSED_COMMETHODS

        def EchoLpStr(self, value):
            result = c_char_p()
            try:
                self.__com_EchoLpStr(value, byref(result))
                return result.value
            finally:
                _CoTaskMemFree(cast(result, c_void_p))

        def GetCharacterBytes(self, value):
            result = c_byte()
            self.__com_GetCharacterBytes(value, byref(result))
            return result.value

    return IByteCompatibilityTest


def create_proposed_arr_output_buf() -> type[_PByteCompatibilityTest]:
    class IByteCompatibilityTest(IUnknown):
        _case_insensitive_ = True
        _iid_ = IID_IByteCompatibilityTest
        _idlflags_ = ["oleautomation"]
        _methods_ = PROPOSED_COMMETHODS

        def EchoLpStr(self, value):
            result = c_char_p()
            try:
                self.__com_EchoLpStr(value, byref(result))
                return result.value
            finally:
                _CoTaskMemFree(cast(result, c_void_p))

        def GetCharacterBytes(self, value):
            result = (c_byte * 1)()
            self.__com_GetCharacterBytes(value, result)
            return result[0]

    return IByteCompatibilityTest


def create_server(interface: type[_PByteCompatibilityTest]) -> _PByteCompatibilityTest:
    return CreateObject(
        ComtypesCppTestSrvLib.CoByteCompatibilityTest,
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=interface,
    )


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class ByteCompatibilityTest(unittest.TestCase):
    """Compare saved legacy and proposed interfaces in the same native calls."""

    def setUp(self):
        self.legacy = create_server(create_legacy_itf())
        self.proposed = {
            "automatic": create_server(create_proposed_itf()),
            "pointer": create_server(create_proposed_char_ptr_output_buf()),
            "reference": create_server(create_proposed_char_ref_output_buf()),
            "array": create_server(create_proposed_arr_output_buf()),
        }

    def test_generated_mapping(self):
        self.assertEqual(
            ComtypesCppTestSrvLib.IByteCompatibilityTest._methods_, PROPOSED_COMMETHODS
        )

    def test_legacy_int8_implicit_output(self):
        # STRING cannot allocate a scalar output. Keep that legacy behavior
        # visible rather than silently adapting it to the proposed result type.
        for value in (b"\x80", b"\xf9", b"\xff", b"\0", b"\x01", b"\x7f"):
            with self.subTest(mapping="legacy", value=value):
                with self.assertRaises(TypeError):
                    self.legacy.EchoInt8(value)

    def test_legacy_int8_explicit_output(self):
        for value in (b"\x80", b"\xf9", b"\xff", b"\0", b"\x01", b"\x7f"):
            with self.subTest(mapping="legacy", value=value):
                result = create_string_buffer(2)
                self.legacy._IByteCompatibilityTest__com_EchoInt8(value, result)
                self.assertEqual(result.raw[:1], value)

    def test_negative_int8(self):
        for name, server in self.proposed.items():
            for value in (-128, -7, -1):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoInt8(value), value)

    def test_nonnegative_int8(self):
        for name, server in self.proposed.items():
            for value in (0, 1, 127):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoInt8(value), value)

    def test_uint8(self):
        for name, server in [("legacy", self.legacy), *self.proposed.items()]:
            for value in (0, 1, 42, 128, 255):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoUint8(value), value)

    def test_lpstr(self):
        for name, server in [("legacy", self.legacy), *self.proposed.items()]:
            for value in (b"", b"hello", b"\x80\xff"):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoLpStr(value), value)

    def test_char_pointer_bytes_input(self):
        for value, expected in (
            (b"\0", 0),
            (b"h", 104),
            (b"\x80", 128),
            (b"\xff", 255),
        ):
            with self.subTest(mapping="legacy", value=value):
                self.assertEqual(self.legacy.ReadCharPointer(value), expected)
            for name, server in self.proposed.items():
                with self.subTest(mapping=name, value=value):
                    with self.assertRaises(ArgumentError):
                        server.ReadCharPointer(value)

    def test_char_pointer_input(self):
        for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
            argument = pointer(c_byte(value))
            with self.subTest(mapping="legacy", value=value):
                with self.assertRaises(ArgumentError):
                    self.legacy.ReadCharPointer(argument)
            for name, server in self.proposed.items():
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.ReadCharPointer(argument), expected)

    def test_char_reference_input(self):
        for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
            argument = byref(c_byte(value))
            with self.subTest(mapping="legacy", value=value):
                with self.assertRaises(ArgumentError):
                    self.legacy.ReadCharPointer(argument)
            for name, server in self.proposed.items():
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.ReadCharPointer(argument), expected)

    def test_char_array_input(self):
        for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
            argument = (c_byte * 1)(value)
            with self.subTest(mapping="legacy", value=value):
                with self.assertRaises(ArgumentError):
                    self.legacy.ReadCharPointer(argument)
            for name, server in self.proposed.items():
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.ReadCharPointer(argument), expected)

    def test_character_bytes_output(self):
        # Reserve a local terminator; native plain char* writes one byte only.
        for value, expected in (
            (0, b""),
            (1, b"\x01"),
            (104, b"h"),
            (128, b"\x80"),
            (255, b"\xff"),
        ):
            with self.subTest(mapping="legacy", value=value):
                self.assertEqual(self.legacy.GetCharacterBytes(value), expected)

    def test_character_numeric_output(self):
        for name, server in self.proposed.items():
            for value, expected in ((0, 0), (1, 1), (104, 104), (128, -128), (255, -1)):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.GetCharacterBytes(value), expected)


@unittest.skipIf(IMPORT_SERVER_FAILED, "Requires the C++ out-of-process COM server.")
class LegacyMappingCompatibilityTest(unittest.TestCase):
    """Old calling styles still fail on the proposed generated interface."""

    def setUp(self):
        self.server = create_server(ComtypesCppTestSrvLib.IByteCompatibilityTest)

    @unittest.expectedFailure
    def test_char_pointer_bytes_input(self):
        self.assertEqual(self.server.ReadCharPointer(b"h"), 104)

    @unittest.expectedFailure
    def test_character_bytes_output_buffer(self):
        result = create_string_buffer(2)
        self.server._IByteCompatibilityTest__com_GetCharacterBytes(104, result)
        self.assertEqual(result.value, b"h")


if __name__ == "__main__":
    unittest.main()
