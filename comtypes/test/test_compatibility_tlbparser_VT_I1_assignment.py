"""Compare the VT_I1 mappings across the native COM boundary.

Mapping and compatibility discussion: https://github.com/enthought/comtypes/issues/969
"""

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

# The generated definitions are checked in TestGeneratedMapping below.


class _PByteCompatibilityTest(Protocol):
    def EchoInt8(self, value: Any) -> Any: ...
    def EchoUint8(self, value: Any) -> Any: ...
    def EchoLpStr(self, value: Any) -> Any: ...
    def ReadCharPointer(self, value: Any) -> Any: ...
    def GetCharacterBytes(self, value: Any) -> Any: ...


def create_legacy_automatic_int8_output() -> type[_PByteCompatibilityTest]:
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


def create_legacy_buffered_int8_output() -> type[_PByteCompatibilityTest]:
    class IByteCompatibilityTest(create_legacy_automatic_int8_output()):
        def EchoInt8(self, value):
            # STRING cannot allocate a scalar output. Supply character storage
            # and retain one byte, including a possible NUL value.
            result = create_string_buffer(2)
            self._IByteCompatibilityTest__com_EchoInt8(value, result)
            return result.raw[:1]

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
    def setUp(self):
        self.legacy_automatic = create_server(create_legacy_automatic_int8_output())
        self.legacy_buffered = create_server(create_legacy_buffered_int8_output())
        self.proposed = {
            "automatic": create_server(create_proposed_itf()),
            "pointer": create_server(create_proposed_char_ptr_output_buf()),
            "reference": create_server(create_proposed_char_ref_output_buf()),
            "array": create_server(create_proposed_arr_output_buf()),
            "generated": create_server(ComtypesCppTestSrvLib.IByteCompatibilityTest),
        }


class TestGeneratedMapping(ByteCompatibilityTest):
    def test_generated_mapping(self):
        self.assertEqual(
            ComtypesCppTestSrvLib.IByteCompatibilityTest._methods_, PROPOSED_COMMETHODS
        )
        # The legacy mapping describes different Python argument/result types.
        self.assertNotEqual(
            ComtypesCppTestSrvLib.IByteCompatibilityTest._methods_, LEGACY_COMMETHODS
        )


class TestEchoInt8(ByteCompatibilityTest):
    def test_takes_and_returns_as_intended(self):
        for value, expected in (
            (b"\x80", b"\x80"),
            (b"\xf9", b"\xf9"),
            (b"\xff", b"\xff"),
            (b"\0", b"\0"),
            (b"\x01", b"\x01"),
            (b"\x7f", b"\x7f"),
        ):
            with self.subTest(mapping="legacy_buffered", value=value):
                self.assertEqual(self.legacy_buffered.EchoInt8(value), expected)
        for name, server in self.proposed.items():
            for value, expected in (
                (-128, -128),
                (-7, -7),
                (-1, -1),
                (0, 0),
                (1, 1),
                (127, 127),
            ):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoInt8(value), expected)

    def test_raises_type_error(self):
        for value in (b"\x80", b"\xf9", b"\xff", b"\0", b"\x01", b"\x7f"):
            with self.subTest(mapping="legacy_automatic", value=value):
                with self.assertRaises(TypeError):
                    self.legacy_automatic.EchoInt8(value)

    def test_raises_argument_error(self):
        for name, server in self.proposed.items():
            for value in (b"\x80", b"h", b"\xff"):
                with self.subTest(mapping=name, value=value):
                    with self.assertRaises(ArgumentError):
                        server.EchoInt8(value)


class TestEchoUint8(ByteCompatibilityTest):
    def test_takes_and_returns_as_intended(self):
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
            *self.proposed.items(),
        ]:
            for value, expected in ((0, 0), (1, 1), (42, 42), (128, 128), (255, 255)):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoUint8(value), expected)

    def test_raises_argument_error(self):
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
            *self.proposed.items(),
        ]:
            with self.subTest(mapping=name):
                with self.assertRaises(ArgumentError):
                    server.EchoUint8(object())


class TestEchoLpStr(ByteCompatibilityTest):
    def test_takes_and_returns_as_intended(self):
        # Use helpers that free the native CoTaskMem allocation after copying it.
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
            *[(n, s) for n, s in self.proposed.items() if n != "generated"],
        ]:
            for value, expected in (
                (b"", b""),
                (b"hello", b"hello"),
                (b"\x80\xff", b"\x80\xff"),
            ):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.EchoLpStr(value), expected)

    def test_raises_argument_error(self):
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
            *self.proposed.items(),
        ]:
            with self.subTest(mapping=name):
                with self.assertRaises(ArgumentError):
                    server.EchoLpStr(object())


class TestReadCharPointer(ByteCompatibilityTest):
    def test_takes_and_returns_as_intended(self):
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
        ]:
            for value, expected in (
                (b"\0", 0),
                (b"h", 104),
                (b"\x80", 128),
                (b"\xff", 255),
            ):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.ReadCharPointer(value), expected)
        for name, server in self.proposed.items():
            for value, expected in ((0, 0), (104, 104), (-128, 128), (-1, 255)):
                for argument in (
                    pointer(c_byte(value)),
                    byref(c_byte(value)),
                    (c_byte * 1)(value),
                ):
                    with self.subTest(mapping=name, value=value, argument=argument):
                        self.assertEqual(server.ReadCharPointer(argument), expected)

    def test_raises_argument_error(self):
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
        ]:
            for argument in (
                pointer(c_byte(104)),
                byref(c_byte(104)),
                (c_byte * 1)(104),
            ):
                with self.subTest(mapping=name, argument=argument):
                    with self.assertRaises(ArgumentError):
                        server.ReadCharPointer(argument)
        # This includes the generated interface: bytes input compatibility is
        # changed, even though signed-byte pointers now work as intended.
        for name, server in self.proposed.items():
            for value in (b"\0", b"h", b"\x80", b"\xff"):
                with self.subTest(mapping=name, value=value):
                    with self.assertRaises(ArgumentError):
                        server.ReadCharPointer(value)


class TestGetCharacterBytes(ByteCompatibilityTest):
    def test_takes_and_returns_as_intended(self):
        for name, server in [
            ("legacy_automatic", self.legacy_automatic),
            ("legacy_buffered", self.legacy_buffered),
        ]:
            for value, expected in (
                (0, b""),
                (1, b"\x01"),
                (104, b"h"),
                (128, b"\x80"),
                (255, b"\xff"),
            ):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.GetCharacterBytes(value), expected)
        for name, server in self.proposed.items():
            for value, expected in ((0, 0), (1, 1), (104, 104), (128, -128), (255, -1)):
                with self.subTest(mapping=name, value=value):
                    self.assertEqual(server.GetCharacterBytes(value), expected)
                    # The new signed numeric result is not the old bytes value.
                    self.assertNotEqual(server.GetCharacterBytes(value), bytes([value]))

    def test_raises_argument_error(self):
        for name, server in self.proposed.items():
            with self.subTest(mapping=name):
                result = create_string_buffer(2)
                # Legacy character storage is incompatible with POINTER(c_byte).
                with self.assertRaises(ArgumentError):
                    server._IByteCompatibilityTest__com_GetCharacterBytes(104, result)


if __name__ == "__main__":
    unittest.main()
