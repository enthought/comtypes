import unittest as ut

from ctypes import POINTER, byref, c_bool, c_ubyte
import comtypes
import comtypes.client

comtypes.client.GetModule("portabledeviceapi.dll")
from comtypes.gen.PortableDeviceApiLib import IStream


STGC_DEFAULT = 0
STREAM_SEEK_SET = 0
STREAM_SEEK_CUR = 1
STREAM_SEEK_END = 2


def _create_stream() -> IStream:
    # Create an IStream
    stream = POINTER(IStream)()  # type: ignore
    comtypes._ole32.CreateStreamOnHGlobal(None, c_bool(True), byref(stream))
    return stream  # type: ignore


class Test_RemoteWrite(ut.TestCase):
    def test_RemoteWrite(self):
        stream = _create_stream()
        test_data = "Some data".encode("utf-8")
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))

        written = stream.RemoteWrite(pv, len(test_data))

        # Verification
        self.assertEqual(written, len(test_data))


class Test_RemoteRead(ut.TestCase):
    def test_RemoteRead(self):
        stream = _create_stream()
        test_data = "Some data".encode("utf-8")
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        stream.RemoteWrite(pv, len(test_data))

        # Make sure the data actually gets written before trying to read back
        stream.Commit(STGC_DEFAULT)
        # Move the stream back to the beginning
        stream.RemoteSeek(0, STREAM_SEEK_SET)

        buffer_size = 1024

        read_buffer, data_read = stream.RemoteRead(buffer_size)

        # Verification
        self.assertEqual(data_read, len(test_data))
        self.assertEqual(bytearray(read_buffer)[0:data_read], test_data)


class Test_RemoteSeek(ut.TestCase):
    def _create_sample_stream(self) -> IStream:
        stream = _create_stream()
        test_data = b"spam egg bacon ham"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        stream.RemoteWrite(pv, len(test_data))
        stream.Commit(STGC_DEFAULT)
        return stream

    def test_takes_STREAM_SEEK_SET_as_origin(self):
        stream = self._create_sample_stream()
        newpos = stream.RemoteSeek(9, STREAM_SEEK_SET)
        self.assertEqual(newpos, 9)
        buf, read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], b"bacon ham")

    def test_takes_STREAM_SEEK_CUR_as_origin(self):
        stream = self._create_sample_stream()
        stream.RemoteSeek(8, STREAM_SEEK_SET)
        newpos = stream.RemoteSeek(7, STREAM_SEEK_CUR)
        self.assertEqual(newpos, 15)
        buf, read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], b"ham")

    def test_takes_STREAM_SEEK_END_as_origin(self):
        stream = self._create_sample_stream()
        stream.RemoteSeek(8, STREAM_SEEK_SET)
        newpos = stream.RemoteSeek(-13, STREAM_SEEK_END)
        self.assertEqual(newpos, 5)
        buf, read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], b"egg bacon ham")


if __name__ == "__main__":
    ut.main()
