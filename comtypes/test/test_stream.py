import contextlib
import ctypes
import struct
import unittest as ut
from collections.abc import Iterator
from ctypes import (
    HRESULT,
    POINTER,
    OleDLL,
    WinDLL,
    byref,
    c_size_t,
    c_ubyte,
    c_ulonglong,
    pointer,
)
from ctypes.wintypes import (
    BOOL,
    HDC,
    HGLOBAL,
    HWND,
    INT,
    LONG,
    LPVOID,
    UINT,
    ULARGE_INTEGER,
)
from typing import Optional

import comtypes.client
from comtypes import hresult

comtypes.client.GetModule("portabledeviceapi.dll")
# The stdole module is generated automatically during the portabledeviceapi
# module generation.
import comtypes.gen.stdole as stdole
from comtypes.gen.PortableDeviceApiLib import IStream

SIZE_T = c_size_t

STATFLAG_DEFAULT = 0
STGC_DEFAULT = 0
STGTY_STREAM = 2
STREAM_SEEK_SET = 0
STREAM_SEEK_CUR = 1
STREAM_SEEK_END = 2

_ole32 = OleDLL("ole32")

_CreateStreamOnHGlobal = _ole32.CreateStreamOnHGlobal
_CreateStreamOnHGlobal.argtypes = [HGLOBAL, BOOL, POINTER(POINTER(IStream))]
_CreateStreamOnHGlobal.restype = HRESULT

_shlwapi = OleDLL("shlwapi")

_IStream_Size = _shlwapi.IStream_Size
_IStream_Size.argtypes = [POINTER(IStream), POINTER(ULARGE_INTEGER)]
_IStream_Size.restype = HRESULT


def _create_stream(
    handle: Optional[int] = None, delete_on_release: bool = True
) -> IStream:
    # Create an IStream
    stream = POINTER(IStream)()  # type: ignore
    _CreateStreamOnHGlobal(handle, delete_on_release, byref(stream))
    return stream  # type: ignore


class Test_RemoteWrite(ut.TestCase):
    def test_RemoteWrite(self):
        stream = _create_stream()
        test_data = b"Some data"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))

        written = stream.RemoteWrite(pv, len(test_data))

        # Verification
        self.assertEqual(written, len(test_data))


class Test_RemoteRead(ut.TestCase):
    def test_RemoteRead(self):
        stream = _create_stream()
        test_data = b"Some data"
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


class Test_SetSize(ut.TestCase):
    def test_SetSize(self):
        stream = _create_stream()
        stream.SetSize(42)
        pui = pointer(c_ulonglong())
        _IStream_Size(stream, pui)
        self.assertEqual(pui.contents.value, 42)


class Test_RemoteCopyTo(ut.TestCase):
    def test_RemoteCopyTo(self):
        src = _create_stream()
        dst = _create_stream()
        test_data = b"parrot"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        src_written = src.RemoteWrite(pv, len(test_data))
        src.Commit(STGC_DEFAULT)
        src.RemoteSeek(0, STREAM_SEEK_SET)
        cpy_read, cpy_written = src.RemoteCopyTo(dst, src_written)
        self.assertEqual(cpy_read, len(test_data))
        self.assertEqual(cpy_written, len(test_data))
        dst.Commit(STGC_DEFAULT)
        dst.RemoteSeek(0, STREAM_SEEK_SET)
        dst_buf, dst_read = dst.RemoteRead(1024)
        self.assertEqual(bytearray(dst_buf)[0:dst_read], test_data)


class Test_Stat(ut.TestCase):
    # https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-istream-stat
    # https://learn.microsoft.com/en-us/windows/win32/api/objidl/ns-objidl-statstg
    def test_returns_statstg_from_no_modified_stream(self):
        stream = _create_stream()
        statstg = stream.Stat(STATFLAG_DEFAULT)
        self.assertIsNone(statstg.pwcsName)
        self.assertEqual(statstg.type, STGTY_STREAM)
        self.assertEqual(statstg.cbSize, 0)
        mt, ct, at = statstg.mtime, statstg.ctime, statstg.atime
        self.assertTrue(mt.dwLowDateTime == ct.dwLowDateTime == at.dwLowDateTime)
        self.assertTrue(mt.dwHighDateTime == ct.dwHighDateTime == at.dwHighDateTime)
        self.assertEqual(statstg.grfMode, 0)
        self.assertEqual(statstg.grfLocksSupported, 0)
        self.assertEqual(statstg.clsid, comtypes.GUID())
        self.assertEqual(statstg.grfStateBits, 0)


class Test_Clone(ut.TestCase):
    def test_Clone(self):
        orig = _create_stream()
        test_data = b"spam egg bacon ham"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        orig.RemoteWrite(pv, len(test_data))
        orig.Commit(STGC_DEFAULT)
        orig.RemoteSeek(0, STREAM_SEEK_SET)
        new_stm = orig.Clone()
        buf, read = new_stm.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], test_data)


_user32 = WinDLL("user32")

_GetDC = _user32.GetDC
_GetDC.argtypes = (HWND,)
_GetDC.restype = HDC

_ReleaseDC = _user32.ReleaseDC
_ReleaseDC.argtypes = (HWND, HDC)
_ReleaseDC.restype = INT

_kernel32 = WinDLL("kernel32")

_GlobalAlloc = _kernel32.GlobalAlloc
_GlobalAlloc.argtypes = (UINT, SIZE_T)
_GlobalAlloc.restype = HGLOBAL

_GlobalFree = _kernel32.GlobalFree
_GlobalFree.argtypes = (HGLOBAL,)
_GlobalFree.restype = HGLOBAL

_GlobalLock = _kernel32.GlobalLock
_GlobalLock.argtypes = (HGLOBAL,)
_GlobalLock.restype = LPVOID

_GlobalUnlock = _kernel32.GlobalUnlock
_GlobalUnlock.argtypes = (HGLOBAL,)
_GlobalUnlock.restype = BOOL

_oleaut32 = WinDLL("oleaut32")

_OleLoadPicture = _oleaut32.OleLoadPicture
_OleLoadPicture.argtypes = (
    POINTER(IStream),  # lpstm
    LONG,  # lSize
    BOOL,  # fSave
    POINTER(comtypes.GUID),  # riid
    POINTER(POINTER(comtypes.IUnknown)),  # ppvObj
)
_OleLoadPicture.restype = HRESULT

# Constants for the type of a picture object
PICTYPE_BITMAP = 1

GMEM_FIXED = 0x0000
GMEM_ZEROINIT = 0x0040

BI_RGB = 0  # No compression


@contextlib.contextmanager
def get_dc(hwnd: int) -> Iterator[int]:
    """Context manager to get and release a device context (DC)."""
    dc = _GetDC(hwnd)
    assert dc, "Failed to get device context."
    try:
        yield dc
    finally:
        # Release the device context
        _ReleaseDC(hwnd, dc)


@contextlib.contextmanager
def global_alloc(uflags: int, dwbytes: int) -> Iterator[int]:
    """Context manager to allocate and free a global memory handle."""
    handle = _GlobalAlloc(uflags, dwbytes)
    assert handle, "Failed to GlobalAlloc"
    try:
        yield handle
    finally:
        _GlobalFree(handle)


@contextlib.contextmanager
def global_lock(handle: int) -> Iterator[int]:
    """Context manager to lock a global memory handle and obtain a pointer."""
    lp_mem = _GlobalLock(handle)
    assert lp_mem, "Failed to GlobalLock"
    try:
        yield lp_mem
    finally:
        _GlobalUnlock(handle)


def create_24bit_pixel_data(
    red: int,
    green: int,
    blue: int,
    width: int,
    height: int,
) -> bytes:
    # Generates width x height pixel 24-bit BGR BMP binary data with 0 DPI.
    SIZEOF_BITMAPFILEHEADER = 14
    SIZEOF_BITMAPINFOHEADER = 40
    pixel_data = b""
    for _ in range(height):
        # Each row is padded to a 4-byte boundary.
        # For 24bpp, each pixel is 3 bytes (BGR).
        # Row size without padding: width * 3 bytes
        row_size = width * 3
        # Calculate padding bytes (to make row_size a multiple of 4)
        padding_bytes = (4 - (row_size % 4)) % 4
        for _ in range(width):
            # B, G, R
            pixel_data += struct.pack(b"BBB", blue, green, red)
        pixel_data += b"\x00" * padding_bytes
    BITMAP_DATA_OFFSET = SIZEOF_BITMAPFILEHEADER + SIZEOF_BITMAPINFOHEADER
    file_size = BITMAP_DATA_OFFSET + len(pixel_data)
    bmp_header = struct.pack(
        b"<2sIHHI",
        b"BM",  # File type signature "BM"
        file_size,  # Total file size
        0,  # Reserved1
        0,  # Reserved2
        BITMAP_DATA_OFFSET,  # Offset to pixel data
    )
    info_header = struct.pack(
        b"<IiiHHIIiiII",
        SIZEOF_BITMAPINFOHEADER,  # Size of BITMAPINFOHEADER
        width,  # Image width
        height,  # Image height
        1,  # Planes
        24,  # Bits per pixel (for BGR)
        BI_RGB,  # Compression
        len(pixel_data),  # Size of image data
        # Set to 0 DPI (0 px/m) is focusing on pixel data integrity;
        # this ensures environment-independent results and not reliably preserve resolution metadata.
        0,  # X pixels per meter (0 DPI)
        0,  # Y pixels per meter (0 DPI)
        # Setting biClrUsed and biClrImportant to 0 signifies that the bitmap
        # uses the maximum number of colors for the given bit depth (2^24) and
        # that all pixels are essential for rendering.
        0,  # Colors used
        0,  # Colors important
    )
    return bmp_header + info_header + pixel_data


class Test_Picture(ut.TestCase):
    def test_ole_load_picture(self):
        width, height = 1, 1
        data = create_24bit_pixel_data(255, 0, 0, width, height)  # Red pixel
        # Allocate global memory with `GMEM_FIXED` (fixed-size) and
        # `GMEM_ZEROINIT` (initialize to zero) and copy BMP data.
        with global_alloc(GMEM_FIXED | GMEM_ZEROINIT, len(data)) as handle:
            with global_lock(handle) as lp_mem:
                ctypes.memmove(lp_mem, data, len(data))
            pstm = _create_stream(handle, delete_on_release=False)
            # Load picture from the stream
            pic: stdole.IPicture = POINTER(stdole.IPicture)()  # type: ignore
            hr = _OleLoadPicture(
                pstm,
                len(data),  # lSize
                False,  # fSave
                byref(stdole.IPicture._iid_),
                byref(pic),
            )
            self.assertEqual(hr, hresult.S_OK)
            self.assertEqual(pic.Type, PICTYPE_BITMAP)
            pstm.RemoteSeek(0, STREAM_SEEK_SET)
            buf, read = pstm.RemoteRead(len(data))
        self.assertEqual(bytes(buf)[:read], data)


if __name__ == "__main__":
    ut.main()
