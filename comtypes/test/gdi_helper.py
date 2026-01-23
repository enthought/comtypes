from ctypes import POINTER, Structure, WinDLL, c_void_p
from ctypes.wintypes import (
    BOOL,
    DWORD,
    HANDLE,
    HDC,
    HGDIOBJ,
    HWND,
    INT,
    LONG,
    UINT,
    WORD,
)

_user32 = WinDLL("user32")

_GetDC = _user32.GetDC
_GetDC.argtypes = (HWND,)
_GetDC.restype = HDC

_ReleaseDC = _user32.ReleaseDC
_ReleaseDC.argtypes = (HWND, HDC)
_ReleaseDC.restype = INT

_gdi32 = WinDLL("gdi32")

_CreateCompatibleDC = _gdi32.CreateCompatibleDC
_CreateCompatibleDC.argtypes = (HDC,)
_CreateCompatibleDC.restype = HDC

_DeleteDC = _gdi32.DeleteDC
_DeleteDC.argtypes = (HDC,)
_DeleteDC.restype = BOOL

_SelectObject = _gdi32.SelectObject
_SelectObject.argtypes = (HDC, HGDIOBJ)
_SelectObject.restype = HGDIOBJ

_DeleteObject = _gdi32.DeleteObject
_DeleteObject.argtypes = (HGDIOBJ,)
_DeleteObject.restype = BOOL

_GdiFlush = _gdi32.GdiFlush
_GdiFlush.argtypes = []
_GdiFlush.restype = BOOL


class BITMAPINFOHEADER(Structure):
    _fields_ = [
        ("biSize", DWORD),
        ("biWidth", LONG),
        ("biHeight", LONG),
        ("biPlanes", WORD),
        ("biBitCount", WORD),
        ("biCompression", DWORD),
        ("biSizeImage", DWORD),
        ("biXPelsPerMeter", LONG),
        ("biYPelsPerMeter", LONG),
        ("biClrUsed", DWORD),
        ("biClrImportant", DWORD),
    ]


class BITMAPINFO(Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", DWORD * 1),  # Placeholder for color table, not used for 32bpp
    ]


_CreateDIBSection = _gdi32.CreateDIBSection
_CreateDIBSection.argtypes = (
    HDC,
    POINTER(BITMAPINFO),
    UINT,  # DIB_RGB_COLORS
    POINTER(c_void_p),  # lplpBits
    HANDLE,  # hSection
    DWORD,  # dwOffset
)
_CreateDIBSection.restype = HGDIOBJ
