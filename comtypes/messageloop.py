import ctypes
from ctypes import WinDLL, WinError, byref
from ctypes.wintypes import MSG
from typing import TYPE_CHECKING, List, SupportsIndex

if TYPE_CHECKING:
    from ctypes import _CArgObject
    from typing import Any, Callable, Iterable

    _FilterCallable = Callable[["_CArgObject"], Iterable[Any]]  # type: ignore

_user32 = WinDLL("user32")

GetMessage = _user32.GetMessageA
GetMessage.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint]
TranslateMessage = _user32.TranslateMessage
DispatchMessage = _user32.DispatchMessageA


class _MessageLoop:
    def __init__(self) -> None:
        self._filters: List["_FilterCallable"] = []

    def insert_filter(self, obj: "_FilterCallable", index: SupportsIndex = -1) -> None:
        self._filters.insert(index, obj)

    def remove_filter(self, obj: "_FilterCallable") -> None:
        self._filters.remove(obj)

    def run(self) -> None:
        msg = MSG()
        lpmsg = byref(msg)
        while 1:
            ret = GetMessage(lpmsg, 0, 0, 0)
            if ret == -1:
                raise WinError()
            elif ret == 0:
                return  # got WM_QUIT
            if not self.filter_message(lpmsg):
                TranslateMessage(lpmsg)
                DispatchMessage(lpmsg)

    def filter_message(self, lpmsg: "_CArgObject") -> bool:
        return any(list(filter(lpmsg)) for filter in self._filters)


_messageloop = _MessageLoop()

run = _messageloop.run
insert_filter = _messageloop.insert_filter
remove_filter = _messageloop.remove_filter

__all__ = ["run", "insert_filter", "remove_filter"]
