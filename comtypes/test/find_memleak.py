import gc
from collections.abc import Callable
from ctypes import Structure, WinDLL, WinError, byref, c_size_t, c_void_p, sizeof
from ctypes import c_ulonglong as ULONGLONG
from ctypes.wintypes import BOOL, DWORD, HANDLE
from typing import TYPE_CHECKING, Any, Literal, TypeVar

if TYPE_CHECKING:
    from _ctypes import _CField
    from ctypes import _SimpleCData

    _CIntLikeT = TypeVar("_CIntLikeT", bound=_SimpleCData[int])
    _CIntLikeField = _CField[_CIntLikeT, int, _CIntLikeT | int]

################################################################


class PROCESS_MEMORY_COUNTERS(Structure):
    if TYPE_CHECKING:
        cb: _CIntLikeField[DWORD]
        PageFaultCount: _CIntLikeField[DWORD]
        PeakWorkingSetSize: _CIntLikeField[c_size_t]
        WorkingSetSize: _CIntLikeField[c_size_t]
        QuotaPeakPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPeakNonPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaNonPagedPoolUsage: _CIntLikeField[c_size_t]
        PagefileUsage: _CIntLikeField[c_size_t]
        PeakPagefileUsage: _CIntLikeField[c_size_t]

    _fields_ = [
        ("cb", DWORD),
        ("PageFaultCount", DWORD),
        ("PeakWorkingSetSize", c_size_t),
        ("WorkingSetSize", c_size_t),
        ("QuotaPeakPagedPoolUsage", c_size_t),
        ("QuotaPagedPoolUsage", c_size_t),
        ("QuotaPeakNonPagedPoolUsage", c_size_t),
        ("QuotaNonPagedPoolUsage", c_size_t),
        ("PagefileUsage", c_size_t),
        ("PeakPagefileUsage", c_size_t),
    ]

    def __init__(self):
        self.cb = sizeof(self)

    def dump(self):
        for n, _ in self._fields_[2:]:
            print(n, getattr(self, n) / 1e6)


class PROCESS_MEMORY_COUNTERS_EX(Structure):
    if TYPE_CHECKING:
        cb: _CIntLikeField[DWORD]
        PageFaultCount: _CIntLikeField[DWORD]
        PeakWorkingSetSize: _CIntLikeField[c_size_t]
        WorkingSetSize: _CIntLikeField[c_size_t]
        QuotaPeakPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPeakNonPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaNonPagedPoolUsage: _CIntLikeField[c_size_t]
        PagefileUsage: _CIntLikeField[c_size_t]
        PeakPagefileUsage: _CIntLikeField[c_size_t]
        PrivateUsage: _CIntLikeField[c_size_t]

    _fields_ = [
        ("cb", DWORD),
        ("PageFaultCount", DWORD),
        ("PeakWorkingSetSize", c_size_t),
        ("WorkingSetSize", c_size_t),
        ("QuotaPeakPagedPoolUsage", c_size_t),
        ("QuotaPagedPoolUsage", c_size_t),
        ("QuotaPeakNonPagedPoolUsage", c_size_t),
        ("QuotaNonPagedPoolUsage", c_size_t),
        ("PagefileUsage", c_size_t),
        ("PeakPagefileUsage", c_size_t),
        ("PrivateUsage", c_size_t),
    ]

    def __init__(self):
        self.cb = sizeof(self)


class PROCESS_MEMORY_COUNTERS_EX2(Structure):
    if TYPE_CHECKING:
        cb: _CIntLikeField[DWORD]
        PageFaultCount: _CIntLikeField[DWORD]
        PeakWorkingSetSize: _CIntLikeField[c_size_t]
        WorkingSetSize: _CIntLikeField[c_size_t]
        QuotaPeakPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaPeakNonPagedPoolUsage: _CIntLikeField[c_size_t]
        QuotaNonPagedPoolUsage: _CIntLikeField[c_size_t]
        PagefileUsage: _CIntLikeField[c_size_t]
        PeakPagefileUsage: _CIntLikeField[c_size_t]
        PrivateUsage: _CIntLikeField[c_size_t]
        PrivateWorkingSetSize: _CIntLikeField[c_size_t]
        SharedCommitUsage: _CIntLikeField[ULONGLONG]

    _fields_ = [
        ("cb", DWORD),
        ("PageFaultCount", DWORD),
        ("PeakWorkingSetSize", c_size_t),
        ("WorkingSetSize", c_size_t),
        ("QuotaPeakPagedPoolUsage", c_size_t),
        ("QuotaPagedPoolUsage", c_size_t),
        ("QuotaPeakNonPagedPoolUsage", c_size_t),
        ("QuotaNonPagedPoolUsage", c_size_t),
        ("PagefileUsage", c_size_t),
        ("PeakPagefileUsage", c_size_t),
        ("PrivateUsage", c_size_t),
        ("PrivateWorkingSetSize", c_size_t),
        ("SharedCommitUsage", ULONGLONG),
    ]

    def __init__(self):
        self.cb = sizeof(self)


_psapi = WinDLL("psapi")

_GetProcessMemoryInfo = _psapi.GetProcessMemoryInfo
# Use `c_void_p` to accept any of the `PROCESS_MEMORY_COUNTERS*` structures.
_GetProcessMemoryInfo.argtypes = [HANDLE, c_void_p, DWORD]
_GetProcessMemoryInfo.restype = BOOL

_PMI = (
    PROCESS_MEMORY_COUNTERS_EX2 | PROCESS_MEMORY_COUNTERS_EX | PROCESS_MEMORY_COUNTERS
)
_T_pmi = TypeVar("_T_pmi", bound=_PMI)


def _get_current_pmi(typ: type[_T_pmi]) -> tuple[OSError | None, _T_pmi]:
    pmi = typ()
    # -1 is a pseudo-handle representing the current process.
    issuccess = _GetProcessMemoryInfo(-1, byref(pmi), sizeof(pmi))
    return None if issuccess else WinError(), pmi


def _get_memory_info() -> _PMI:
    """Get process memory info using the most detailed structure available."""
    # Try EX2 first (Windows 10 22H2+ / Windows 11 22H2+)
    err_ex2, pmi_ex2 = _get_current_pmi(PROCESS_MEMORY_COUNTERS_EX2)
    if err_ex2 is None:
        return pmi_ex2
    # Try EX (Windows Vista+)
    err_ex, pmi_ex = _get_current_pmi(PROCESS_MEMORY_COUNTERS_EX)
    if err_ex is None:
        return pmi_ex
    # Fall back to basic
    err, pmi = _get_current_pmi(PROCESS_MEMORY_COUNTERS)
    if err is None:
        return pmi
    # We should use `ExceptionGroup` in Python 3.10+.
    raise err


def get_working_set_size() -> int:
    # Return the working set size (memory used by process)
    pmi = _get_memory_info()
    return pmi.WorkingSetSize


def get_private_usage() -> int:
    """Return the private memory usage (Commit Charge) in bytes.

    This is a better indicator of actual memory leaks than Working Set Size,
    because it represents the total amount of private memory committed for
    the process, regardless of whether pages are currently resident in RAM.
    """
    pmi = _get_memory_info()
    if isinstance(pmi, PROCESS_MEMORY_COUNTERS):
        # Fall back to PagefileUsage for basic structure
        # (same as PrivateUsage on older Windows)
        return pmi.PagefileUsage
    return pmi.PrivateUsage


LOOPS = 10, 1000


def find_memleak(
    func: Callable[[], Any],
    loops: tuple[int, int] = LOOPS,
    target: Literal["working_set", "private_usage"] = "working_set",
) -> int:
    """Detect memory leaks by measuring memory usage before and after repeated
    calls.

    Args:
        func: Callable to test for memory leaks.
        loops: Tuple of (outer_loops, inner_loops) for warmup and measurement
            phases.
        target: Method to use for leak detection:
            - "working_set": Use WorkingSetSize (physical RAM pages) - default
            - "private_usage": Use PrivateUsage (Commit Charge)

    Returns:
        The increase in memory usage in bytes (0 if no increase or decrease).

    Note:
        PrivateUsage (Commit Charge) is more reliable than WorkingSetSize for
        detecting actual memory leaks because:
        - WorkingSetSize only measures pages currently in physical RAM
        - PrivateUsage measures all committed private memory regardless of residency
        - An increase in PrivateUsage directly indicates unreleased allocations
    """
    # call 'func' several times, so that memory consumption
    # stabilizes:
    if target == "working_set":
        get_mem = get_working_set_size
    elif target == "private_usage":
        get_mem = get_private_usage
    else:
        raise TypeError(f"Unknown target: {target!r}")

    for j in range(loops[0]):
        for k in range(loops[1]):
            func()
    gc.collect()
    gc.collect()
    gc.collect()
    baseline = get_mem()
    # call 'func' several times, recording the difference in
    # memory consumption before and after the call.  Repeat this a
    # few times, and return a list containing the memory
    # consumption differences.
    for j in range(loops[0]):
        for k in range(loops[1]):
            func()
    gc.collect()
    gc.collect()
    gc.collect()
    # return the increased in process size
    result = get_mem() - baseline
    # Sometimes the process size did decrease, we do not report leaks
    # in this case:
    return max(result, 0)
