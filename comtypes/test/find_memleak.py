import gc
from collections.abc import Callable
from ctypes import Structure, WinDLL, WinError, byref, c_size_t, c_void_p, sizeof
from ctypes import c_ulonglong as ULONGLONG
from ctypes.wintypes import BOOL, DWORD, HANDLE
from typing import TYPE_CHECKING, Any, Literal, Protocol, TypeVar

from comtypes import COMObject, hresult
from comtypes.malloc import IMallocSpy, _CoRegisterMallocSpy, _CoRevokeMallocSpy

if TYPE_CHECKING:
    from _ctypes import _CField
    from ctypes import _Pointer, _SimpleCData

    from comtypes import hints  # type: ignore

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


class BytesTrackingMallocSpy(COMObject):
    _com_interfaces_ = [IMallocSpy]

    def __init__(self):
        super().__init__()
        self._sizes: dict[int, int] = {}  # pointer -> allocated size
        self._pending: int = 0  # size from PreAlloc/PreRealloc
        self._allocated: int = 0  # cumulative allocated bytes
        self._freed: int = 0  # cumulative freed bytes

    def IMallocSpy_PreAlloc(self, this: int, cbRequest: int) -> int:
        self._pending = cbRequest
        return cbRequest

    def IMallocSpy_PostAlloc(self, this: int, pActual: int) -> int:
        if pActual:
            self._sizes[pActual] = self._pending
            self._allocated += self._pending
        return pActual

    def IMallocSpy_PreFree(self, this: int, pRequest: int, fSpyed: int) -> int:
        if fSpyed and pRequest:
            self._freed += self._sizes.pop(pRequest, 0)
        return pRequest

    def IMallocSpy_PostFree(self, this: int, fSpyed: int) -> None: ...

    def IMallocSpy_PreRealloc(
        self,
        this: int,
        pRequest: int,
        cbRequest: int,
        ppNewRequest: "_Pointer[c_void_p]",
        fSpyed: int,
    ) -> int:
        self._pending = cbRequest
        if fSpyed and pRequest:
            self._freed += self._sizes.pop(pRequest, 0)
        # `ppNewRequest` is an [out] parameter that `PreRealloc` must write.
        # Write `pRequest` to pass through the original pointer unchanged.
        # (A spy that extends/modifies the allocation would write a different
        # pointer here.)
        # See also https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-imallocspy-prerealloc
        if ppNewRequest:
            ppNewRequest[0] = pRequest
        return cbRequest

    def IMallocSpy_PostRealloc(self, this: int, pActual: int, fSpyed: int) -> int:
        if pActual:
            self._sizes[pActual] = self._pending
            self._allocated += self._pending
        return pActual

    def IMallocSpy_PreGetSize(self, this: int, pRequest: int, fSpyed: int) -> int:
        return pRequest

    def IMallocSpy_PostGetSize(self, this: int, cbActual: int, fSpyed: int) -> int:
        return cbActual

    def IMallocSpy_PreDidAlloc(self, this: int, pRequest: int, fSpyed: int) -> int:
        return pRequest

    def IMallocSpy_PostDidAlloc(
        self, this: int, pRequest: int, fSpyed: int, fActual: int
    ) -> int:
        return fActual

    def IMallocSpy_PreHeapMinimize(self, this: int) -> None: ...

    def IMallocSpy_PostHeapMinimize(self, this: int) -> None: ...

    def get_allocated_bytes(self) -> int:
        return self._allocated

    def get_freed_bytes(self) -> int:
        return self._freed

    def get_net_bytes(self) -> int:
        return self._allocated - self._freed


class MemoryUsageDiffDetector(Protocol):
    """Protocol for context managers that track memory usage difference.

    While inside the context manager, `get_diff()` returns the difference
    between the current memory usage.
    """

    def __enter__(self) -> "hints.Self": ...

    def __exit__(self, exc_type, exc_val, exc_tb) -> None: ...

    def get_diff(self) -> int: ...


class ProcessMemoryDiffDetector:
    def __init__(self, get_memory_func: Callable[[], int]):
        self._get_memory_func = get_memory_func

    def __enter__(self) -> "ProcessMemoryDiffDetector":
        # Measure baseline
        self._baseline = self._get_memory_func()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None: ...

    def get_diff(self) -> int:
        return self._get_memory_func() - self._baseline


class ComAllocationDiffDetector:
    def __init__(self):
        self._spy = BytesTrackingMallocSpy()

    def __enter__(self) -> "ComAllocationDiffDetector":
        _CoRegisterMallocSpy(self._spy)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        # Avoid failing here when revoking the spy raises an exception because
        # some memory it tracked has not been freed yet.
        err = None
        for _ in (0, 1):
            try:
                # `CoRevokeMallocSpy` fails with `E_ACCESSDENIED` while any
                # memory block the spy tracked has not been freed yet.
                _CoRevokeMallocSpy()
            except OSError as e:
                err = e
                if e.winerror != hresult.E_ACCESSDENIED:
                    raise e
                # Run `gc.collect()` instead of calling `CoTaskMemFree` here.
                # This detector frees the blocks by releasing the Python
                # objects that own them, avoiding a use-after-free.
                gc.collect()
                # `get_diff()` is expected to be called inside the `with`
                # block, so this cleanup does not affect the measured leak.
            if err is None:
                return
            # Try to revoke the spy again, expecting the memory blocks to have
            # been freed.
        raise err

    def get_diff(self) -> int:
        return self._spy.get_net_bytes()


def create_memdiff_detector(
    target: Literal["working_set", "private_usage", "mallocspy"],
) -> MemoryUsageDiffDetector:
    if target == "working_set":
        return ProcessMemoryDiffDetector(get_working_set_size)
    elif target == "private_usage":
        return ProcessMemoryDiffDetector(get_private_usage)
    elif target == "mallocspy":
        return ComAllocationDiffDetector()
    raise TypeError(f"Unknown target: {target!r}")


LOOPS = 10, 1000


def find_memleak(
    func: Callable[[], Any],
    loops: tuple[int, int] = LOOPS,
    target: Literal["working_set", "private_usage", "mallocspy"] = "working_set",
) -> int:
    """Detect memory leaks by measuring memory usage before and after repeated
    calls.

    Args:
        func: Callable to test for memory leaks.
        loops: Tuple of (outer_loops, inner_loops) for warmup and measurement
               phases.
        target: Method to use for leak detection:
            - 'working_set': Use WorkingSetSize (physical RAM pages) - default
            - 'private_usage': Use PrivateUsage (Commit Charge)
            - 'mallocspy': Use IMallocSpy to track COM allocations

    Returns:
        The increase in memory usage in bytes (0 if no increase or decrease).
        For 'mallocspy', returns the net allocated bytes (allocated - freed).

    Note:
        PrivateUsage (Commit Charge) is more reliable than WorkingSetSize for
        detecting actual memory leaks because:
          - Measures all committed private memory regardless of residency
          - Indicates directly unreleased allocations
    """
    # call 'func' several times, so that memory consumption
    # stabilizes:
    for j in range(loops[0]):
        for k in range(loops[1]):
            func()
    gc.collect()
    gc.collect()
    gc.collect()
    # call 'func' several times, recording the difference in
    # memory consumption before and after the call.  Repeat this a
    # few times, and return a list containing the memory
    # consumption differences.
    with create_memdiff_detector(target) as d:
        # Measurement phase
        for j in range(loops[0]):
            for k in range(loops[1]):
                func()
        gc.collect()
        gc.collect()
        gc.collect()
        # Sometimes the process size did decrease, we do not report leaks
        # in this case:
        return max(d.get_diff(), 0)
