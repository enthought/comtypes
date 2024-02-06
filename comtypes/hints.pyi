# This stub contains...
# - symbols those what might occur recursive imports in runtime.
# - utilities for type hints.
import sys
from typing import (
    Any,
    Callable,
    Generic,
    Iterator,
    List,
    NoReturn,
    Tuple,
    Type,
    TypeVar,
    Optional,
    overload,
    Union as _UnionT,
)

if sys.version_info >= (3, 8):
    from typing import Protocol
else:
    from typing_extensions import Protocol
if sys.version_info >= (3, 10):
    from typing import Concatenate, ParamSpec
else:
    from typing_extensions import Concatenate, ParamSpec
if sys.version_info >= (3, 11):
    from typing import Self
else:
    from typing_extensions import Self

from comtypes.automation import IDispatch as IDispatch, VARIANT as VARIANT
from comtypes.server import IClassFactory as IClassFactory
from comtypes.typeinfo import ITypeInfo as ITypeInfo

class _MethodTypeDesc(Protocol):
    arguments: List[Tuple[Any, str, List[str], Optional[Any]]]
    idlflags: List[str]
    name: str

_P_Put = ParamSpec("_P_Put")
_R_Put = TypeVar("_R_Put")
_P_PutRef = ParamSpec("_P_PutRef")
_R_PutRef = TypeVar("_R_PutRef")

def put_or_putref(
    put: Callable[_P_Put, _R_Put], putref: Callable[_P_PutRef, _R_PutRef]
) -> _UnionT[Callable[_P_Put, _R_Put], Callable[_P_PutRef, _R_PutRef]]: ...

_T_Inst = TypeVar("_T_Inst")
_T_SetVal = TypeVar("_T_SetVal")
_P_Get = ParamSpec("_P_Get")
_R_Get = TypeVar("_R_Get")
_P_Set = ParamSpec("_P_Set")

class _GetSetNormalProperty(Generic[_T_Inst, _R_Get, _T_SetVal]):
    fget: Callable[[_T_Inst], Any]
    fset: Callable[[_T_Inst, _T_SetVal], Any]

    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]) -> _R_Get: ...
    def __set__(self, instance: _T_Inst, value: _T_SetVal) -> None: ...

class _GetOnlyNormalProperty(Generic[_T_Inst, _R_Get]):
    fget: Callable[[_T_Inst], Any]

    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]) -> _R_Get: ...
    def __set__(self, instance: _T_Inst, value: Any) -> NoReturn: ...

class _SetOnlyNormalProperty(Generic[_T_Inst, _T_SetVal]):
    fget: Callable[[_T_Inst], Any]
    fset: Callable[[_T_Inst, _T_SetVal], Any]

    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(
        self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]
    ) -> NoReturn: ...
    def __set__(self, instance: _T_Inst, value: _T_SetVal) -> None: ...

@overload
def normal_property(
    fget: Callable[[_T_Inst], _R_Get],
) -> _GetOnlyNormalProperty[_T_Inst, _R_Get]: ...
@overload
def normal_property(
    *, fset: Callable[[_T_Inst, _T_SetVal], Any]
) -> _SetOnlyNormalProperty[_T_Inst, _T_SetVal]: ...
@overload
def normal_property(
    fget: Callable[[_T_Inst], _R_Get], fset: Callable[[_T_Inst, _T_SetVal], Any]
) -> _GetSetNormalProperty[_T_Inst, _R_Get, _T_SetVal]: ...

class _GetSetBoundNamedProperty(Generic[_T_Inst, _P_Get, _R_Get, _P_Set]):
    name: str
    fget: Callable[Concatenate[_T_Inst, _P_Get], _R_Get]
    fset: Callable[Concatenate[_T_Inst, _P_Set], Any]
    __doc__: Optional[str]
    def __getitem__(self, index: Any) -> _R_Get: ...
    def __call__(self, *args: _P_Get.args, **kwargs: _P_Get.kwargs) -> _R_Get: ...
    def __setitem__(self, index: Any, value: Any) -> None: ...
    def __iter__(self) -> NoReturn: ...

class _GetSetNamedProperty(Generic[_T_Inst, _P_Get, _R_Get, _P_Set]):
    name: str
    fget: Callable[Concatenate[_T_Inst, _P_Get], _R_Get]
    fset: Callable[Concatenate[_T_Inst, _P_Set], Any]
    __doc__: Optional[str]

    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(
        self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]
    ) -> _GetSetBoundNamedProperty[_T_Inst, _P_Get, _R_Get, _P_Set]: ...
    def __set__(self, instance: _T_Inst, value: Any) -> NoReturn: ...

class _GetOnlyBoundNamedProperty(Generic[_T_Inst, _P_Get, _R_Get]):
    name: str
    fget: Callable[Concatenate[_T_Inst, _P_Get], _R_Get]
    __doc__: Optional[str]
    def __getitem__(self, index: Any) -> _R_Get: ...
    def __call__(self, *args: _P_Get.args, **kwargs: _P_Get.kwargs) -> _R_Get: ...
    def __setitem__(self, index: Any, value: Any) -> NoReturn: ...
    def __iter__(self) -> NoReturn: ...

class _GetOnlyNamedProperty(Generic[_T_Inst, _P_Get, _R_Get]):
    name: str
    fget: Callable[Concatenate[_T_Inst, _P_Get], _R_Get]
    __doc__: Optional[str]

    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(
        self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]
    ) -> _GetOnlyBoundNamedProperty[_T_Inst, _P_Get, _R_Get]: ...
    def __set__(self, instance: _T_Inst, value: Any) -> NoReturn: ...

class _SetOnlyBoundNamedProperty(Generic[_T_Inst, _P_Set]):
    name: str
    fset: Callable[Concatenate[_T_Inst, _P_Set], Any]
    __doc__: Optional[str]
    def __getitem__(self, index: Any) -> NoReturn: ...
    def __call__(self, *args: Any, **kwargs: Any) -> NoReturn: ...
    def __setitem__(self, index: Any, value: Any) -> None: ...
    def __iter__(self) -> NoReturn: ...

class _SetOnlyNamedProperty(Generic[_T_Inst, _P_Set]):
    name: str
    fset: Callable[Concatenate[_T_Inst, _P_Set], Any]
    __doc__: Optional[str]

    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(
        self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]
    ) -> _SetOnlyBoundNamedProperty[_T_Inst, _P_Set]: ...
    def __set__(self, instance: _T_Inst, value: Any) -> NoReturn: ...

@overload
def named_property(
    name: str, fget: Callable[Concatenate[_T_Inst, _P_Get], _R_Get]
) -> _GetOnlyNamedProperty[_T_Inst, _P_Get, _R_Get]: ...
@overload
def named_property(
    name: str, *, fset: Callable[Concatenate[_T_Inst, _P_Set], Any]
) -> _SetOnlyNamedProperty[_T_Inst, _P_Set]: ...
@overload
def named_property(
    name: str,
    fget: Callable[Concatenate[_T_Inst, _P_Get], _R_Get],
    fset: Callable[Concatenate[_T_Inst, _P_Set], Any],
) -> _GetSetNamedProperty[_T_Inst, _P_Get, _R_Get, _P_Set]: ...

# for dunder methods those what be patched to ComInterface by metaclasses.

class _Descriptor(Protocol[_T_Inst, _R_Get]):
    @overload
    def __get__(self, instance: None, owner: Type[_T_Inst]) -> Self: ...
    @overload
    def __get__(self, instance: _T_Inst, owner: Optional[Type[_T_Inst]]) -> _R_Get: ...

# `__len__` for objects with `Count`
@overload
def to_dunder_len(count: _Descriptor[_T_Inst, int]) -> Callable[[_T_Inst], int]: ...
@overload
def to_dunder_len(count: Any) -> Callable[..., NoReturn]: ...

# `__iter__` for objects with `_NewEnum`
_T_E = TypeVar("_T_E")

@overload
def to_dunder_iter(
    newenum: _UnionT[
        _Descriptor[_T_Inst, Iterator[_T_E]], Callable[[_T_Inst], Iterator[_T_E]]
    ]
) -> Callable[[_T_Inst], Iterator[_T_E]]: ...
@overload
def to_dunder_iter(newenum: Any) -> Callable[..., NoReturn]: ...

# ... for objects with `Item`
# `__call__`
@overload
def to_dunder_call(
    item: _UnionT[
        _GetSetNamedProperty[_T_Inst, _P_Get, _R_Get, ...],
        _GetOnlyNamedProperty[_T_Inst, _P_Get, _R_Get],
        Callable[Concatenate[_T_Inst, _P_Get], _R_Get],
    ]
) -> Callable[Concatenate[_T_Inst, _P_Get], _R_Get]: ...
@overload
def to_dunder_call(item: Any) -> Callable[..., NoReturn]: ...

# `__getitem__`
@overload
def to_dunder_getitem(
    item: _UnionT[
        _GetSetNamedProperty[_T_Inst, _P_Get, _R_Get, ...],
        _GetOnlyNamedProperty[_T_Inst, _P_Get, _R_Get],
        Callable[Concatenate[_T_Inst, _P_Get], _R_Get],
    ]
) -> Callable[Concatenate[_T_Inst, _P_Get], _R_Get]: ...
@overload
def to_dunder_getitem(item: Any) -> Callable[..., NoReturn]: ...

# `__setitem__`
@overload
def to_dunder_setitem(
    item: _UnionT[
        _GetSetNamedProperty[_T_Inst, ..., Any, _P_Set],
        _SetOnlyNamedProperty[_T_Inst, _P_Set],
    ]
) -> Callable[Concatenate[_T_Inst, _P_Set], Any]: ...
@overload
def to_dunder_setitem(item: Any) -> Callable[..., NoReturn]: ...
