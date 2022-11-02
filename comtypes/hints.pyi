from typing import (
    Any, Callable, Generic, NoReturn, Optional, overload, SupportsIndex,
    Type, TypeVar, Union as _UnionT,
)


def AnnoField() -> Any:
    """**THIS IS `TYPE_CHECKING` ONLY SYMBOL.

    This is workaround for class field type annotations for old
    python versions.

    Examples:
        # instead of class field annotation, like below

        class Foo:
            # spam: int  # <- not available in old versions.
            if TYPE_CHECKING:
                spam = AnnoField()  # type: int  # <- available in old versions.
    """
    ...
