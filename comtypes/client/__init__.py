"""comtypes.client - High level client level COM support package."""

################################################################
#
# TODO:
#
# - refactor some code into modules
#
################################################################

import ctypes
import logging
import os
import sys
from typing import TYPE_CHECKING, Any, Optional, Type, TypeVar, overload
from typing import Union as _UnionT

import comtypes
import comtypes.client.dynamic
from comtypes import GUID, CoClass, IUnknown, automation, typeinfo
from comtypes.client._code_cache import _find_gen_dir
from comtypes.client._constants import Constants
from comtypes.client._events import GetEvents, PumpEvents, ShowEvents
from comtypes.client._generate import GetModule
from comtypes.client._misc import GetBestInterface, _manage, wrap_outparam  # noqa
from comtypes.hresult import *  # noqa

if TYPE_CHECKING:
    from comtypes import hints  # type: ignore

gen_dir = _find_gen_dir()
import comtypes.gen

### for testing
##gen_dir = None

_T_IUnknown = TypeVar("_T_IUnknown", bound=IUnknown)
logger = logging.getLogger(__name__)


# backwards compatibility:
wrap = GetBestInterface

# Should we do this for POINTER(IUnknown) also?
ctypes.POINTER(automation.IDispatch).__ctypes_from_outparam__ = wrap_outparam  # type: ignore

from comtypes.client._misc import (
    CoGetObject,
    CreateObject,
    GetActiveObject,
    GetClassObject,
)

# fmt: off
__all__ = [
    "CreateObject", "GetActiveObject", "CoGetObject", "GetEvents",
    "ShowEvents", "PumpEvents", "GetModule", "GetClassObject",
]
# fmt: on
