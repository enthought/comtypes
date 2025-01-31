"""comtypes.client - High level client level COM support package."""

import ctypes
import logging

from comtypes import automation
from comtypes.client import dynamic  # noqa
from comtypes.client._code_cache import _find_gen_dir
from comtypes.client._constants import Constants  # noqa
from comtypes.client._events import GetEvents, PumpEvents, ShowEvents
from comtypes.client._generate import GetModule
from comtypes.client.misc import GetBestInterface, _manage, wrap_outparam  # noqa
from comtypes.hresult import *  # noqa

gen_dir = _find_gen_dir()
import comtypes.gen  # noqa

### for testing
##gen_dir = None

logger = logging.getLogger(__name__)


# backwards compatibility:
wrap = GetBestInterface

# Should we do this for POINTER(IUnknown) also?
ctypes.POINTER(automation.IDispatch).__ctypes_from_outparam__ = wrap_outparam  # type: ignore

from comtypes.client.misc import (
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
