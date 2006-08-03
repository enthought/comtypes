import sys, os
import logging
logging.basicConfig()
##logging.basicConfig(level=logging.DEBUG)
##logger = logging.getLogger(__name__)

import ctypes
import comtypes
from comtypes.hresult import *
import comtypes.client
import comtypes.errorinfo
import comtypes.server
import comtypes.server.automation
import comtypes.typeinfo

################################################################

# pathname of the type library file

tlbfile = os.path.join(os.path.dirname(__file__), "TestComServer.tlb")

# Create the wrapper in the comtypes.gen package, it will be named
# TestComServerLib; the name is derived from the 'library ' statement
# in the IDL file
if not hasattr(sys, "frozen"):
    # if running as frozen app (dll or exe), the wrapper should be in
    # the library archive, so we don't need to generate it.
    comtypes.client.GetModule(tlbfile)

# Import the wrapper
from comtypes.gen import TestComServerLib

################################################################

# Implement the CoClass.  Use the coclass from the wrapper as base
# class, and use DualDispMixin as base class which provides default
# implementations of IDispatch, IProvideClassInfo, IProvideClassInfo2
# interfaces.  ISupportErrorInfo is implemented by the COMObject base
# class.
class TestComServer(
    TestComServerLib.TestComServer, # the coclass from the typelib wrapper
    comtypes.server.automation.DualDispImplMixin, # other mixins
    ):

    # The default interface from the typelib MUST be the first
    # interface, other interfaces can follow

    _com_interfaces_ = TestComServerLib.TestComServer._com_interfaces_ + \
                       [comtypes.typeinfo.IProvideClassInfo2,
                        comtypes.errorinfo.ISupportErrorInfo,
                        ]

    # registry entries
    _reg_threading_ = "Both"
    _reg_progid_ = "TestComServerLib.TestComServer.1"
    _reg_novers_progid_ = "TestComServerLib.TestComServer"
    _reg_desc_ = "comtypes COM server sample for testing"
    _reg_clsctx_ = comtypes.CLSCTX_INPROC_SERVER | comtypes.CLSCTX_LOCAL_SERVER

    # needed for DualDispImplMixin:
    _typelib_path_ = tlbfile

    ################################
    # ITestComServer methods

    def ITestComServer__get_array(self, this, parray):
        # Hm, why is assignment to value needed?

        # these leak
##        parray[0].value = (1, "2", None, 3.14)
##        parray[0].value = (1, "2", None)

##        parray[0].value = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
##        parray[0].value = (1,), (4,)

##        parray[0].value = (), ()
##        parray[0].value = (), 0

        # these don't leak
##        parray[0].value = (1, 2, 3)
##        parray[0].value = (1, 2, None)
##        parray[0].value = (1, 3.14)
##        parray[0].value = [1, "(1, 2, 3)"]
##        parray[0].value = (1, "2")
##        parray[0].value = [1, "2"]
##        parray[0].value = (None, None, None)

        parray[0].value = (),

        return S_OK

    def ITestComServer__get_id(self, this, pid):
        pid[0] = id(self)
        return S_OK

    _name = u"spam, spam, spam"

    def ITestComServer__get_name(self, this, pname):
        pname[0] = u"spam, spam, spam"
        return S_OK

    def ITestComServer__set_name(self, this, name):
        self._name = name
        return S_OK

if __name__ == "__main__":
    try:
        from comtypes.server.register import UseCommandLine
##    logging.basicConfig(level=logging.DEBUG)
        UseCommandLine(TestComServer)
    except Exception:
        import traceback
        traceback.print_exc()
        raw_input()
