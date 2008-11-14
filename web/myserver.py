import comtypes
import comtypes.server.localserver

from comtypes.client import GetModule
GetModule("mytypelib.tlb")

from comtypes.gen.MyTypeLib import MyObject

class MyObjectImpl(MyObject):
    # registry entries
    _reg_threading_ = "Both"
    _reg_progid_ = "MyTypeLib.MyObject.1"
    _reg_novers_progid_ = "MyTypeLib.MyObject"
    _reg_desc_ = "Simple COM server for testing"
    _reg_clsctx_ = comtypes.CLSCTX_INPROC_SERVER | comtypes.CLSCTX_LOCAL_SERVER
    _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE

    def MyMethod(self, a, b):
        return a + b

if __name__ == "__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(MyObjectImpl)
