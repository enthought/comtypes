from ctypes import *
from comtypes import IUnknown, GUID, STDMETHOD, HRESULT

################################################################
# Interfaces
class IClassFactory(IUnknown):
    _iid_ = GUID("{00000001-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(HRESULT, "CreateInstance", [c_int, POINTER(GUID), POINTER(c_ulong)]),
        STDMETHOD(HRESULT, "LockServer", [c_int])]

class IExternalConnection(IUnknown):
    _iid_ = GUID("{00000019-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(HRESULT, "AddConnection", [c_ulong, c_ulong]),
        STDMETHOD(HRESULT, "ReleaseConnection", [c_ulong, c_ulong, c_ulong])]
