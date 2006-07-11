import ctypes
import comtypes.automation
import comtypes.typeinfo
import comtypes.client

def Dispatch(obj):
    # Wrap an object in a Dispatch instance, exposing methods and properties
    # via fully dynamic dispatch
    if isinstance(obj, _Dispatch):
        return obj
    if isinstance(obj, ctypes.POINTER(comtypes.automation.IDispatch)):
        return _Dispatch(obj)
    return obj

class _Dispatch(object):
    # Expose methods and properties via fully dynamic dispatch
    def __init__(self, comobj):
        self._comobj = comobj

    def __enum(self):
        e = self._comobj.Invoke(-4) # DISPID_NEWENUM
        return e.QueryInterface(comtypes.automation.IEnumVARIANT)

    def __getitem__(self, index):
        enum = self.__enum()
        if index > 0:
            if 0 != enum.Skip(index):
                raise IndexError, "index out of range"
        item, fetched = enum.Next(1)
        if not fetched:
            raise IndexError, "index out of range"
        return item

    def QueryInterface(self, *args):
        "QueryInterface is forwarded to the real com object."
        return self._comobj.QueryInterface(*args)

    def __getattr__(self, name):
##        tc = self._comobj.GetTypeInfo(0).QueryInterface(comtypes.typeinfo.ITypeComp)
##        dispid = tc.Bind(name)[1].memid
        dispid = self._comobj.GetIDsOfNames(name)[0]
        flags = comtypes.automation.DISPATCH_PROPERTYGET
        return self._comobj.Invoke(dispid,
                                   _invkind=flags)

    def __iter__(self):
        return _Collection(self.__enum())

##    def __setitem__(self, index, value):
##        self._comobj.Invoke(-3, index, value,
##                            _invkind=comtypes.automation.DISPATCH_PROPERTYPUT|comtypes.automation.DISPATCH_PROPERTYPUTREF)

class _Collection(object):
    def __init__(self, enum):
        self.enum = enum

    def next(self):
        item, fetched = self.enum.Next(1)
        if fetched:
            return item
        raise StopIteration

    def __iter__(self):
        return self

__all__ = ["Dispatch"]
