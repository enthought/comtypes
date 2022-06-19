################################################################
#
# Typelib constants
#
################################################################
import comtypes
import comtypes.automation


class Constants(object):
    """This class loads the type library from the supplied object,
    then exposes constants in the type library as attributes."""
    def __init__(self, obj):
        obj = obj.QueryInterface(comtypes.automation.IDispatch)
        tlib, index = obj.GetTypeInfo(0).GetContainingTypeLib()
        self.tcomp = tlib.GetTypeComp()

    def __getattr__(self, name):
        try:
            kind, desc = self.tcomp.Bind(name)
        except (WindowsError, comtypes.COMError):
            raise AttributeError(name)
        if kind != "variable":
            raise AttributeError(name)
        return desc._.lpvarValue[0].value

    def _bind_type(self, name):
        return self.tcomp.BindType(name)
