from ctypes import _SimpleCData, windll


class BSTR(_SimpleCData):
    "The windows BSTR data type"
    _type_ = "X"
    _needsfree = False

    def __repr__(self):
        return "%s(%r)" % (self.__class__.__name__, self.value)

    def __ctypes_from_outparam__(self):
        self._needsfree = True
        return self.value

    def __del__(self, _free=windll.oleaut32.SysFreeString):
        # Free the string if self owns the memory
        # or if instructed by __ctypes_from_outparam__.
        if self._b_base_ is None or self._needsfree:
            _free(self)

    @classmethod
    def from_param(cls, value):
        """Convert into a foreign function call parameter."""
        if isinstance(value, cls):
            return value
        # Although the builtin SimpleCData.from_param call does the
        # right thing, it doesn't ensure that SysFreeString is called
        # on destruction.
        return cls(value)
