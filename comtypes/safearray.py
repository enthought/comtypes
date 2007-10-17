from ctypes import *
from comtypes import _safearray
from comtypes.partial import partial

_safearray_type_cache = {}

################################################################
# This is THE PUBLIC function: the gateway to the SAFEARRAY functionality.
def _midlSAFEARRAY(itemtype):
    """This function mimics the 'SAFEARRAY(aType)' IDL idiom.  It
    returns a subtype of SAFEARRAY, instances will be built with a
    typecode VT_...  corresponding to the aType, which must be one of
    the supported ctypes.
    """
    try:
        return POINTER(_safearray_type_cache[itemtype])
    except KeyError:
        sa_type = _make_safearray_type(itemtype)
        _safearray_type_cache[itemtype] = sa_type
        return POINTER(sa_type)

def _make_safearray_type(itemtype):
    # Create and return a subclass of tagSAFEARRAY
    from comtypes.automation import _ctype_to_vartype, VT_RECORD

    meta = type(_safearray.tagSAFEARRAY)
    sa_type = meta.__new__(meta,
                           "SAFEARRAY_%s" % itemtype.__name__,
                           (_safearray.tagSAFEARRAY,), {})

    try:
        vartype = _ctype_to_vartype[itemtype]
        extra = None
    except KeyError:
        if issubclass(itemtype, Structure):
            from comtypes.typeinfo import GetRecordInfoFromGuids
            extra = GetRecordInfoFromGuids(*itemtype._recordinfo_)
            vartype = VT_RECORD
        else:
            raise TypeError(itemtype)

    class _(partial, POINTER(sa_type)):
        # Should explain the ideas how SAFEARRAY is used in comtypes
        _itemtype_ = itemtype # a ctypes type
        _vartype_ = vartype # a VARTYPE value: VT_...
        _needsfree = False

        @classmethod
        def create(cls, value, extra=None):
            """Create a POINTER(SAFEARRAY_...) instance of the correct type;
            value is a sequence containing the items to store."""

            # XXX XXX
            #
            # For VT_UNKNOWN or VT_DISPATCH, extra must be a pointer to
            # the GUID of the interface.
            #
            # For VT_RECORD, extra must be a pointer to an IRecordInfo
            # describing the record.

            # XXX How to specify the lbound (3. parameter to CreateVectorEx)?
            # XXX How to test lbound != 0?
            pa = _safearray.SafeArrayCreateVectorEx(cls._vartype_,
                                                    0,
                                                    len(value),
                                                    extra)
            # We now have a POINTER(tagSAFEARRAY) instance which we must cast
            # to the correct type:
            pa = cast(pa, cls)
            # Now, fill the data in:
            ptr = POINTER(cls._itemtype_)() # container for the values
            _safearray.SafeArrayAccessData(pa, byref(ptr))
            try:
                for index, item in enumerate(value):
                    ptr[index] = item
            finally:
                _safearray.SafeArrayUnaccessData(pa)
            return pa

        @classmethod
        def from_param(cls, value):
            result = cls.create(value, extra)
            result._needsfree = True
            return result

        def __getitem__(self, index):
            # pparray[0] returns the whole array contents.
            if index != 0:
                raise IndexError("Only index 0 allowed")
            return self.unpack()

        def __setitem__(self, index, value):
            raise TypeError("Setting items not allowed")

        def __ctypes_from_outparam__(self):
            self._needsfree = True
            return self[0]
            
        def __del__(self):
            if self._needsfree:
                _safearray.SafeArrayDestroy(self)

        def _get_size(self, dim):
            "Return the number of elements for dimension 'dim'"
            return _safearray.SafeArrayGetUBound(self, dim)+1 - _safearray.SafeArrayGetLBound(self, dim)

        def unpack(self):
            """Unpack a POINTER(SAFEARRAY_...) into a Python tuple."""
            dim = _safearray.SafeArrayGetDim(self)

            if dim == 1:
                num_elements = self._get_size(1)
                return tuple(self._get_elements_raw(num_elements))
            elif dim == 2:
                # get the number of elements in each dimension
                rows, cols = self._get_size(1), self._get_size(2)
                # get all elements
                result = self._get_elements_raw(rows * cols)
                # transpose the result, because it is in VB order
                result = [tuple(result[r::rows]) for r in range(rows)]
                return tuple(result)
            else:
                lowerbounds = [_safearray.SafeArrayGetLBound(self, d) for d in range(1, dim+1)]
                indexes = (c_long * dim)(*lowerbounds)
                upperbounds = [_safearray.SafeArrayGetUBound(self, d) for d in range(1, dim+1)]
                return self._get_row(0, indexes, lowerbounds, upperbounds)

        def _get_elements_raw(self, num_elements):
            """Returns a flat list containing ALL elements in the safearray."""
            from comtypes.automation import VARIANT
            # XXX Not sure this is true:
            # For VT_UNKNOWN and VT_DISPATCH, we should retrieve the
            # interface iid by SafeArrayGetIID().
##            from comtypes.automation import VT_UNKNOWN, VT_DISPATCH
##            if self._vartype_ in (VT_UNKNOWN, VT_DISPATCH):
##                raise "HALT"
            ptr = POINTER(self._itemtype_)() # container for the values
            _safearray.SafeArrayAccessData(self, byref(ptr))
            try:
                if self._itemtype_ == VARIANT:
                    return [i.value for i in ptr[:num_elements]]
                else:
                    return ptr[:num_elements]
            finally:
                _safearray.SafeArrayUnaccessData(self)

        def _get_row(self, dim, indices, lowerbounds, upperbounds):
            # loop over the index of dimension 'dim'
            # we have to restore the index of the dimension we're looping over
            restore = indices[dim]

            result = []
            obj = self._itemtype_()
            pobj = byref(obj)
            if dim+1 == len(indices):
                # It should be faster to lock the array and get a whole row at once?
                # How to calculate the pointer offset?
                for i in range(indices[dim], upperbounds[dim]+1):
                    indices[dim] = i
                    _safearray.SafeArrayGetElement(self, indices, pobj)
                    result.append(obj.value)
            else:
                for i in range(indices[dim], upperbounds[dim]+1):
                    indices[dim] = i
                    result.append(self._get_row(dim+1, indices, lowerbounds, upperbounds))
            indices[dim] = restore
            return tuple(result) # for compatibility with pywin32.

    class _(partial, POINTER(POINTER(sa_type))):

        @classmethod
        def from_param(cls, value):
            if isinstance(value, cls._type_):
                return byref(value)
            return byref(cls._type_.create(value))

        def __setitem__(self, index, value):
            # create an LP_SAFEARRAY_... instance
            pa = self._type_.create(value)
            # XXX Must we destroy the currently contained data? 
            # fill it into self
            super(POINTER(POINTER(sa_type)), self).__setitem__(index, pa)

    return sa_type
