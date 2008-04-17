import logging

from ctypes import *
from comtypes.hresult import *

from comtypes import COMObject, IUnknown
from comtypes.automation import IEnumVARIANT

logger = logging.getLogger(__name__)

__all__ = ["VARIANTEnumerator"]

class VARIANTEnumerator(COMObject):
    _com_interfaces_ = [IEnumVARIANT]

    def __init__(self, itemtype, jobs):
        self.jobs = jobs # keep, so that we can restore our iterator (in Reset, and Clone).
        self.itemtype = itemtype
        self.item_interface = itemtype._com_interfaces_[0]
        self.seq = iter(self.jobs)
        super(VARIANTEnumerator, self).__init__()

    def Next(self, this, celt, rgVar, pCeltFetched):
        if not rgVar: return E_POINTER
        if not pCeltFetched: pCeltFetched = [None]
        pCeltFetched[0] = 0
        try:
            for index in range(celt):
                job = self.itemtype(self.seq.next())
                p = POINTER(self.item_interface)()
                job.IUnknown_QueryInterface(None,
                                            pointer(p._iid_),
                                            byref(p))
                rgVar[index].value = p
                pCeltFetched[0] += 1
        except StopIteration:
            pass
        if pCeltFetched[0] == celt:
            return S_OK
        return S_FALSE

    def Skip(self, this, celt):
        # skip some elements.
        try:
            for _ in range(celt):
                self.seq.next()
        except StopIteration:
            return S_FALSE
        return S_OK

    def Reset(self, this):
        self.seq = iter(self.jobs)
        return S_OK

    # Clone

################################################################

class COMCollection(COMObject):
    """Abstract base class which implements Count, Item, and _NewEnum."""
    def __init__(self, itemtype, collection):
        self.collection = collection
        self.itemtype = itemtype
        super(COMCollection, self).__init__()

    def _get_Item(self, this, pathname, pitem):
        if not pitem:
            return E_POINTER
        item = self.itemtype(pathname)
        return item.IUnknown_QueryInterface(None,
                                            pointer(pitem[0]._iid_),
                                            pitem)

    def _get_Count(self, this, pcount):
        if not pcount:
            return E_POINTER
        pcount[0] = len(self.collection)
        return S_OK

    def _get__NewEnum(self, this, penum):
        if not penum:
            return E_POINTER
        enum = VARIANTEnumerator(self.itemtype, self.collection)
        return enum.IUnknown_QueryInterface(None,
                                            pointer(IUnknown._iid_),
                                            penum)

