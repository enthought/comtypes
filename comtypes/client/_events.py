import ctypes
import comtypes
from comtypes.hresult import *
import comtypes.automation
import comtypes.typeinfo
import comtypes.connectionpoints
import logging
logger = logging.getLogger(__name__)

class _AdviseConnection(object):
    def __init__(self, source, interface, receiver):
        cpc = source.QueryInterface(comtypes.connectionpoints.IConnectionPointContainer)
        self.cp = cpc.FindConnectionPoint(ctypes.byref(interface._iid_))
        logger.debug("Start advise %s", interface)
        self.cookie = self.cp.Advise(receiver)
        self.receiver = receiver

    def disconnect(self):
        if self.cookie:
            self.cp.Unadvise(self.cookie)
            logger.debug("Unadvised %s", self.cp)
            self.cp = None
            self.cookie = None
            del self.receiver

    def __del__(self):
        try:
            if self.cookie is not None:
                self.cp.Unadvise(self.cookie)
        except (comtypes.COMError, WindowsError):
            # Are we sure we want to ignore errors here?
            pass

def FindOutgoingInterface(source):
    """XXX Describe the strategy that is used..."""
    # If the COM object implements IProvideClassInfo2, it is easy to
    # find the default autgoing interface.
    try:
        pci = source.QueryInterface(comtypes.typeinfo.IProvideClassInfo2)
        guid = pci.GetGUID(1)
    except comtypes.COMError:
        pass
    else:
        # another try: block needed?
        try:
            interface = comtypes.com_interface_registry[str(guid)]
        except KeyError:
            tinfo = pci.GetClassInfo()
            tlib, index = tinfo.GetContainingTypeLib()
            from comtypes.client import GetModule
            GetModule(tlib)
            interface = comtypes.com_interface_registry[str(guid)]
        logger.debug("%s using sinkinterface %s", source, interface)
        return interface

    # If we can find the CLSID of the COM object, we can look for a
    # registered outgoing interface (__clsid has been set by
    # comtypes.client):
    clsid = source.__dict__.get('__clsid')
    try:
        interface = comtypes.com_coclass_registry[clsid]._outgoing_interfaces_[0]
    except KeyError:
        pass
    else:
        logger.debug("%s using sinkinterface from clsid %s", source, interface)
        return interface

##    interface = find_single_connection_interface(source)
##    if interface:
##        return interface

    raise TypeError("cannot determine source interface")

def find_single_connection_interface(source):
    # Enumerate the connection interfaces.  If we find a single one,
    # return it, if there are more, we give up since we cannot
    # determine which one to use.
    cpc = source.QueryInterface(comtypes.connectionpoints.IConnectionPointContainer)
    enum = cpc.EnumConnectionPoints()
    iid = enum.next().GetConnectionInterface()
    try:
        enum.next()
    except StopIteration:
        try:
            interface = comtypes.com_interface_registry[str(iid)]
        except KeyError:
            return None
        else:
            logger.debug("%s using sinkinterface from iid %s", source, interface)
            return interface
    else:
        logger.debug("%s has nore than one connection point", source)

    return None

from comtypes._comobject import _MethodFinder
class _SinkMethodFinder(_MethodFinder):
    def __init__(self, inst, sink):
        super(_SinkMethodFinder, self).__init__(inst)
        self.sink = sink

    def find_method(self, fq_name, mthname):
        try:
            return super(_SinkMethodFinder, self).find_method(fq_name, mthname)
        except AttributeError:
            try:
                return getattr(self.sink, fq_name)
            except AttributeError:
                return getattr(self.sink, mthname)

def CreateEventReceiver(interface, sink):

    class Sink(comtypes.COMObject):
        _com_interfaces_ = [interface]

        def _get_method_finder_(self, itf):
            # Use a special MethodFinder that will first try 'self',
            # then the sink.
            return _SinkMethodFinder(self, sink)

    return Sink()

def GetEvents(source, sink, interface=None):
    """Receive COM events from 'source'.  Events will call methods on
    the 'sink' object.  'interface' is the source interface to use.
    """
    # When called from CreateObject, the sourceinterface has already
    # been determined by the coclass.  Otherwise, the only thing that
    # makes sense is to use IProvideClassInfo2 to get the default
    # source interface.
    if interface is None:
        interface = FindOutgoingInterface(source)

    rcv = CreateEventReceiver(interface, sink)
    return _AdviseConnection(source, interface, rcv)

class EventDumper(object):
    """Universal sink for COM events."""

    def __getattr__(self, name):
        "Create event handler methods on demand"
        if name.startswith("__") and name.endswith("__"):
            raise AttributeError(name)
        print "# event found:", name
        def handler(self, this, *args, **kw):
            # XXX handler is called with 'this'.  Should we really print "None" instead?
            args = (None,) + args
            print "Event %s(%s)" % (name, ", ".join([repr(a) for a in args]))
        import new
        return new.instancemethod(handler, EventDumper, self)

def ShowEvents(source, interface=None):
    """Receive COM events from 'source'.  A special event sink will be
    used that first prints the names of events that are found in the
    outgoing interface, and will also print out the events when they
    are fired.
    """
    return comtypes.client.GetEvents(source, sink=EventDumper(), interface=interface)

def PumpEvents(timeout):
    """This following code waits for 'timeout' seconds in the way
    required for COM, internally doing the correct things depending
    on the COM appartment of the current thread.  It is possible to
    terminate the message loop by pressing CTRL+C, which will raise
    a KeyboardInterrupt.
    """
    # XXX Should there be a way to pass additional event handles which
    # can terminate this function?
    hevt = ctypes.windll.kernel32.CreateEventA(None, True, False, None)
    handles = (ctypes.c_void_p * 1)(hevt)
    RPC_S_CALLPENDING = -2147417835

    @ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_uint)
    def HandlerRoutine(dwCtrlType):
        if dwCtrlType == 0: # CTRL+C
            ctypes.windll.kernel32.SetEvent(hevt)
            return 1
        return 0

    ctypes.windll.kernel32.SetConsoleCtrlHandler(HandlerRoutine, 1)

    try:
        try:
            res = ctypes.oledll.ole32.CoWaitForMultipleHandles(0,
                                                               int(timeout * 1000),
                                                               len(handles), handles,
                                                               ctypes.byref(ctypes.c_ulong()))
        except WindowsError, details:
            if details[0] != RPC_S_CALLPENDING: # timeout expired
                raise
        else:
            raise KeyboardInterrupt
    finally:
        ctypes.windll.kernel32.CloseHandle(hevt)
        ctypes.windll.kernel32.SetConsoleCtrlHandler(HandlerRoutine, 0)

