from ctypes import *
import ctypes

from ._safearray import SAFEARRAY
from . import (
    IUnknown,
    helpstring,
    GUID,
    BSTR,
    COMMETHOD
)
from .automation import (
    DECIMAL,
    VARTYPE,
    CY,
    IDispatch,
    SCODE
)
from ctypes.wintypes import (
    WORD,
    DWORD,
    FILETIME,
    LARGE_INTEGER,
    ULARGE_INTEGER,
    VARIANT_BOOL,
    LPSTR,
    LPWSTR,
    LPOLESTR,
    BOOLEAN
)

LPSAFEARRAY = POINTER(SAFEARRAY)
CLSID = GUID
PROPID = c_ulong
REFFMTID = POINTER(GUID)
FMTID = GUID
REFCLSID = POINTER(GUID)
IID = GUID
OLECHAR = c_wchar
SNB = POINTER(LPOLESTR)


class tagPROPVARIANT(ctypes.Structure):
    pass


PROPVARIANT = tagPROPVARIANT


class tagPROPSPEC(ctypes.Structure):
    class DUMMYUNIONNAME(ctypes.Union):
        _fields_ = [
            ('propid', PROPID),
            ('lpwstr', LPOLESTR),
        ]

    _fields_ = [
        ('ulKind', c_ulong),
        ('DUMMYUNIONNAME', DUMMYUNIONNAME),
    ]
    _anonymous_ = ('DUMMYUNIONNAME',)


PROPSPEC = tagPROPSPEC


class tagSTATPROPSTG(ctypes.Structure):
    _fields_ = [
        ('lpwstrName', LPOLESTR),
        ('propid', PROPID),
        ('vt', VARTYPE),
    ]


STATPROPSTG = tagSTATPROPSTG


class tagSTATPROPSETSTG(ctypes.Structure):
    _fields_ = [
        ('fmtid', FMTID),
        ('clsid', CLSID),
        ('grfFlags', DWORD),
        ('mtime', FILETIME),
        ('ctime', FILETIME),
        ('atime', FILETIME),
        ('dwOSVersion', DWORD),
    ]


STATPROPSETSTG = tagSTATPROPSETSTG


class tagSERIALIZEDPROPERTYVALUE(ctypes.Structure):
    _fields_ = [
        ('dwType', DWORD),
        ('rgb', c_byte * 1),
    ]


SERIALIZEDPROPERTYVALUE = tagSERIALIZEDPROPERTYVALUE


class tagSTATSTG(ctypes.Structure):
    _fields_ = [
        ('pwcsName', LPOLESTR),
        ('type', DWORD),
        ('cbSize', ULARGE_INTEGER),
        ('mtime', FILETIME),
        ('ctime', FILETIME),
        ('atime', FILETIME),
        ('grfMode', DWORD),
        ('grfLocksSupported', DWORD),
        ('clsid', CLSID),
        ('grfStateBits', DWORD),
        ('reserved', DWORD),
    ]


STATSTG = tagSTATSTG


class tagBLOB(ctypes.Structure):
    _fields_ = [
        ('cbSize', c_ulong),
        ('pBlobData', POINTER(c_byte)),
    ]


BLOB = tagBLOB


class tagBSTRBLOB(ctypes.Structure):
    _fields_ = [
        ('cbSize', c_ulong),
        ('pData', POINTER(c_byte))
    ]


BSTRBLOB = tagBSTRBLOB


class tagCLIPDATA(ctypes.Structure):
    _fields_ = [
        ('cbSize', c_ulong),
        ('ulClipFmt', c_long),
        ('pClipData', POINTER(c_byte))
    ]


CLIPDATA = tagCLIPDATA


class tagCAC(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_char)),
    ]


CAC = tagCAC


class tagCAUB(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_ubyte)),
    ]


CAUB = tagCAUB


class tagCAI(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_short)),
    ]


CAI = tagCAI


class tagCAUI(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_ushort)),
    ]


CAUI = tagCAUI


class tagCAL(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_long)),
    ]


CAL = tagCAL


class tagCAUL(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_ulong)),
    ]


CAUL = tagCAUL


class tagCAFLT(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_float)),
    ]


CAFLT = tagCAFLT


class tagCADBL(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_double)),
    ]


CADBL = tagCADBL


class tagCACY(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(CY)),
    ]


CACY = tagCACY


class tagCADATE(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(c_double)),
    ]


CADATE = tagCADATE


class tagCABSTR(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(BSTR)),
    ]


CABSTR = tagCABSTR


class tagCABSTRBLOB(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(BSTRBLOB)),
    ]


CABSTRBLOB = tagCABSTRBLOB


class tagCABOOL(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(VARIANT_BOOL)),
    ]


CABOOL = tagCABOOL


class tagCASCODE(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(SCODE)),
    ]


CASCODE = tagCASCODE


class tagCAPROPVARIANT(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(PROPVARIANT)),
    ]


CAPROPVARIANT = tagCAPROPVARIANT


class tagCAH(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(LARGE_INTEGER)),
    ]


CAH = tagCAH


class tagCAUH(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(ULARGE_INTEGER)),
    ]


CAUH = tagCAUH


class tagCALPSTR(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(LPSTR)),
    ]


CALPSTR = tagCALPSTR


class tagCALPWSTR(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(LPWSTR)),
    ]


CALPWSTR = tagCALPWSTR


class tagCAFILETIME(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(FILETIME)),
    ]


CAFILETIME = tagCAFILETIME


class tagCACLIPDATA(ctypes.Structure):
    _fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(CLIPDATA)),
    ]


CACLIPDATA = tagCACLIPDATA


class tagCACLSID(ctypes.Structure):
    fields_ = [
        ('cElems', c_ulong),
        ('pElems', POINTER(CLSID)),
    ]


CACLSID = tagCACLSID

# =================  Interface Forward Declerations  =================
IID_IEnumSTATPROPSTG = GUID("{00000139-0000-0000-C000-000000000046}")


class IEnumSTATPROPSTG(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IEnumSTATPROPSTG
    _idlflags_ = []


IID_IEnumSTATPROPSETSTG = GUID("{0000013B-0000-0000-C000-000000000046}")


class IEnumSTATPROPSETSTG(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IEnumSTATPROPSETSTG
    _idlflags_ = []


IID_ISequentialStream = GUID("{0C733A30-2A1C-11CE-ADE5-00AA0044773D}")


class ISequentialStream(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_ISequentialStream
    _idlflags_ = []


IID_IStream = GUID("{0000000C-0000-0000-C000-000000000046}")


class IStream(ISequentialStream):
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IStream


IID_IStorage = GUID("{0000000B-0000-0000-C000-000000000046}")


class IStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IStorage
    _idlflags_ = []


IID_IEnumSTATSTG = GUID("{0000000D-0000-0000-C000-000000000046}")


class IEnumSTATSTG(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IEnumSTATSTG
    _idlflags_ = []


# ====================================================================

IID_IPropertyStorage = GUID("{00000138-0000-0000-C000-000000000046}")


class IPropertyStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertyStorage
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method ReadMultiple')],
            HRESULT,
            'ReadMultiple',
            (['in'], c_ulong, 'cpspec'),
            (['in'], PROPSPEC * 0, 'rgpspec'),
            (['out'], PROPVARIANT * 0, 'rgpropvar'),
        ),
        COMMETHOD(
            [helpstring('Method WriteMultiple')],
            HRESULT,
            'WriteMultiple',
            (['in'], c_ulong, 'cpspec'),
            (['in'], PROPSPEC * 0, 'rgpspec'),
            (['in'], PROPVARIANT * 0, 'rgpropvar'),
            (['in'], PROPID, 'propidNameFirst'),
        ),
        COMMETHOD(
            [helpstring('Method DeleteMultiple')],
            HRESULT,
            'DeleteMultiple',
            (['in'], c_ulong, 'cpspec'),
            (['in'], PROPSPEC * 0, 'rgpspec'),
        ),
        COMMETHOD(
            [helpstring('Method ReadPropertyNames')],
            HRESULT,
            'ReadPropertyNames',
            (['in'], c_ulong, 'cpropid'),
            (['in'], PROPID * 0, 'rgpropid'),
            (['out'], LPOLESTR * 0, 'rglpwstrName'),
        ),
        COMMETHOD(
            [helpstring('Method WritePropertyNames')],
            HRESULT,
            'WritePropertyNames',
            (['in'], c_ulong, 'cpropid'),
            (['in'], PROPID * 0, 'rgpropid'),
            (['in'], LPOLESTR * 0, 'rglpwstrName'),
        ),
        COMMETHOD(
            [helpstring('Method DeletePropertyNames')],
            HRESULT,
            'DeletePropertyNames',
            (['in'], c_ulong, 'cpropid'),
            (['in'], PROPID * 0, 'rgpropid'),
        ),
        COMMETHOD(
            [helpstring('Method Commit')],
            HRESULT,
            'Commit',
            (['in'], DWORD, 'grfCommitFlags'),
        ),
        COMMETHOD(
            [helpstring('Method Revert')],
            HRESULT,
            'Revert',
        ),
        COMMETHOD(
            [helpstring('Method Enum')],
            HRESULT,
            'Enum',
            (
                ['out'],
                POINTER(POINTER(IEnumSTATPROPSTG)),
                'ppenum'
            ),
        ),
        COMMETHOD(
            [helpstring('Method SetTimes')],
            HRESULT,
            'SetTimes',
            (['in'], POINTER(FILETIME), 'pctime'),
            (['in'], POINTER(FILETIME), 'patime'),
            (['in'], POINTER(FILETIME), 'pmtime'),
        ),
        COMMETHOD(
            [helpstring('Method SetClass')],
            HRESULT,
            'SetClass',
            (['in'], REFCLSID, 'clsid'),
        ),
        COMMETHOD(
            [helpstring('Method Stat')],
            HRESULT,
            'Stat',
            (['out'], POINTER(STATPROPSETSTG), 'pstatpsstg'),
        ),
    ]


IID_IPropertySetStorage = GUID("{0000013A-0000-0000-C000-000000000046}")


class IPropertySetStorage(IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IPropertySetStorage
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method Create')],
            HRESULT,
            'Create',
            (['in'], REFFMTID, 'rfmtid'),
            (['unique', 'in'], POINTER(CLSID), 'pclsid'),
            (['in'], DWORD, 'grfFlags'),
            (['in'], DWORD, 'grfMode'),
            (
                ['out'],
                POINTER(POINTER(IPropertyStorage)),
                'ppprstg'
            ),
        ),
        COMMETHOD(
            [helpstring('Method Open')],
            HRESULT,
            'Open',
            (['in'], REFFMTID, 'rfmtid'),
            (['in'], DWORD, 'grfMode'),
            (
                ['out'],
                POINTER(POINTER(IPropertyStorage)),
                'ppprstg'
            ),
        ),
        COMMETHOD(
            [helpstring('Method Delete')],
            HRESULT,
            'Delete',
            (['in'], REFFMTID, 'rfmtid'),
        ),
        COMMETHOD(
            [helpstring('Method Enum')],
            HRESULT,
            'Enum',
            (
                ['out'],
                POINTER(POINTER(IEnumSTATPROPSETSTG)),
                'ppenum'
            ),
        ),
    ]


class tagVersionedStream(ctypes.Structure):
    _fields_ = [
        ('guidVersion', GUID),
        ('pStream', POINTER(IStream)),
    ]


VERSIONEDSTREAM = tagVersionedStream
LPVERSIONEDSTREAM = POINTER(tagVersionedStream)

# Flags for IPropertySetStorage::Create
PROPSETFLAG_DEFAULT = 0
PROPSETFLAG_NONSIMPLE = 1
PROPSETFLAG_ANSI = 2

# (This flag is only supported on StgCreatePropStg & StgOpenPropStg
PROPSETFLAG_UNBUFFERED = 4
# (This flag causes a version-1 property set to be created
PROPSETFLAG_CASE_SENSITIVE = 8
# Flags for the reserved PID_BEHAVIOR property
PROPSET_BEHAVIOR_CASE_SENSITIVE = 1

PROPVAR_PAD1 = WORD
PROPVAR_PAD2 = WORD
PROPVAR_PAD3 = WORD


# =======================  PROPVARIANT  =======================

class _Union_1(ctypes.Union):
    class tag_inner_PROPVARIANT(ctypes.Structure):
        class _Union_2(ctypes.Union):
            _fields_ = [
                ('cVal', c_char),
                ('bVal', c_ubyte),
                ('iVal', c_short),
                ('uiVal', c_ushort),
                ('lVal', c_long),
                ('ulVal', c_ulong),
                ('intVal', c_int),
                ('uintVal', c_uint),
                ('hVal', LARGE_INTEGER),
                ('uhVal', ULARGE_INTEGER),
                ('fltVal', c_float),
                ('dblVal', c_double),
                ('boolVal', VARIANT_BOOL),
                ('bool', VARIANT_BOOL),
                ('scode', SCODE),
                ('cyVal', CY),
                ('date', c_double),
                ('filetime', FILETIME),
                ('puuid', POINTER(CLSID)),
                ('pclipdata', POINTER(CLIPDATA)),
                ('bstrVal', BSTR),
                ('bstrblobVal', BSTRBLOB),
                ('blob', BLOB),
                ('pszVal', LPSTR),
                ('pwszVal', LPWSTR),
                ('punkVal', POINTER(IUnknown)),
                ('pdispVal', POINTER(IDispatch)),
                ('pStream', POINTER(IStream)),
                ('pStorage', POINTER(IStorage)),
                ('pVersionedStream', LPVERSIONEDSTREAM),
                ('parray', LPSAFEARRAY),
                ('cac', CAC),
                ('caub', CAUB),
                ('cai', CAI),
                ('caui', CAUI),
                ('cal', CAL),
                ('caul', CAUL),
                ('cah', CAH),
                ('cauh', CAUH),
                ('caflt', CAFLT),
                ('cadbl', CADBL),
                ('cabool', CABOOL),
                ('cascode', CASCODE),
                ('cacy', CACY),
                ('cadate', CADATE),
                ('cafiletime', CAFILETIME),
                ('cauuid', CACLSID),
                ('caclipdata', CACLIPDATA),
                ('cabstr', CABSTR),
                ('cabstrblob', CABSTRBLOB),
                ('calpstr', CALPSTR),
                ('calpwstr', CALPWSTR),
                ('capropvar', CAPROPVARIANT),
                ('pcVal', POINTER(c_char)),
                ('pbVal', POINTER(c_ubyte)),
                ('piVal', POINTER(c_short)),
                ('puiVal', POINTER(c_ushort)),
                ('plVal', POINTER(c_long)),
                ('pulVal', POINTER(c_ulong)),
                ('pintVal', POINTER(c_int)),
                ('puintVal', POINTER(c_uint)),
                ('pfltVal', POINTER(c_float)),
                ('pdblVal', POINTER(c_double)),
                ('pboolVal', POINTER(VARIANT_BOOL)),
                ('pdecVal', POINTER(DECIMAL)),
                ('pscode', POINTER(SCODE)),
                ('pcyVal', POINTER(CY)),
                ('pdate', POINTER(c_double)),
                ('pbstrVal', POINTER(BSTR)),
                ('ppunkVal', POINTER(POINTER(IUnknown))),
                ('ppdispVal', POINTER(POINTER(IDispatch))),
                ('pparray', POINTER(LPSAFEARRAY)),
                ('pvarVal', POINTER(PROPVARIANT)),
            ]

        _fields_ = [
            ('vt', VARTYPE),
            ('wReserved1', PROPVAR_PAD1),
            ('wReserved2', PROPVAR_PAD2),
            ('wReserved3', PROPVAR_PAD3),
            ('_Union_2', _Union_2),
        ]

        _anonymous_ = ('_Union_2',)

    _fields_ = [
        ('tag_inner_PROPVARIANT', tag_inner_PROPVARIANT),
        ('decVal', DECIMAL),
    ]

    _anonymous_ = ('tag_inner_PROPVARIANT',)


tagPROPVARIANT._fields_ = [
    ('_Union_1', _Union_1),
]

tagPROPVARIANT._anonymous_ = ('_Union_1',)

# =============================================================

# ===================  Interface Methods  =====================

IEnumSTATSTG._methods_ = [
    COMMETHOD(
        [helpstring('Method Next'), 'local'],
        HRESULT,
        'Next',
        (['in'], c_ulong, 'celt'),
        (['out'], POINTER(STATSTG), 'rgelt'),
        (['out'], POINTER(c_ulong), 'pceltFetched'),
    ),
    COMMETHOD(
        [helpstring('Method Skip')],
        HRESULT,
        'Skip',
        (['in'], c_ulong, 'celt'),
    ),
    COMMETHOD(
        [helpstring('Method Reset')],
        HRESULT,
        'Reset',
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (['out'], POINTER(POINTER(IEnumSTATSTG)), 'ppenum'),
    ),
]

IStorage._methods_ = [
    COMMETHOD(
        [helpstring('Method CreateStream')],
        HRESULT,
        'CreateStream',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['in'], DWORD, 'grfMode'),
        (['in'], DWORD, 'reserved1'),
        (['in'], DWORD, 'reserved2'),
        (['out'], POINTER(POINTER(IStream)), 'ppstm'),
    ),
    COMMETHOD(
        [helpstring('Method OpenStream'), 'local'],
        HRESULT,
        'OpenStream',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['in'], POINTER(c_void_p), 'reserved1'),
        (['in'], DWORD, 'grfMode'),
        (['in'], DWORD, 'reserved2'),
        (['out'], POINTER(POINTER(IStream)), 'ppstm'),
    ),
    COMMETHOD(
        [helpstring('Method CreateStorage')],
        HRESULT,
        'CreateStorage',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['in'], DWORD, 'grfMode'),
        (['in'], DWORD, 'reserved1'),
        (['in'], DWORD, 'reserved2'),
        (['out'], POINTER(POINTER(IStorage)), 'ppstg'),
    ),
    COMMETHOD(
        [helpstring('Method OpenStorage')],
        HRESULT,
        'OpenStorage',
        (['unique', 'in'], POINTER(OLECHAR), 'pwcsName'),
        (['unique', 'in'], POINTER(IStorage), 'pstgPriority'),
        (['in'], DWORD, 'grfMode'),
        (['unique', 'in'], SNB, 'snbExclude'),
        (['in'], DWORD, 'reserved'),
        (['out'], POINTER(POINTER(IStorage)), 'ppstg'),
    ),
    COMMETHOD(
        [helpstring('Method CopyTo'), 'local'],
        HRESULT,
        'CopyTo',
        (['in'], DWORD, 'ciidExclude'),
        (['in'], POINTER(IID), 'rgiidExclude'),
        (['in'], SNB, 'snbExclude'),
        (['in'], POINTER(IStorage), 'pstgDest'),
    ),
    COMMETHOD(
        [helpstring('Method MoveElementTo')],
        HRESULT,
        'MoveElementTo',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
        (['unique', 'in'], POINTER(IStorage), 'pstgDest'),
        (['in'], POINTER(OLECHAR), 'pwcsNewName'),
        (['in'], DWORD, 'grfFlags'),
    ),
    COMMETHOD(
        [helpstring('Method Commit')],
        HRESULT,
        'Commit',
        (['in'], DWORD, 'grfCommitFlags'),
    ),
    COMMETHOD(
        [helpstring('Method Revert')],
        HRESULT,
        'Revert',
    ),
    COMMETHOD(
        [helpstring('Method EnumElements'), 'local'],
        HRESULT,
        'EnumElements',
        (['in'], DWORD, 'reserved1'),
        (['in'], POINTER(c_void_p), 'reserved2'),
        (['in'], DWORD, 'reserved3'),
        (['out'], POINTER(POINTER(IEnumSTATSTG)), 'ppenum'),
    ),
    COMMETHOD(
        [helpstring('Method DestroyElement')],
        HRESULT,
        'DestroyElement',
        (['in'], POINTER(OLECHAR), 'pwcsName'),
    ),
    COMMETHOD(
        [helpstring('Method RenameElement')],
        HRESULT,
        'RenameElement',
        (['in'], POINTER(OLECHAR), 'pwcsOldName'),
        (['in'], POINTER(OLECHAR), 'pwcsNewName'),
    ),
    COMMETHOD(
        [helpstring('Method SetElementTimes')],
        HRESULT,
        'SetElementTimes',
        (['unique', 'in'], POINTER(OLECHAR), 'pwcsName'),
        (['unique', 'in'], POINTER(FILETIME), 'pctime'),
        (['unique', 'in'], POINTER(FILETIME), 'patime'),
        (['unique', 'in'], POINTER(FILETIME), 'pmtime'),
    ),
    COMMETHOD(
        [helpstring('Method SetClass')],
        HRESULT,
        'SetClass',
        (['in'], REFCLSID, 'clsid'),
    ),
    COMMETHOD(
        [helpstring('Method SetStateBits')],
        HRESULT,
        'SetStateBits',
        (['in'], DWORD, 'grfStateBits'),
        (['in'], DWORD, 'grfMask'),
    ),
    COMMETHOD(
        [helpstring('Method Stat')],
        HRESULT,
        'Stat',
        (['out'], POINTER(STATSTG), 'pstatstg'),
        (['in'], DWORD, 'grfStatFlag'),
    ),
]

ISequentialStream._methods_ = [
    COMMETHOD(
        [helpstring('Method Read'), 'local'],
        HRESULT,
        'Read',
        (['out'], POINTER(c_void_p), 'pv'),
        (['in'], c_ulong, 'cb'),
        (['out'], POINTER(c_ulong), 'pcbRead'),
    ),
    COMMETHOD(
        [helpstring('Method Write'), 'local'],
        HRESULT,
        'Write',
        (['in'], POINTER(c_void_p), 'pv'),
        (['in'], c_ulong, 'cb'),
        (['out'], POINTER(c_ulong), 'pcbWritten'),
    ),
]

IStream._methods_ = [
    COMMETHOD(
        [helpstring('Method Seek'), 'local'],
        HRESULT,
        'Seek',
        (['in'], LARGE_INTEGER, 'dlibMove'),
        (['in'], DWORD, 'dwOrigin'),
        (
            ['out'],
            POINTER(ULARGE_INTEGER),
            'plibNewPosition'
        ),
    ),
    COMMETHOD(
        [helpstring('Method SetSize')],
        HRESULT,
        'SetSize',
        (['in'], ULARGE_INTEGER, 'libNewSize'),
    ),
    COMMETHOD(
        [helpstring('Method CopyTo'), 'local'],
        HRESULT,
        'CopyTo',
        (['in'], POINTER(IStream), 'pstm'),
        (['in'], ULARGE_INTEGER, 'cb'),
        (['out'], POINTER(ULARGE_INTEGER), 'pcbRead'),
        (['out'], POINTER(ULARGE_INTEGER), 'pcbWritten'),
    ),
    COMMETHOD(
        [helpstring('Method Commit')],
        HRESULT,
        'Commit',
        (['in'], DWORD, 'grfCommitFlags'),
    ),
    COMMETHOD(
        [helpstring('Method Revert')],
        HRESULT,
        'Revert',
    ),
    COMMETHOD(
        [helpstring('Method LockRegion')],
        HRESULT,
        'LockRegion',
        (['in'], ULARGE_INTEGER, 'libOffset'),
        (['in'], ULARGE_INTEGER, 'cb'),
        (['in'], DWORD, 'dwLockType'),
    ),
    COMMETHOD(
        [helpstring('Method UnlockRegion')],
        HRESULT,
        'UnlockRegion',
        (['in'], ULARGE_INTEGER, 'libOffset'),
        (['in'], ULARGE_INTEGER, 'cb'),
        (['in'], DWORD, 'dwLockType'),
    ),
    COMMETHOD(
        [helpstring('Method Stat')],
        HRESULT,
        'Stat',
        (['out'], POINTER(STATSTG), 'pstatstg'),
        (['in'], DWORD, 'grfStatFlag'),
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (['out'], POINTER(POINTER(IStream)), 'ppstm'),
    ),
]

IEnumSTATPROPSTG._methods_ = [
    COMMETHOD(
        [helpstring('Method Next'), 'local'],
        HRESULT,
        'Next',
        (['in'], c_ulong, 'celt'),
        (['out'], POINTER(STATPROPSTG), 'rgelt'),
        (['out'], POINTER(c_ulong), 'pceltFetched'),
    ),
    COMMETHOD(
        [helpstring('Method Skip')],
        HRESULT,
        'Skip',
        (['in'], c_ulong, 'celt'),
    ),
    COMMETHOD(
        [helpstring('Method Reset')],
        HRESULT,
        'Reset',
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (
            ['out'],
            POINTER(POINTER(IEnumSTATPROPSTG)),
            'ppenum'
        ),
    ),
]

IEnumSTATPROPSETSTG._methods_ = [
    COMMETHOD(
        [helpstring('Method Next'), 'local'],
        HRESULT,
        'Next',
        (['in'], c_ulong, 'celt'),
        (['out'], POINTER(STATPROPSETSTG), 'rgelt'),
        (['out'], POINTER(c_ulong), 'pceltFetched'),
    ),
    COMMETHOD(
        [helpstring('Method Skip')],
        HRESULT,
        'Skip',
        (['in'], c_ulong, 'celt'),
    ),
    COMMETHOD(
        [helpstring('Method Reset')],
        HRESULT,
        'Reset',
    ),
    COMMETHOD(
        [helpstring('Method Clone')],
        HRESULT,
        'Clone',
        (
            ['out'],
            POINTER(POINTER(IEnumSTATPROPSETSTG)),
            'ppenum'
        ),
    ),
]

# =============================================================

# Property IDs for the DiscardableInformation Property Set
PIDDI_THUMBNAIL = 0x00000002  # VT_BLOB

# Property IDs for the SummaryInformation Property Set
PIDSI_TITLE = 0x00000002  # VT_LPSTR
PIDSI_SUBJECT = 0x00000003  # VT_LPSTR
PIDSI_AUTHOR = 0x00000004  # VT_LPSTR
PIDSI_KEYWORDS = 0x00000005  # VT_LPSTR
PIDSI_COMMENTS = 0x00000006  # VT_LPSTR
PIDSI_TEMPLATE = 0x00000007  # VT_LPSTR
PIDSI_LASTAUTHOR = 0x00000008  # VT_LPSTR
PIDSI_REVNUMBER = 0x00000009  # VT_LPSTR
PIDSI_EDITTIME = 0x0000000A  # VT_FILETIME (UTC)
PIDSI_LASTPRINTED = 0x0000000B  # VT_FILETIME (UTC)
PIDSI_CREATE_DTM = 0x0000000C  # VT_FILETIME (UTC)
PIDSI_LASTSAVE_DTM = 0x0000000D  # VT_FILETIME (UTC)
PIDSI_PAGECOUNT = 0x0000000E  # VT_I4
PIDSI_WORDCOUNT = 0x0000000F  # VT_I4
PIDSI_c_charCOUNT = 0x00000010  # VT_I4
PIDSI_THUMBNAIL = 0x00000011  # VT_CF
PIDSI_APPNAME = 0x00000012  # VT_LPSTR
PIDSI_DOC_SECURITY = 0x00000013  # VT_I4

# Property IDs for the DocSummaryInformation Property Set
PIDDSI_CATEGORY = 0x00000002  # VT_LPSTR
PIDDSI_PRESFORMAT = 0x00000003  # VT_LPSTR
PIDDSI_BYTECOUNT = 0x00000004  # VT_I4
PIDDSI_LINECOUNT = 0x00000005  # VT_I4
PIDDSI_PARCOUNT = 0x00000006  # VT_I4
PIDDSI_SLIDECOUNT = 0x00000007  # VT_I4
PIDDSI_NOTECOUNT = 0x00000008  # VT_I4
PIDDSI_HIDDENCOUNT = 0x00000009  # VT_I4
PIDDSI_MMCLIPCOUNT = 0x0000000A  # VT_I4
PIDDSI_SCALE = 0x0000000B  # VT_BOOL
PIDDSI_HEADINGPAIR = 0x0000000C  # VT_VARIANT | VT_VECTOR
PIDDSI_DOCPARTS = 0x0000000D  # VT_LPSTR | VT_VECTOR
PIDDSI_MANAGER = 0x0000000E  # VT_LPSTR
PIDDSI_COMPANY = 0x0000000F  # VT_LPSTR
PIDDSI_LINKSDIRTY = 0x00000010  # VT_BOOL

# FMTID_MediaFileSummaryInfo - Property IDs
PIDMSI_EDITOR = 0x00000002  # VT_LPWSTR
PIDMSI_SUPPLIER = 0x00000003  # VT_LPWSTR
PIDMSI_SOURCE = 0x00000004  # VT_LPWSTR
PIDMSI_SEQUENCE_NO = 0x00000005  # VT_LPWSTR
PIDMSI_PROJECT = 0x00000006  # VT_LPWSTR
PIDMSI_STATUS = 0x00000007  # VT_UI4
PIDMSI_OWNER = 0x00000008  # VT_LPWSTR
PIDMSI_RATING = 0x00000009  # VT_LPWSTR
PIDMSI_PRODUCTION = 0x0000000A  # VT_FILETIME (UTC)
PIDMSI_COPYRIGHT = 0x0000000B  # VT_LPWSTR

# =======================  FUNCTIONS  =========================


ole32 = ctypes.windll.OLE32

# _Check_return_ WINOLEAPI PropVariantCopy(
#     _Out_ PROPVARIANT* pvarDest,
#     _In_ PROPVARIANT * pvarSrc
# );
PropVariantCopy = ole32.PropVariantCopy
PropVariantCopy.restype = HRESULT

# WINOLEAPI PropVariantClear(_Inout_ PROPVARIANT* pvar);
PropVariantClear = ole32.PropVariantClear
PropVariantClear.restype = HRESULT

# WINOLEAPI FreePropVariantArray(
#     _In_ c_ulong cVariants,
#     _Inout_updates_(cVariants) PROPVARIANT* rgvars
# );
FreePropVariantArray = ole32.FreePropVariantArray
FreePropVariantArray.restype = HRESULT


def PropVariantInit(pvar):
    return ctypes.memset(pvar, 0, ctypes.sizeof(PROPVARIANT))


# EXTERN_C
# _Success_(TRUE) // Raises status on failure
# SERIALIZEDPROPERTYVALUE* __stdcall
# StgConvertVariantToProperty(
#     _In_ PROPVARIANT* pvar,
#     _In_ c_ushort CodePage,
#     _Out_writes_bytes_opt_(*pcb) SERIALIZEDPROPERTYVALUE* pprop,
#     _Inout_ c_ulong* pcb,
#     _In_ PROPID pid,
#     _Reserved_ BOOLEAN fReserved,
#     _Inout_opt_ c_ulong* pcIndirect
# );
StgConvertVariantToProperty = ole32.StgConvertVariantToProperty
StgConvertVariantToProperty.restype = SERIALIZEDPROPERTYVALUE

# EXTERN_C
# _Success_(TRUE) // Raises status on failure
# BOOLEAN __stdcall
# StgConvertPropertyToVariant(
#     _In_ SERIALIZEDPROPERTYVALUE* pprop,
#     _In_ c_ushort CodePage,
#     _Out_ PROPVARIANT* pvar,
#     _In_ PMemoryAllocator* pma
# );
StgConvertPropertyToVariant = ole32.StgConvertPropertyToVariant
StgConvertPropertyToVariant.restype = BOOLEAN

# Additional Prototypes for ALL interfaces
oleaut32 = ctypes.windll.OLEAUT32

# c_ulong BSTR_UserSize(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in BSTR *
# );
BSTR_UserSize = oleaut32.BSTR_UserSize
BSTR_UserSize.restype = c_ulong

# c_ubyte * BSTR_UserMarshal(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in BSTR *
# );
BSTR_UserMarshal = oleaut32.BSTR_UserMarshal
BSTR_UserMarshal.restype = POINTER(c_ubyte)

# c_ubyte * BSTR_UserUnmarshal(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out BSTR *
# );
BSTR_UserUnmarshal = oleaut32.BSTR_UserUnmarshal
BSTR_UserUnmarshal.restype = POINTER(c_ubyte)

# VOID BSTR_UserFree(
#     __RPC__in c_ulong *,
#     __RPC__in BSTR *
# );
BSTR_UserFree = oleaut32.BSTR_UserFree
BSTR_UserFree.restype = c_void_p

# c_ulong LPSAFEARRAY_UserSize(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserSize = oleaut32.LPSAFEARRAY_UserSize
LPSAFEARRAY_UserSize.restype = c_ulong

# c_ubyte * LPSAFEARRAY_UserMarshal(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserMarshal = oleaut32.LPSAFEARRAY_UserMarshal
LPSAFEARRAY_UserMarshal.restype = POINTER(c_ubyte)

# c_ubyte * LPSAFEARRAY_UserUnmarshal(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out LPSAFEARRAY *
# );
LPSAFEARRAY_UserUnmarshal = oleaut32.LPSAFEARRAY_UserUnmarshal
LPSAFEARRAY_UserUnmarshal.restype = POINTER(c_ubyte)

# VOID LPSAFEARRAY_UserFree(
#     __RPC__in c_ulong *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserFree = oleaut32.LPSAFEARRAY_UserFree
LPSAFEARRAY_UserFree.restype = c_void_p

# c_ulong BSTR_UserSize64(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in BSTR *
# );
BSTR_UserSize64 = oleaut32.BSTR_UserSize64
BSTR_UserSize64.restype = c_ulong

# c_ubyte * BSTR_UserMarshal64(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in BSTR *
# );
BSTR_UserMarshal64 = oleaut32.BSTR_UserMarshal64
BSTR_UserMarshal64.restype = POINTER(c_ubyte)

# c_ubyte * BSTR_UserUnmarshal64(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out BSTR *
# );
BSTR_UserUnmarshal64 = oleaut32.BSTR_UserUnmarshal64
BSTR_UserUnmarshal64.restype = POINTER(c_ubyte)

# VOID BSTR_UserFree64(
#     __RPC__in c_ulong *,
#     __RPC__in BSTR *
# );
BSTR_UserFree64 = oleaut32.BSTR_UserFree64
BSTR_UserFree64.restype = c_void_p

# c_ulong LPSAFEARRAY_UserSize64(
#     __RPC__in c_ulong *,
#     c_ulong,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserSize64 = oleaut32.LPSAFEARRAY_UserSize64
LPSAFEARRAY_UserSize64.restype = c_ulong

# c_ubyte * LPSAFEARRAY_UserMarshal64(
#     __RPC__in c_ulong *,
#     __RPC__inout_xcount(0) c_ubyte *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserMarshal64 = oleaut32.LPSAFEARRAY_UserMarshal64
LPSAFEARRAY_UserMarshal64.restype = POINTER(c_ubyte)

# c_ubyte * LPSAFEARRAY_UserUnmarshal64(
#     __RPC__in c_ulong *,
#     __RPC__in_xcount(0) c_ubyte *,
#     __RPC__out LPSAFEARRAY *
# );
LPSAFEARRAY_UserUnmarshal64 = oleaut32.LPSAFEARRAY_UserUnmarshal64
LPSAFEARRAY_UserUnmarshal64.restype = POINTER(c_ubyte)

# VOID LPSAFEARRAY_UserFree64(
#     __RPC__in c_ulong *,
#     __RPC__in LPSAFEARRAY *
# );
LPSAFEARRAY_UserFree64 = oleaut32.LPSAFEARRAY_UserFree64
LPSAFEARRAY_UserFree64.restype = c_void_p

# =============================================================
