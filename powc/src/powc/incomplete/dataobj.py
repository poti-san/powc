"""COMデータオブジェクト"""

from ctypes import (
    POINTER,
    Structure,
    Union,
    _Pointer,
    byref,
    c_int32,
    c_uint16,
    c_uint32,
    c_void_p,
    c_wchar_p,
)
from enum import IntEnum, IntFlag
from typing import TYPE_CHECKING, Any, Iterator

from comtypes import GUID, STDMETHOD, IUnknown
from powc.core import ComResult, check_hresult, cr, query_interface
from powc.stream import IStream

from .. import _ole32

# from .__statdata import IEnumSTATDATA

ClipFormat = c_uint16
"""CLIPFORMAT"""


class DataDirection(IntEnum):
    """DATADIR"""

    GET = 1
    SET = 2


class MediumType(IntFlag):
    """TYMED"""

    HGLOBAL = 1
    FILE = 2
    ISTREAM = 4
    ISTORAGE = 8
    GDI = 16
    MFPICT = 32
    ENHMF = 64
    NULL = 0


class STGMEDIUM(Structure):
    class DummyUnion(Union):
        _fields_ = (
            ("bitmap_handle", c_void_p),
            ("metafilepict_handle", c_void_p),
            ("enhmetafile_handle", c_void_p),
            ("global_handle", c_void_p),
            ("file", c_wchar_p),
            ("stream", POINTER(IStream)),
            # TODO POINTER(IStorage)
            ("storage", POINTER(IUnknown)),
        )

    _anonymous_ = ("u",)
    _fields_ = (
        ("typed", c_uint32),
        ("u", DummyUnion),
        ("unk_for_release", POINTER(IUnknown)),
    )

    __slots__ = ()

    if TYPE_CHECKING:
        typed: int
        bitmap_handle: int
        metafilepict_handle: int
        enhmetafile_handle: int
        global_handle: int
        file: str
        stream: _Pointer[IStream]
        storage: _Pointer[IUnknown]  # TODO: IStorage
        unk_for_release: _Pointer[IUnknown]

    def __del__(self) -> None:
        _ReleaseStgMedium(self)


_ReleaseStgMedium = _ole32.ReleaseStgMedium
_ReleaseStgMedium.argtypes = (POINTER(STGMEDIUM),)
_ReleaseStgMedium.restype = c_int32


class DeviceTargetDevice(Structure):
    """DVTARGETDEVICE"""

    _fields_ = (
        ("size", c_uint32),
        ("driver_name_offset", c_uint16),
        ("device_name_offset", c_uint16),
        ("port_name_offset", c_uint16),
        ("ext_device_mode_offset", c_uint16),
        # ("data", c_byte * N)
    )

    __slots__ = ()


class FormatEtc(Structure):
    _fields_ = (
        ("format", ClipFormat),
        ("ptd", c_void_p),
        ("aspect", c_uint32),
        ("index", c_int32),
        ("_tymed", c_uint32),
    )

    __slots__ = ()

    if TYPE_CHECKING:
        format: ClipFormat
        ptd: int  # POINTER(DVTARGETDEVICE)
        aspect: int
        index: int

    @property
    def tymed(self) -> MediumType:
        return MediumType(self._tymed)

    @tymed.setter
    def tymed(self, x: MediumType) -> None:
        self._tymed = int(x)

    @staticmethod
    def create_simple(format: ClipFormat) -> "FormatEtc":
        fmtetc = FormatEtc()
        fmtetc.format = format
        fmtetc.index = -1
        fmtetc.tymed = MediumType.HGLOBAL
        return fmtetc


class IEnumFORMATETC(IUnknown):
    _iid_ = GUID("{00000103-0000-0000-C000-000000000046}")

    __slots__ = ()


IEnumFORMATETC._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(FormatEtc), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IEnumFORMATETC)),)),
]


class FormatEtcEnumerator:
    __o: Any  # POINTER(IEnumFORMATETC)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumFORMATETC)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __repr__(self) -> str:
        return "FormatEtcEnumerator"

    def __str__(self) -> str:
        return "FormatEtcEnumerator"

    def __iter__(self) -> Iterator[FormatEtc]:
        hr = 0
        fmtetc = FormatEtc()
        while (hr := self.__o.Next(1, byref(fmtetc), None)) == 0:
            yield fmtetc
            fmtetc = FormatEtc()
        check_hresult(hr)

    @property
    def items(self) -> tuple[FormatEtc, ...]:
        return tuple(iter(self))


class IDataObject(IUnknown):
    _iid_ = GUID("{0000010e-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "GetData", (POINTER(FormatEtc), POINTER(STGMEDIUM))),
        STDMETHOD(c_int32, "GetDataHere", (POINTER(FormatEtc), POINTER(STGMEDIUM))),
        STDMETHOD(c_int32, "QueryGetData", (POINTER(FormatEtc),)),
        STDMETHOD(c_int32, "GetCanonicalFormatEtc", (POINTER(FormatEtc), POINTER(FormatEtc))),
        STDMETHOD(c_int32, "SetData", (POINTER(FormatEtc), POINTER(STGMEDIUM), c_int32)),
        STDMETHOD(c_int32, "EnumFormatEtc", (c_uint32, POINTER(POINTER(IEnumFORMATETC)))),
        # TODO: IAdviseSink
        STDMETHOD(c_int32, "DAdvise", (POINTER(FormatEtc), c_uint32, POINTER(IUnknown), POINTER(c_uint32))),
        STDMETHOD(c_int32, "DUnadvise", (c_uint32,)),
        # TODO: STDMETHOD(c_int32, "EnumDAdvise", (POINTER(POINTER(IEnumSTATDATA)),)),
        STDMETHOD(c_int32, "EnumDAdvise", (POINTER(POINTER(IUnknown)),)),
    ]

    __slots__ = ()


class DataObject:
    __o: Any  # POINTER(IDataObject)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDataObject)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    # TODO: format: str
    def get_data_nothrow(self, format: ClipFormat) -> ComResult[STGMEDIUM]:
        fmtetc = FormatEtc.create_simple(format)
        stgmed = STGMEDIUM()
        return cr(self.__o.GetData(byref(fmtetc), byref(stgmed)), stgmed)

    def get_data(self, format: ClipFormat) -> STGMEDIUM:
        return self.get_data_nothrow(format).value

    # STDMETHOD(c_int32, "GetDataHere", (POINTER(FormatEtc), POINTER(STGMEDIUM))),
    # STDMETHOD(c_int32, "QueryGetData", (POINTER(FormatEtc),)),
    # STDMETHOD(c_int32, "GetCanonicalFormatEtc", (POINTER(FormatEtc), POINTER(FormatEtc))),
    # STDMETHOD(c_int32, "SetData", (POINTER(FormatEtc), POINTER(STGMEDIUM), c_int32)),
    # STDMETHOD(c_int32, "EnumFormatEtc", (POINTER(c_uint32), POINTER(POINTER(IEnumFORMATETC)))),

    def get_formatetc_enum_nothrow(self, direction: DataDirection) -> ComResult[FormatEtcEnumerator]:
        p = POINTER(IEnumFORMATETC)()
        return cr(self.__o.EnumFormatEtc(int(direction), byref(p)), FormatEtcEnumerator(p))

    def get_formatetc_enum(self, direction: DataDirection) -> FormatEtcEnumerator:
        return self.get_formatetc_enum_nothrow(direction).value

    def iter_formatetc_getter(self) -> Iterator[FormatEtc]:
        yield from self.get_formatetc_enum(DataDirection.GET)

    def iter_formatetc_setter(self) -> Iterator[FormatEtc]:
        yield from self.get_formatetc_enum(DataDirection.SET)

    # # TODO: IAdviseSink
    # STDMETHOD(c_int32, "DAdvise", (POINTER(FormatEtc), c_uint32, POINTER(IUnknown), POINTER(c_uint32))),
    # STDMETHOD(c_int32, "DUnadvise", (c_uint32,)),
    # STDMETHOD(c_int32, "EnumDAdvise", (POINTER(POINTER(IEnumSTATDATA)),)),

    @staticmethod
    def get_clipboard_nothrow() -> "ComResult[DataObject]":
        p = POINTER(IDataObject)()
        return cr(_OleGetClipboard(byref(p)), DataObject(p))

    @staticmethod
    def get_clipboard() -> "DataObject":
        return DataObject.get_clipboard_nothrow().value

    def set_clipboard_nothrow(self, flush: bool) -> ComResult[None]:
        ret = cr(_OleSetClipboard(self.__o), None)
        if not ret:
            return ret
        return cr(_OleFlushClipboard(), None) if flush else ret

    def set_clipboard(self, flush: bool) -> None:
        return self.set_clipboard_nothrow(flush).value


_OleGetClipboard = _ole32.OleGetClipboard
_OleGetClipboard.argtypes = (POINTER(POINTER(IDataObject)),)
_OleGetClipboard.restype = c_int32

_OleSetClipboard = _ole32.OleSetClipboard
_OleSetClipboard.argtypes = (POINTER(IDataObject),)
_OleSetClipboard.restype = c_int32

_OleFlushClipboard = _ole32.OleFlushClipboard
_OleFlushClipboard.argtypes = ()
_OleFlushClipboard.restype = c_int32
