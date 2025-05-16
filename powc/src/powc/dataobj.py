"""COMデータオブジェクト機能を提供します。主なクラスは :class:`DataObject` です。"""

from ctypes import (
    POINTER,
    Structure,
    Union,
    _Pointer,
    byref,
    c_byte,
    c_int16,
    c_int32,
    c_uint16,
    c_uint32,
    c_void_p,
    c_wchar,
    c_wchar_p,
)
from enum import IntEnum, IntFlag
from typing import TYPE_CHECKING, Any, Iterator

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, check_hresult, cr, query_interface
from powc.globalmem import globalmem_lock
from powc.stream import ComStream, IStream

from . import _ole32, _user32

# from .__statdata import IEnumSTATDATA


class ClipboardFormat:
    __fmt: int

    def __init__(self, fmt: int) -> None:
        self.__fmt = fmt

    @property
    def value(self) -> int:
        return self.__fmt

    def __int__(self) -> int:
        return self.__fmt

    def __str__(self) -> str:
        match self.value:
            case 1:
                return "CF_TEXT"
            case 2:
                return "CF_BITMAP"
            case 3:
                return "CF_METAFILEPICT"
            case 4:
                return "CF_SYLK"
            case 5:
                return "CF_DIB"
            case 6:
                return "CF_TIFF"
            case 7:
                return "CF_OEMTEXT"
            case 8:
                return "CF_DIB"
            case 9:
                return "CF_PALETTE"
            case 10:
                return "CF_PENDATA"
            case 11:
                return "CF_RIFF"
            case 12:
                return "CF_WAVE"
            case 13:
                return "CF_UNICODETEXT"
            case 14:
                return "CF_ENHMETAFILE"
            case 15:
                return "CF_HDROP"
            case 16:
                return "CF_LOCALE"
            case 17:
                return "CF_DIBV5"
            case 0x0080:
                return "CF_OWNERDISPLAY"
            case 0x0081:
                return "CF_DSPTEXT"
            case 0x0082:
                return "CF_DSPBITMAP"
            case 0x0083:
                return "CF_DSPMETAFILEPICT"
            case 0x008E:
                return "CF_DSPENHMETAFILE"
            case _:
                return f'"{fmtname}"' if (fmtname := self.formatname) else f"#{self.__fmt}"

    def __repr__(self) -> str:
        return f"ClipboardFormat({self.__str__()})"

    @property
    def formatname(self) -> str | None:
        fmt = self.__fmt
        for len in range(0xFF, 0x7FFFFFFF, 0xFF):
            buf = (c_wchar * len)()
            copied = _GetClipboardFormatNameW(fmt, buf, len)
            if copied <= len:
                return buf.value
        raise OverflowError

    @property
    def is_standard(self) -> bool:
        return 1 <= self.__fmt <= 18

    @property
    def is_private(self) -> bool:
        return 0x200 <= self.__fmt <= 0x2FF

    @property
    def is_gdiobj(self) -> bool:
        return 0x300 <= self.__fmt <= 0x3FF


_GetClipboardFormatNameW = _user32.GetClipboardFormatNameW
_GetClipboardFormatNameW.argtypes = (c_uint32, c_wchar_p, c_int32)


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


class StorageMedium(Structure):
    """STGMEDIUM構造体。"""

    class DummyUnion(Union):
        _fields_ = (
            ("bitmap_handle", c_void_p),
            ("metafilepict_handle", c_void_p),
            ("enhmetafile_handle", c_void_p),
            ("global_handle", c_void_p),
            ("file", c_wchar_p),
            ("stream_ptr", POINTER(IStream)),
            ("storage_ptr", POINTER(IUnknown)),  # TODO POINTER(IStorage)
        )

    __slots__ = ()
    _anonymous_ = ("u",)
    _fields_ = (
        ("_tymed", c_uint32),
        ("u", DummyUnion),
        ("unk_for_release", POINTER(IUnknown)),
    )

    if TYPE_CHECKING:
        bitmap_handle: int
        metafilepict_handle: int
        enhmetafile_handle: int
        global_handle: int
        file: str
        stream_ptr: _Pointer[IStream]
        storage_ptr: _Pointer[IUnknown]  # TODO: IStorage
        unk_for_release: _Pointer[IUnknown]

    @property
    def tymed(self) -> MediumType:
        return MediumType(self._tymed)

    def __del__(self) -> None:
        _ReleaseStgMedium(self)

    @property
    def bytes(self) -> bytes:
        match self.tymed:
            case MediumType.NULL:
                return bytes()
            case MediumType.HGLOBAL:
                with globalmem_lock(self.global_handle, c_void_p) as (p, size):
                    t = c_byte * size
                    return bytes(t.from_address(p.value or 0))
            case MediumType.ISTREAM:
                return ComStream(self.stream_ptr).read_bytes_all()
            # TODO: IStorage
            case _:
                raise TypeError


_ReleaseStgMedium = _ole32.ReleaseStgMedium
_ReleaseStgMedium.argtypes = (POINTER(StorageMedium),)
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
        ("_format", c_int16),
        ("ptd", c_void_p),
        ("aspect", c_uint32),
        ("index", c_int32),
        ("_tymed", c_uint32),
    )

    __slots__ = ()

    if TYPE_CHECKING:
        ptd: int  # POINTER(DVTARGETDEVICE)
        aspect: int
        index: int

    @property
    def format(self) -> ClipboardFormat:
        return ClipboardFormat(self._format)

    @property
    def tymed(self) -> MediumType:
        return MediumType(self._tymed)

    @tymed.setter
    def tymed(self, x: MediumType) -> None:
        self._tymed = int(x)

    @staticmethod
    def create_simple(format: ClipboardFormat) -> "FormatEtc":
        fmtetc = FormatEtc()
        fmtetc._format = format.value
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
        STDMETHOD(c_int32, "GetData", (POINTER(FormatEtc), POINTER(StorageMedium))),
        STDMETHOD(c_int32, "GetDataHere", (POINTER(FormatEtc), POINTER(StorageMedium))),
        STDMETHOD(c_int32, "QueryGetData", (POINTER(FormatEtc),)),
        STDMETHOD(c_int32, "GetCanonicalFormatEtc", (POINTER(FormatEtc), POINTER(FormatEtc))),
        STDMETHOD(c_int32, "SetData", (POINTER(FormatEtc), POINTER(StorageMedium), c_int32)),
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

    def get_data_nothrow(self, format: ClipboardFormat) -> ComResult[StorageMedium]:
        fmtetc = FormatEtc.create_simple(format)
        stgmed = StorageMedium()
        return cr(self.__o.GetData(byref(fmtetc), byref(stgmed)), stgmed)

    def get_data(self, format: ClipboardFormat) -> StorageMedium:
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
