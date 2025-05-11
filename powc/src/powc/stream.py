"""IStreamラッパー。"""

from collections.abc import Buffer
from contextlib import contextmanager
from ctypes import (
    POINTER,
    Structure,
    byref,
    c_byte,
    c_int32,
    c_int64,
    c_uint32,
    c_uint64,
    c_void_p,
    c_wchar_p,
    cast,
)
from dataclasses import dataclass
from datetime import datetime
from enum import IntEnum, IntFlag
from typing import Any, Generator

from comtypes import GUID, STDMETHOD, IUnknown
from comtypes.hresult import E_FAIL, S_OK

from . import _shlwapi
from .core import ComResult, cotaskmem, cr, query_interface
from .datetime import filetimeint64_to_datetime


class ISequentialStream(IUnknown):
    _iid_ = GUID("{0c733a30-2a1c-11ce-ade5-00aa0044773d}")
    _methods_ = [
        STDMETHOD(c_int32, "Read", (c_void_p, c_uint32, POINTER(c_uint32))),
        STDMETHOD(c_int32, "Write", (c_void_p, c_uint32, POINTER(c_uint32))),
    ]

    __slots__ = ()


class StorageMode(IntFlag):
    """STGM"""

    DIRECT = 0x00000000
    TRANSACTED = 0x00010000
    SIMPLE = 0x08000000
    READ = 0x00000000
    WRITE = 0x00000001
    READWRITE = 0x00000002
    SHARE_DENY_NONE = 0x00000040
    SHARE_DENY_READ = 0x00000030
    SHARE_DENY_WRITE = 0x00000020
    SHARE_EXCLUSIVE = 0x00000010
    PRIORITY = 0x00040000
    DELETEONRELEASE = 0x04000000
    NOSCRATCH = 0x00100000
    CREATE = 0x00001000
    CONVERT = 0x00020000
    FAILIFTHERE = 0x00000000
    NOSNAPSHOT = 0x00200000
    DIRECT_SWMR = 0x00400000


class _STATSTG(Structure):
    _fields_ = (
        ("pwcsName", c_void_p),
        ("type", c_uint32),
        ("cbSize", c_uint64),
        ("mtime", c_int64),
        ("ctime", c_int64),
        ("atime", c_int64),
        ("grfMode", c_uint32),
        ("grfLocksSupported", c_uint32),
        ("clsid", GUID),
        ("grfStateBits", c_uint32),
        ("reserved", c_uint32),
    )


@dataclass
class ComStorageStat:
    name: str | None
    type: int
    size: int
    mtime: datetime
    ctime: datetime
    atime: datetime
    mode: int
    locks_supported: int
    clsid: GUID
    statebits: int
    reserved: int


class StorageType(IntEnum):
    """STGTY"""

    STORAGE = 1
    STREAM = 2
    LOCKBYTES = 3
    PROPERTY = 4


class StreamSeek(IntEnum):
    """STREAM_SEEK"""

    SET = 0
    CUR = 1
    END = 2


class LockType(IntFlag):
    """LOCKTYPE"""

    LOCK_WRITE = 1
    LOCK_EXCLUSIVE = 2
    LOCK_ONLYONCE = 4


class StorageCommit(IntFlag):
    """STGC"""

    DEFAULT = 0
    OVERWRITE = 1
    ONLYIFCURRENT = 2
    DANGEROUSLYCOMMITMERELYTODISKCACHE = 4
    CONSOLIDATE = 8


class StatFlag(IntEnum):
    """STATFLAG"""

    DEFAULT = 0
    NONAME = 1
    # NOOPEN = 2 # Unused


class IStream(ISequentialStream):
    """"""

    _iid_ = GUID("{0000000c-0000-0000-C000-000000000046}")

    __slots__ = ()


IStream._methods_ = [
    STDMETHOD(c_int32, "Seek", (c_int64, c_uint32, POINTER(c_uint64))),
    STDMETHOD(c_int32, "SetSize", (c_uint64,)),
    STDMETHOD(c_int32, "CopyTo", (POINTER(IStream), c_uint64, POINTER(c_uint64), POINTER(c_uint64))),
    STDMETHOD(c_int32, "Commit", (c_uint32,)),
    STDMETHOD(c_int32, "Revert", ()),
    STDMETHOD(c_int32, "LockRegion", (c_uint64, c_uint64, c_uint32)),
    STDMETHOD(c_int32, "UnlockRegion", (c_uint64, c_uint64, c_uint32)),
    STDMETHOD(c_int32, "Stat", (POINTER(_STATSTG), c_uint32)),
    STDMETHOD(c_int32, "Clone", (POINTER(IStream),)),
]

_SHCreateStreamOnFileEx = _shlwapi.SHCreateStreamOnFileEx
_SHCreateStreamOnFileEx.argtypes = (c_wchar_p, c_uint32, c_uint32, c_int32, POINTER(IStream), POINTER(POINTER(IStream)))
_SHCreateStreamOnFileEx.restype = c_int32

_SHCreateMemStream = _shlwapi.SHCreateMemStream
_SHCreateMemStream.argtypes = (POINTER(c_byte), c_uint32)
_SHCreateMemStream.restype = POINTER(IStream)


class ComStream:
    """COMストリーム。IStreamインターフェイスのラッパーです。"""

    __o: Any  # IStream

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IStream)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_on_file_nothrow(path: str, mode: StorageMode | int, creates: bool) -> "ComResult[ComStream]":
        p = POINTER(IStream)()
        return cr(_SHCreateStreamOnFileEx(path, int(mode), 0, 1 if creates else 0, None, byref(p)), ComStream(p))

    @staticmethod
    def create_on_file(path: str, mode: StorageMode | int, creates: bool) -> "ComStream":
        return ComStream.create_on_file_nothrow(path, mode, creates).value

    @staticmethod
    def openread_on_file_nothrow(path: str) -> ComResult["ComStream"]:
        return ComStream.create_on_file_nothrow(path, StorageMode.READ, False)

    @staticmethod
    def openread_on_file(path: str) -> "ComStream":
        return ComStream.create_on_file(path, StorageMode.READ, False)

    @staticmethod
    def create_on_mem_nothrow(buffer: Buffer) -> "ComResult[ComStream]":
        if buffer is None:
            p = _SHCreateMemStream(None, 0)
        else:
            mv = memoryview(buffer)
            if mv.readonly:
                buf = (c_byte * len(mv)).from_buffer_copy(mv)
            else:
                buf = (c_byte * len(mv)).from_buffer(mv)
            p = _SHCreateMemStream(buf, len(buf))
        return cr(S_OK if p else E_FAIL, ComStream(p))

    @staticmethod
    def create_on_mem(buffer: Buffer) -> "ComStream":
        return ComStream.create_on_mem_nothrow(buffer).value

    @dataclass
    class BytesAndSize:
        array: bytes
        size: int

    def read_bytes_nothrow(self, size: int) -> ComResult[BytesAndSize]:
        a = (c_byte * size)()
        x = c_uint32()
        return cr(self.__o.Read(a, size, byref(x)), ComStream.BytesAndSize(bytes(a), x.value))

    def read_bytes(self, size: int) -> BytesAndSize:
        return self.read_bytes_nothrow(size).value

    def write_bytes_nothrow(self, data: Buffer) -> ComResult[int]:
        x = c_uint32()
        mv = memoryview(data)
        t = c_byte * len(mv)
        buf = t.from_buffer_copy(mv) if mv.readonly else t.from_buffer(mv)
        return cr(self.__o.Write(buf, len(buf), byref(x)), x.value)

    def write_bytes(self, data) -> int:
        return self.write_bytes_nothrow(data).value

    def seek_nothrow(self, move: int, origin: StreamSeek | int) -> ComResult[int]:
        x = c_uint64()
        return cr(self.__o.Seek(move, origin, byref(x)), x.value)

    def seek(self, move: int, origin: StreamSeek | int) -> int:
        return self.seek_nothrow(move, origin).value

    @property
    def size(self) -> int:
        return self.stat.size

    @size.setter
    def size(self, size: int) -> None:
        return self.set_size_nothrow(size).value

    @property
    def size_nothrow(self) -> ComResult[int]:
        x = self.stat_nothrow
        return cr(x.hr, x.value_unchecked.size if x else 0)  # type: ignore

    def set_size_nothrow(self, size: int) -> ComResult[None]:
        return cr(self.__o.SetSize(size), None)

    # STDMETHOD(c_int32, "CopyTo", (POINTER(IStream), c_uint64, POINTER(c_uint64), POINTER(c_uint64))),

    def commit_nothrow(self, flags: StorageCommit | int) -> ComResult[None]:
        return cr(self.__o.Commit(int(flags)), None)

    def commit(self, flags: StorageCommit | int) -> None:
        return self.commit_nothrow(flags).value

    def revert_nothrow(self) -> ComResult[None]:
        return cr(self.__o.Revert(), None)

    def revert(self) -> None:
        return self.revert_nothrow().value

    def lock_region_nothrow(self, offset: int, size: int, locktype: LockType | int) -> ComResult[None]:
        return cr(self.__o.LockRegion(offset, size, int(locktype)), None)

    def unlock_region_nothrow(self, offset: int, size: int, locktype: LockType | int) -> ComResult[None]:
        return cr(self.__o.UnlockRegion(offset, size, int(locktype)), None)

    def lock_region(self, offset: int, size: int, locktype: LockType | int) -> None:
        return self.lock_region_nothrow(offset, size, locktype).value

    def unlock_region(self, offset: int, size: int, locktype: LockType | int) -> None:
        return self.unlock_region_nothrow(offset, size, locktype).value

    def get_stat_nothrow(self, flags: StatFlag | int) -> ComResult[ComStorageStat]:
        x = _STATSTG()
        hr = self.__o.Stat(byref(x), int(flags))
        if hr < 0:
            return cr(hr, ComStorageStat("", 0, 0, datetime.now(), datetime.now(), datetime.now(), 0, 0, GUID(), 0, 0))
        with cotaskmem(x.pwcsName):
            return cr(
                hr,
                ComStorageStat(
                    cast(x.pwcsName, c_wchar_p).value,
                    x.type,
                    x.cbSize,
                    filetimeint64_to_datetime(x.mtime),
                    filetimeint64_to_datetime(x.ctime),
                    filetimeint64_to_datetime(x.atime),
                    x.grfMode,
                    x.grfLocksSupported,
                    x.clsid,
                    x.grfStateBits,
                    x.reserved,
                ),
            )

    def get_stat(self, flags: StatFlag | int) -> ComStorageStat:
        return self.get_stat_nothrow(flags).value

    @property
    def stat_nothrow(self) -> ComResult[ComStorageStat]:
        return self.get_stat_nothrow(StatFlag.DEFAULT)

    @property
    def stat(self) -> ComStorageStat:
        return self.get_stat(StatFlag.DEFAULT)

    def clone_nothrow(self) -> "ComResult[ComStream]":
        x = POINTER(IStream)()
        return cr(self.__o.Clone(byref(x)), ComStream(x))

    def clone(self) -> "ComStream":
        return self.clone_nothrow().value

    @property
    def pos_nothrow(self) -> ComResult[int]:
        return self.seek_nothrow(0, StreamSeek.SET)

    @property
    def pos(self) -> int:
        return self.seek(0, StreamSeek.SET)

    def set_pos_nothrow(self, pos: int) -> ComResult[int]:
        return self.seek_nothrow(pos, StreamSeek.SET)

    @pos.setter
    def pos(self, pos: int) -> None:
        self.seek(pos, StreamSeek.SET)

    @contextmanager
    def keep_pos(self) -> Generator[None, None, None]:
        pos = self.pos
        yield None
        self.pos = pos

    def read_bytes_all(self) -> bytes:
        with self.keep_pos():
            self.pos = 0
            return self.read_bytes(self.size).array
