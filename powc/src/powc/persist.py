"""永続化機能。IPersist、IPersistFile、IPersistStreamのラッパーです。"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_uint64, c_void_p, c_wchar_p
from typing import Any

from comtypes import GUID, STDMETHOD, IUnknown

from .core import ComResult, cotaskmem, cr, query_interface
from .stream import ComStream, IStream, StorageMode


class IPersist(IUnknown):
    """"""

    __slots__ = ()
    _iid_ = GUID("{0000010c-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "GetClassID", (POINTER(GUID),)),
    ]


class IPersistFile(IPersist):
    """"""

    __slots__ = ()
    _iid_ = GUID("{0000010b-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "IsDirty", ()),
        STDMETHOD(c_int32, "Load", (c_wchar_p, c_uint32)),
        STDMETHOD(c_int32, "Save", (c_wchar_p, c_int32)),
        STDMETHOD(c_int32, "SaveCompleted", (c_wchar_p,)),
        STDMETHOD(c_int32, "GetCurFile", (POINTER(c_wchar_p),)),
    ]


class IPersistStream(IPersist):
    """"""

    __slots__ = ()
    _iid_ = GUID("{00000109-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "IsDirty", ()),
        STDMETHOD(c_int32, "Load", (POINTER(IStream),)),
        STDMETHOD(c_int32, "Save", (POINTER(IStream), c_int32)),
        STDMETHOD(c_int32, "GetSizeMax", (POINTER(c_uint64),)),
    ]


class Persist:
    """IPersistインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IPersist)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPersist)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def get_clsid_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetClassID(byref(x)), x)

    @property
    def get_clsid(self) -> GUID:
        return self.get_clsid_nothrow.value


class PersistFile(Persist):
    """ファイルによる永続化管理。IPersistFileインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IPersistFile)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IPersistFile)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def isdirty_nothrow(self) -> ComResult[bool]:
        hr = self.__o.IsDirty()
        return cr(hr, hr == 0)

    @property
    def isdirty(self) -> bool:
        return self.isdirty_nothrow.value

    def load_nothrow(self, path: str, mode: StorageMode) -> ComResult[bool]:
        hr = self.__o.Load(path, int(mode))
        return cr(hr, hr == 0)

    def load(self, path: str, mode: StorageMode) -> bool:
        return self.load_nothrow(path, mode).value

    def save_nothrow(self, path: str, remembers: bool) -> ComResult[bool]:
        hr = self.__o.Save(path, 1 if remembers else 0)
        return cr(hr, hr == 0)

    def save(self, path: str, remembers: bool) -> bool:
        return self.save_nothrow(path, remembers).value

    def save_completed_nothrow(self, path: str) -> ComResult[bool]:
        hr = self.__o.SaveCompleted(path)
        return cr(hr, hr == 0)

    def savecompleted(self, path: str) -> bool:
        return self.save_completed_nothrow(path).value

    @property
    def curfile_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetCurFile(p), p.value or "")

    @property
    def curfile(self) -> str:
        return self.curfile_nothrow.value


class PersistStream(Persist):
    """ストリームによる永続化管理。IPersistStreamインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IPersistStream)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IPersistStream)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def is_dirty_nothrow(self) -> ComResult[bool]:
        hr = self.__o.IsDirty()
        return cr(hr, hr == 0)

    @property
    def is_dirty(self) -> bool:
        return self.is_dirty_nothrow.value

    def load_nothrow(self, stream: ComStream) -> ComResult[None]:
        return cr(self.__o.Load(stream.wrapped_obj), None)

    def load(self, stream: ComStream) -> None:
        return self.load(stream)

    def save_nothrow(self, stream: ComStream, clears_dirty: bool) -> ComResult[None]:
        return cr(self.__o.Save(stream.wrapped_obj, clears_dirty), None)

    def save(self, stream: ComStream, clears_dirty: bool) -> None:
        return self.save(stream, clears_dirty)

    @property
    def sizemax_nothrow(self) -> ComResult[int]:
        x = c_uint64()
        return cr(self.__o.GetSizeMax(byref(x)), x.value)

    @property
    def sizemax(self) -> int:
        return self.sizemax_nothrow.value
