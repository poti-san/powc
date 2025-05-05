"""COMオブジェクトの基本操作クラス。"""

from ctypes import (
    POINTER,
    Structure,
    byref,
    c_int32,
    c_uint32,
    c_uint64,
    c_void_p,
    c_wchar_p,
    sizeof,
)
from enum import IntEnum, IntFlag
from types import NotImplementedType
from typing import Any, Iterator

from comtypes import GUID, STDMETHOD, IUnknown
from comtypes.hresult import E_FAIL

from . import _ole32
from .core import (
    ComResult,
    IUnknownPointer,
    IUnknownWrapper,
    check_hresult,
    cotaskmem,
    cr,
    query_interface,
)
from .datetime import FILETIME
from .persist import IPersistStream
from .stream import ComStream


class IEnumUnknown(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000100-0000-0000-C000-000000000046}")


IEnumUnknown._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(POINTER(IUnknown)), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(IEnumUnknown),)),
]


class IUnknownEnumerator:
    """IEnumUnknownインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IEnumUnknown)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumUnknown)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __iter__(self) -> Iterator[IUnknownPointer]:
        check_hresult(self.__o.Reset())
        hr = 0
        x = POINTER(IUnknown)()
        while (hr := self.__o.Next(1, byref(x), None)) == 0:
            yield x
            x = POINTER(IUnknown)()  # TODO: 必要？
        check_hresult(hr)

    def clone_nothrow(self) -> "ComResult[IUnknownEnumerator]":
        x = POINTER(IEnumUnknown)()
        return cr(self.__o.Clone(byref(x)), IUnknownEnumerator(x))

    def clone(self) -> "IUnknownEnumerator":
        return self.clone_nothrow().value


class IEnumString(IUnknown):
    """IEnumStringインターフェイス。"""

    __slots__ = ()
    _iid_ = GUID("{00000101-0000-0000-C000-000000000046}")


IEnumString._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(c_wchar_p), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IEnumString)),)),
    STDMETHOD(c_int32, "", ()),
]


class ComStringEnumerator:
    """IEnumStringインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IEnumString)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumString)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __iter__(self) -> Iterator[str]:
        check_hresult(self.__o.Reset())
        hr = 0
        x = c_wchar_p()
        while (hr := self.__o.Next(1, byref(x), None)) == 0:
            with cotaskmem(x):
                yield x.value or ""
        check_hresult(hr)

    def clone_nothrow(self) -> "ComResult[ComStringEnumerator]":
        x = POINTER(IEnumString)()
        return cr(self.__o.Clone(byref(x)), ComStringEnumerator(x))

    def clone(self) -> "ComStringEnumerator":
        return self.clone_nothrow().value


# 注意
# IRunningObjectTableはIMonikerやIEnumMonikerと循環参照します。
# インターフェイスはここで定義しますが、メソッドリストとクラスはMonikerやMonikerEnumeratorより後で実装します。
class IRunningObjectTable(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000010-0000-0000-C000-000000000046}")


class BindOptions(Structure):
    """BIND_OPTS構造体のラッパー。"""

    _fields_ = (
        ("cb_struct", c_uint32),
        ("flags", c_uint32),
        ("mode", c_uint32),
        ("tickcount_deadline", c_uint32),
    )
    __slots__ = ()


class BindOptions2(BindOptions):
    _fields_ = (
        ("track_flags", c_uint32),
        ("class_context", c_uint32),
        ("locale", c_uint32),
        ("server_info", c_void_p),  # TODO: COSERVERINFO*
    )
    __slots__ = ()


class BindOptions3(BindOptions2):
    _fields_ = (("window_handle", c_void_p),)
    __slots__ = ()


class BindFlag(IntFlag):
    """BIND_FLAGS"""

    MAY_BOTHER_USER = 1
    JUST_TEST_EXISTENCE = 2


# IBindCtxはIRunningObjectTableを介してIMonikerと循環参照します。
# インターフェイスはここで定義しますが、クラスはMonikerやRunningObjectTableより後で実装します。
class IBindCtx(IUnknown):
    """IBindCtxインターフェイスのラッパー。"""

    __slots__ = ()
    _iid_ = GUID("{0000000e-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "RegisterObjectBound", (POINTER(IUnknown),)),
        STDMETHOD(c_int32, "RevokeObjectBound", (POINTER(IUnknown),)),
        STDMETHOD(c_int32, "ReleaseBoundObjects", ()),
        STDMETHOD(c_int32, "SetBindOptions", (c_void_p,)),
        STDMETHOD(c_int32, "GetBindOptions", (c_void_p,)),
        STDMETHOD(c_int32, "GetRunningObjectTable", (POINTER(POINTER(IRunningObjectTable)),)),
        STDMETHOD(c_int32, "RegisterObjectParam", (c_wchar_p, POINTER(IUnknown))),
        STDMETHOD(c_int32, "GetObjectParam", (c_wchar_p, POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "EnumObjectParam", (POINTER(IEnumString),)),
        STDMETHOD(c_int32, "RevokeObjectParam", (c_wchar_p,)),
    ]


class MonikerSystem(IntEnum):
    """MKSYS"""

    NONE = 0
    GENERIC_COMPOSITE = 1
    FILE_MONIKER = 2
    ANTI_MONIKER = 3
    ITEM_MONIKER = 4
    POINTER_MONIKER = 5
    CLASS_MONIKER = 7
    OBJREF_MONIKER = 8
    SESSION_MONIKER = 9
    LUA_MONIKER = 10


class MonikerReduce(IntFlag):
    """MKREDUCE"""

    ONE = 3 << 16
    TOUSER = 2 << 16
    THROUGHUSER = 1 << 16
    ALL = 0


class IMoniker(IPersistStream):
    __slots__ = ()
    _iid_ = GUID("{0000000f-0000-0000-C000-000000000046}")


class IEnumMoniker(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000102-0000-0000-C000-000000000046}")


IMoniker._methods_ = [
    STDMETHOD(c_int32, "BindToObject", (POINTER(IBindCtx), POINTER(IMoniker), POINTER(GUID), POINTER(POINTER(GUID)))),
    STDMETHOD(
        c_int32, "BindToStorage", (POINTER(IBindCtx), POINTER(IMoniker), POINTER(GUID), POINTER(POINTER(IUnknown)))
    ),
    STDMETHOD(c_int32, "Reduce", (POINTER(IBindCtx), c_uint32, POINTER(POINTER(IMoniker)), POINTER(POINTER(IMoniker)))),
    STDMETHOD(c_int32, "ComposeWith", (POINTER(IMoniker), c_int32, POINTER(POINTER(IMoniker)))),
    STDMETHOD(c_int32, "Enum", (c_int32, POINTER(POINTER(IEnumMoniker)))),
    STDMETHOD(c_int32, "IsEqual", (POINTER(IMoniker),)),
    STDMETHOD(c_int32, "Hash", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "IsRunning", (POINTER(IBindCtx), POINTER(IMoniker), POINTER(IMoniker))),
    STDMETHOD(c_int32, "GetTimeOfLastChange", (POINTER(IBindCtx), POINTER(IMoniker), POINTER(FILETIME))),
    STDMETHOD(c_int32, "Inverse", (POINTER(POINTER(IMoniker)),)),
    STDMETHOD(c_int32, "CommonPrefixWith", (POINTER(IMoniker), POINTER(POINTER(IMoniker)))),
    STDMETHOD(c_int32, "RelativePathTo", (POINTER(IMoniker), POINTER(POINTER(IMoniker)))),
    STDMETHOD(c_int32, "GetDisplayName", (POINTER(IBindCtx), POINTER(IMoniker), POINTER(c_wchar_p))),
    STDMETHOD(
        c_int32,
        "ParseDisplayName",
        (POINTER(IBindCtx), POINTER(IMoniker), c_wchar_p, POINTER(c_uint32), POINTER(POINTER(IMoniker))),
    ),
    STDMETHOD(c_int32, "IsSystemMoniker", (POINTER(c_int32),)),
]

IEnumMoniker._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(POINTER(IMoniker)), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(IEnumMoniker),)),
]


# TODO
class Moniker:
    """IMonikerインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IMoniker)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IMoniker)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def clsid_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetClassID(byref(x)), x)

    @property
    def clsid(self) -> GUID:
        return self.clsid_nothrow.value

    @property
    def is_dirty_nothrow(self) -> ComResult[bool]:
        hr = self.__o.IsDirty()
        return cr(hr, hr == 0)

    @property
    def is_dirty(self) -> bool:
        return self.is_dirty_nothrow.value

    def load_nothrow(self, stream: ComStream) -> ComResult[None]:
        return self.__o.Load(stream.wrapped_obj)

    def load(self, stream: ComStream) -> None:
        return self.load_nothrow(stream).value

    def save_nothrow(self, stream: ComStream, clears_dirty: bool) -> ComResult[None]:
        return self.__o.Save(stream.wrapped_obj, 1 if clears_dirty else 0)

    def save(self, stream: ComStream, clears_dirty: bool) -> None:
        return self.save_nothrow(stream, clears_dirty).value

    @property
    def sizemax_nothrow(self) -> ComResult[int]:
        x = c_uint64()
        return cr(self.__o.GetSizeMax(byref(x)), x.value)

    @property
    def sizemax(self) -> int:
        return self.sizemax_nothrow.value

    # virtual /* [local] */ HRESULT STDMETHODCALLTYPE BindToObject(
    #     /* [annotation][unique][in] */
    #     _In_  IBindCtx *pbc,
    #     /* [annotation][unique][in] */
    #     _In_opt_  IMoniker *pmkToLeft,
    #     /* [annotation][in] */
    #     _In_  REFIID riidResult,
    #     /* [annotation][iid_is][out] */
    #     _Outptr_  void **ppvResult) = 0;

    # virtual /* [local] */ HRESULT STDMETHODCALLTYPE BindToStorage(
    #     /* [annotation][unique][in] */
    #     _In_  IBindCtx *pbc,
    #     /* [annotation][unique][in] */
    #     _In_opt_  IMoniker *pmkToLeft,
    #     /* [annotation][in] */
    #     _In_  REFIID riid,
    #     /* [annotation][iid_is][out] */
    #     _Outptr_  void **ppvObj) = 0;

    # virtual HRESULT STDMETHODCALLTYPE Reduce(
    #     /* [unique][in] */ __RPC__in_opt IBindCtx *pbc,
    #     /* [in] */ DWORD dwReduceHowFar,
    #     /* [unique][out][in] */ __RPC__deref_opt_inout_opt IMoniker **ppmkToLeft,
    #     /* [out] */ __RPC__deref_out_opt IMoniker **ppmkReduced) = 0;

    # virtual HRESULT STDMETHODCALLTYPE ComposeWith(
    #     /* [unique][in] */ __RPC__in_opt IMoniker *pmkRight,
    #     /* [in] */ BOOL fOnlyIfNotGeneric,
    #     /* [out] */ __RPC__deref_out_opt IMoniker **ppmkComposite) = 0;

    def __enum(self, forward: bool) -> "ComResult[MonikerEnumerator]":
        x = POINTER(IEnumMoniker)()
        # IMoniker.Enumは列挙可能なコンポーネントがない場合にS_OKとx=NULLを返します。
        # 成功終了にもかかわらずMonikerEnumerator.__oがNoneで想定外の動作となるため、特別対応します。
        hr = self.__o.Enum(1 if forward else 0, byref(x))
        if hr == 0 and x.contents.value == 0:
            hr = E_FAIL
        return cr(hr, MonikerEnumerator(x))

    @property
    def enum_forward_nothrow(self) -> "ComResult[MonikerEnumerator]":
        return self.__enum(True)

    @property
    def enum_forward(self) -> "MonikerEnumerator":
        return self.enum_forward_nothrow.value

    @property
    def enum_backward_nothrow(self) -> "ComResult[MonikerEnumerator]":
        return self.__enum(False)

    @property
    def enum_backward(self) -> "MonikerEnumerator":
        return self.enum_backward_nothrow.value

    def is_equal_nothrow(self, other: "Moniker") -> ComResult[bool]:
        hr = self.__o.IsEqual(other.wrapped_obj)
        return cr(hr, hr == 0)

    def is_equal(self, other: "Moniker") -> bool:
        return self.is_equal_nothrow(other).value

    def __cmp__(self, other) -> bool | NotImplementedType:
        if isinstance(other, Moniker):
            return self.is_equal(other)
        return NotImplemented

    @property
    def hash_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.Hash(byref(x)), x.value)

    @property
    def hash(self) -> int:
        return self.hash_nothrow.value

    def __hash__(self) -> int:
        return self.hash

    # virtual HRESULT STDMETHODCALLTYPE IsRunning(
    #     /* [unique][in] */ __RPC__in_opt IBindCtx *pbc,
    #     /* [unique][in] */ __RPC__in_opt IMoniker *pmkToLeft,
    #     /* [unique][in] */ __RPC__in_opt IMoniker *pmkNewlyRunning) = 0;

    # virtual HRESULT STDMETHODCALLTYPE GetTimeOfLastChange(
    #     /* [unique][in] */ __RPC__in_opt IBindCtx *pbc,
    #     /* [unique][in] */ __RPC__in_opt IMoniker *pmkToLeft,
    #     /* [out] */ __RPC__out FILETIME *pFileTime) = 0;

    def inversed_nothrow(self) -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(self.__o.Inverse(byref(x)), Moniker(x))

    def inversed(self) -> "Moniker":
        return self.inversed_nothrow().value

    def get_common_prefix_with_nothrow(self, other: "Moniker") -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(self.__o.CommonPrefixWith(other.wrapped_obj, byref(x)), Moniker(x))

    def get_common_prefix_with(self, other: "Moniker") -> "Moniker":
        return self.get_common_prefix_with_nothrow(other).value

    def get_relpath_to_nothrow(self, other: "Moniker") -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(self.__o.RelativePathTo(other.wrapped_obj, byref(x)), Moniker(x))

    def get_relpath_to(self, other: "Moniker") -> "Moniker":
        return self.get_relpath_to_nothrow(other).value

    def get_displayname_nothrow(self, bc: "BindCtx", left: "Moniker | None" = None) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(
                self.__o.GetDisplayName(bc.wrapped_obj, left.wrapped_obj if left else None, byref(p)),
                p.value or "",
            )

    def get_displayname(self, bc: "BindCtx", left: "Moniker | None" = None) -> str:
        return self.get_displayname_nothrow(bc, left).value

    @property
    def displayname_nothrow(self) -> ComResult[str]:
        return self.get_displayname_nothrow(BindCtx.create())

    @property
    def displayname(self) -> str:
        return self.get_displayname_nothrow(BindCtx.create()).value

    def parse_displayname_nothrow(
        self,
        displayname: str,
        left: "Moniker | None" = None,
        bc: "BindCtx | None" = None,
    ) -> "ComResult[Moniker]":
        eaten = c_uint32()
        x = POINTER(IMoniker)()
        return cr(
            self.__o.ParseDisplayName(
                bc.wrapped_obj if bc else None,
                left.wrapped_obj if left else None,
                displayname,
                byref(eaten),
                byref(x),
            ),
            Moniker(x),
        )

    def parse_displayname(
        self,
        displayname: str,
        left: "Moniker | None" = None,
        bc: "BindCtx | None" = None,
    ) -> "Moniker":
        return self.parse_displayname_nothrow(displayname, left, bc).value

    @property
    def is_sysmoniker_nothrow(self) -> ComResult[MonikerSystem]:
        x = c_int32()
        return cr(self.__o.IsSystemMoniker(byref(x)), MonikerSystem(x.value))

    @property
    def is_sysmoniker(self) -> MonikerSystem:
        return self.is_sysmoniker_nothrow.value

    @staticmethod
    def create_classmoniker_nothrow(clsid: GUID) -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(_CreateClassMoniker(clsid, byref(x)), Moniker(x))

    @staticmethod
    def create_classmoniker(clsid: GUID) -> "Moniker":
        return Moniker.create_classmoniker_nothrow(clsid).value

    @staticmethod
    def create_objref_nothrow(wrapper: IUnknownWrapper) -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(_CreateObjrefMoniker(wrapper.wrapped_obj, byref(x)), Moniker(x))

    @staticmethod
    def create_objref(wrapper: IUnknownWrapper) -> "Moniker":
        return Moniker.create_objref_nothrow(wrapper).value

    @staticmethod
    def create_pointer_nothrow(wrapper: IUnknownWrapper) -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(_CreatePointerMoniker(wrapper.wrapped_obj, byref(x)), Moniker(x))

    @staticmethod
    def create_pointer(wrapper: IUnknownWrapper) -> "Moniker":
        return Moniker.create_pointer_nothrow(wrapper).value

    @staticmethod
    def create_file_nothrow(pathname: str) -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(_CreateFileMoniker(pathname, byref(x)), Moniker(x))

    @staticmethod
    def create_file(pathname: str) -> "Moniker":
        return Moniker.create_file_nothrow(pathname).value

    @staticmethod
    def create_item_nothrow(delimiter: str, item: str) -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(_CreateItemMoniker(delimiter, item, byref(x)), Moniker(x))

    @staticmethod
    def create_item(delimiter: str, item: str) -> "Moniker":
        return Moniker.create_item_nothrow(delimiter, item).value

    @staticmethod
    def create_genericcomposite_nothrow(first: "Moniker", rest: "Moniker") -> "ComResult[Moniker]":
        x = POINTER(IMoniker)()
        return cr(_CreateGenericComposite(first.wrapped_obj, rest.wrapped_obj, byref(x)), Moniker(x))

    @staticmethod
    def create_genericcomposite(first: "Moniker", rest: "Moniker") -> "Moniker":
        return Moniker.create_genericcomposite_nothrow(first, rest).value


_CreateClassMoniker = _ole32.CreateClassMoniker
_CreateClassMoniker.argtypes = (POINTER(GUID), POINTER(POINTER(IMoniker)))

_CreateObjrefMoniker = _ole32.CreateObjrefMoniker
_CreateObjrefMoniker.argtypes = (POINTER(IUnknown), POINTER(POINTER(IMoniker)))

_CreatePointerMoniker = _ole32.CreatePointerMoniker
_CreatePointerMoniker.argtypes = (POINTER(IUnknown), POINTER(POINTER(IMoniker)))

_CreateFileMoniker = _ole32.CreateFileMoniker
_CreateFileMoniker.argtypes = (c_wchar_p, POINTER(POINTER(IMoniker)))

_CreateItemMoniker = _ole32.CreateItemMoniker
_CreateItemMoniker.argtypes = (c_wchar_p, c_wchar_p, POINTER(POINTER(IMoniker)))

_CreateGenericComposite = _ole32.CreateGenericComposite
_CreateGenericComposite.argtypes = (POINTER(IMoniker), POINTER(IMoniker), POINTER(POINTER(IMoniker)))


class MonikerEnumerator:
    """IEnumMonikerインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IEnumMoniker)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumMoniker)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __iter__(self) -> Iterator[Moniker]:
        check_hresult(self.__o.Reset())
        hr = 0
        x = POINTER(IMoniker)()
        while (hr := self.__o.Next(1, byref(x), None)) == 0:
            yield Moniker(x)
        check_hresult(hr)

    def clone_nothrow(self) -> "ComResult[MonikerEnumerator]":
        x = POINTER(IEnumMoniker)()
        return cr(self.__o.Clone(byref(x)), MonikerEnumerator(x))

    def clone(self) -> "MonikerEnumerator":
        return self.clone_nothrow().value


# IRunningObjectTableは循環参照するため、MonikerやMonikerEnumeratorより後で定義します。
# 詳細はIRunningObjectTableクラスのコメントを参照してください。
IRunningObjectTable._methods_ = [
    STDMETHOD(c_int32, "Register", (c_uint32, POINTER(IUnknown), POINTER(IMoniker), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Revoke", (c_uint32,)),
    STDMETHOD(c_int32, "IsRunning", (POINTER(IMoniker),)),
    STDMETHOD(c_int32, "GetObject", (POINTER(IMoniker), POINTER(POINTER(IUnknown)))),
    STDMETHOD(c_int32, "NoteChangeTime", (c_uint32, POINTER(FILETIME))),
    STDMETHOD(c_int32, "GetTimeOfLastChange", (POINTER(IMoniker), POINTER(FILETIME))),
    STDMETHOD(c_int32, "EnumRunning", (POINTER(POINTER(IEnumMoniker)),)),
]


class RunningObjectTableFlag(IntFlag):
    """ROTFLAGS_*定数"""

    REGISTRATION_KEEPS_ALIVE = 0x1
    ALLOW_ANY_CLIENT = 0x2


# IRunningObjectTableは循環参照するため、MonikerやMonikerEnumeratorより後で定義します。
# 詳細はIRunningObjectTableクラスのコメントを参照してください。
class RunningObjectTable:
    """IRunningObjectTableインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IRunningObjectTable)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IRunningObjectTable)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_nothrow() -> "ComResult[RunningObjectTable]":
        x = POINTER(IRunningObjectTable)()
        return cr(_GetRunningObjectTable(0, byref(x)), RunningObjectTable(x))

    @staticmethod
    def create() -> "RunningObjectTable":
        return RunningObjectTable.create_nothrow().value

    def register_nothrow(
        self, flags: RunningObjectTableFlag, object: IUnknownPointer, objectname: Moniker
    ) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.Register(int(flags), object, objectname.wrapped_obj, byref(x)), x.value)

    def register(self, flags: RunningObjectTableFlag, object: IUnknownPointer, objectname: Moniker) -> int:
        return self.register_nothrow(flags, object, objectname).value

    def revoke_nothrow(self, register: int) -> ComResult[None]:
        return cr(self.__o.Revoke(register), None)

    def revoke(self, register: int) -> None:
        return self.revoke_nothrow(register).value

    def is_running_nothrow(self, objectname: Moniker) -> ComResult[bool]:
        hr = self.__o.IsRunning(objectname.wrapped_obj)
        return cr(hr, hr == 0)

    def is_running(self, objectname: Moniker) -> bool:
        return self.is_running_nothrow(objectname).value

    def get_object_raw_nothrow(self, objectname: Moniker) -> ComResult[IUnknownPointer]:
        x = POINTER(IUnknown)()
        return cr(self.__o.GetObject(objectname.wrapped_obj, byref(x)), x)

    def get_object_raw(self, objectname: Moniker) -> IUnknownPointer:
        return self.get_object_raw_nothrow(objectname).value

    def note_changetime_nothrow(self, register: c_uint32, time: FILETIME) -> ComResult[None]:
        return cr(self.__o.NoteChangeTime(register, time), None)

    def note_changetime(self, register: c_uint32, time: FILETIME) -> None:
        return self.note_changetime_nothrow(register, time).value

    def get_time_of_lastchange_nothrow(self, objectname: Moniker) -> ComResult[FILETIME]:
        x = FILETIME()
        return cr(self.__o.GetTimeOfLastChange(objectname.wrapped_obj, byref(x)), x)

    def get_time_of_lastchange(self, objectname: Moniker) -> FILETIME:
        return self.get_time_of_lastchange_nothrow(objectname).value

    @property
    def enumrunning_nothrow(self) -> ComResult[MonikerEnumerator]:
        x = POINTER(IEnumMoniker)()
        return cr(self.__o.EnumRunning(byref(x)), MonikerEnumerator(x))

    @property
    def enumrunning(self) -> MonikerEnumerator:
        return self.enumrunning_nothrow.value

    @property
    def moniker_iter(self) -> Iterator[Moniker]:
        yield from self.enumrunning

    @property
    def moniker_items(self) -> tuple[Moniker, ...]:
        return tuple(self.enumrunning)


_GetRunningObjectTable = _ole32.GetRunningObjectTable
_GetRunningObjectTable.argtypes = (c_uint32, POINTER(POINTER(IRunningObjectTable)))


class BindCtx:
    """IBindCtxインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IBindCtx)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IBindCtx)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_nothrow() -> ComResult["BindCtx"]:
        x = POINTER(IBindCtx)()
        return cr(_CreateBindCtx(0, byref(x)), BindCtx(x))

    @staticmethod
    def create() -> "BindCtx":
        return BindCtx.create_nothrow().value

    def register_objectbound_nothrow(self, obj: IUnknownWrapper) -> ComResult[None]:
        return cr(self.__o.RegisterObjectBound(obj.wrapped_obj), None)

    def register_objectbound(self, obj: IUnknownWrapper) -> None:
        return self.register_objectbound_nothrow(obj).value

    def revoke_objectbound_nothrow(self, obj: IUnknownWrapper) -> ComResult[None]:
        return cr(self.__o.RevokeObjectBound(obj.wrapped_obj), None)

    def revoke_objectbound(self, obj: IUnknownWrapper) -> None:
        return self.revoke_objectbound_nothrow(obj).value

    def release_boundobjects_nothrow(self) -> ComResult[None]:
        return cr(self.__o.ReleaseBoundObjects(), None)

    def release_boundobjects(self) -> None:
        return self.release_boundobjects_nothrow().value

    @property
    def bindoptions1_nothrow(self) -> ComResult[BindOptions]:
        x = BindOptions()
        x.cb_struct = sizeof(x)
        return cr(self.__o.GetBindOptions(byref(x)), x)

    @property
    def bindoptions1(self) -> BindOptions:
        return self.bindoptions1_nothrow.value

    def set_bindoptions1_nothrow(self, options: BindOptions) -> ComResult[None]:
        return self.__o.SetBindOptions(options)

    @bindoptions1.setter
    def bindoptions1(self, options: BindOptions) -> None:
        return self.set_bindoptions1_nothrow(options).value

    @property
    def bindoptions2_nothrow(self) -> ComResult[BindOptions2]:
        x = BindOptions2()
        x.cb_struct = sizeof(x)
        return cr(self.__o.GetBindOptions(byref(x)), x)

    @property
    def bindoptions2(self) -> BindOptions2:
        return self.bindoptions2_nothrow.value

    def set_bindoptions2_nothrow(self, options: BindOptions2) -> ComResult[None]:
        return self.__o.SetBindOptions(options)

    @bindoptions2.setter
    def bindoptions2(self, options: BindOptions2) -> None:
        return self.set_bindoptions2_nothrow(options).value

    @property
    def bindoptions3_nothrow(self) -> ComResult[BindOptions3]:
        x = BindOptions3()
        x.cb_struct = sizeof(x)
        return cr(self.__o.GetBindOptions(byref(x)), x)

    @property
    def bindoptions3(self) -> BindOptions3:
        return self.bindoptions3_nothrow.value

    def set_bindoptions3_nothrow(self, options: BindOptions3) -> ComResult[None]:
        return self.__o.SetBindOptions(options)

    @bindoptions3.setter
    def bindoptions3(self, options: BindOptions3) -> None:
        return self.set_bindoptions1_nothrow(options).value

    @property
    def bindoptions_nothrow(self) -> ComResult[BindOptions3]:
        return self.bindoptions3_nothrow

    @property
    def bindoptions(self) -> BindOptions3:
        return self.bindoptions3

    @property
    def rot_nothrow(self) -> ComResult[RunningObjectTable]:
        x = POINTER(IRunningObjectTable)()
        return cr(self.__o.GetRunningObjectTable(byref(x)), RunningObjectTable(x))

    @property
    def rot(self) -> RunningObjectTable:
        return self.rot_nothrow.value

    def register_objectparam_nothrow(self, key: str, obj: IUnknownWrapper) -> ComResult[None]:
        return cr(self.__o.RegisterObjectParam(key, obj.wrapped_obj), None)

    def register_objectparam(self, key: str, obj: IUnknownWrapper) -> None:
        return self.register_objectparam_nothrow(key, obj).value

    def get_objectparam_nothrow[T: IUnknownWrapper](self, key: str, t: type[T]) -> ComResult[T]:
        x = POINTER(IUnknown)()
        return cr(self.__o.RegisterObjectParam(key, byref(x)), t(x))

    def get_objectparam[T: IUnknownWrapper](self, key: str, t: type[T]) -> T:
        return self.get_objectparam_nothrow(key, t).value

    def get_enumobjectparams_nothrow(self) -> ComResult[ComStringEnumerator]:
        x = POINTER(IEnumString)()
        return cr(self.__o.EnumObjectParam(byref(x)), ComStringEnumerator(x))

    def get_enumobjectparams(self) -> ComStringEnumerator:
        return self.get_enumobjectparams_nothrow().value

    def revoke_objectparam_nothrow(self, key: str) -> ComResult[None]:
        return cr(self.__o.RevokeObjectParam(key), None)

    def revoke_objectparam(self, key: str) -> None:
        return self.revoke_objectparam_nothrow(key).value


_CreateBindCtx = _ole32.CreateBindCtx
_CreateBindCtx.argtypes = (c_uint32, POINTER(POINTER(IBindCtx)))
