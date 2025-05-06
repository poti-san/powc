"""基本的なCOM機能。"""

from contextlib import contextmanager
from ctypes import POINTER, WinError, _Pointer, c_int32, c_size_t, c_void_p, memmove
from typing import TYPE_CHECKING, Any, Iterator, NoReturn, Protocol, runtime_checkable

from comtypes import GUID, IUnknown

from . import _ole32


class ComResult[T]:
    """COMメソッドまたは関数の結果を値とエラーコードとして保持します。最低限の機能のみ実装します。"""

    __hr: int
    __value: T

    __slots__ = ("__hr", "__value")

    def __init__(self, hr: int, value: T):
        self.__hr = hr
        self.__value = value

    def __repr__(self):
        return f"ComResult(hr:{self.__hr}, value:{repr(self.__value)})"

    def __str__(self):
        return f"ComResult(hr:{self.__hr}, value:{repr(self.__value)})"

    @property
    def hr(self):
        """エラーコード。"""
        return self.__hr

    @property
    def value_unchecked(self):
        """成否を判定せず値を返します。"""
        return self.__value

    @property
    def value(self):
        """成否を判定して値を返します。失敗時は例外を発生します。
        successや__bool__で有無を確認した場合はvalue_uncheckedの方が高速です。"""
        self.raise_if_error()
        return self.__value

    @property
    def success(self) -> bool:
        """成功時は真。"""
        return self.__hr >= 0

    def __bool__(self) -> bool:
        """成功時は真。"""
        return self.__hr >= 0

    def raise_always(self) -> NoReturn:
        """成否に関わらず例外を発生します。"""
        raise WinError(self.__hr)

    def raise_if_error(self) -> None:
        """失敗時のみ例外を発生します。"""
        if not self.success:
            self.raise_always()

    @property
    def value_or_none(self) -> T | None:
        return self.__value if self.success else None


def cr[T](hr: int, value: T) -> ComResult[T]:
    """ComResultクラスを作成する関数。記述の短縮に使用します。"""
    return ComResult[T](hr, value)


_CoTaskMemFree = _ole32.CoTaskMemFree
_CoTaskMemFree.argtypes = (c_void_p,)
_CoTaskMemFree.restype = None

_CoTaskMemAlloc = _ole32.CoTaskMemAlloc
_CoTaskMemAlloc.argtypes = (c_size_t,)
_CoTaskMemAlloc.restype = c_void_p


@contextmanager
def cotaskmem[T](p: T) -> Iterator[T]:
    """comtypes.c_void_p型等のCOMタスクメモリをスコープ脱出時に解放します。
    Examples:
        >>> from ctypes import c_void_p
        >>> with cotaskmem(c_void_p()) as p:
        >>>     # pのcomtypes.byrefによる確保・使用
        >>>     pass
    """
    try:
        yield p
    finally:
        _CoTaskMemFree(p)


def cotaskmem_alloc(size: int) -> c_void_p:
    """COMメモリを確保します。
    Examples:
        >>> with cotaskmem(cotaskmemalloc(10)) as p:
        >>>     # pの使用
        >>>     pass
    """
    return _CoTaskMemAlloc(size)


def cotaskmem_free(p: Any) -> None:
    """COMメモリを解放します。
    Examples:
        >>> p = cotaskmemalloc(10)
        >>> cotaskmemfree(p)
    """
    _CoTaskMemFree(p)


class CoTaskMem(c_void_p):
    """COMメモリのラッパーです。"""

    __slots__ = ()

    @staticmethod
    def alloc_bytes(bytes: int) -> "CoTaskMem":
        return CoTaskMem(_CoTaskMemAlloc(bytes))

    @staticmethod
    def alloc_unistr(s: str) -> "CoTaskMem":
        b = f"{s}\0".encode("utf-16le")
        p = CoTaskMem.alloc_bytes(len(b))
        memmove(p, b, len(b))
        return p

    def __del__(self) -> None:
        cotaskmem_free(self)

    def detatch(self) -> int:
        x = self.value
        self.value = None
        return x or 0


def hr(code: int) -> int:
    """Pythonのint型をWindowsのHRESULTに変換します。
    Windows用のHRESULT定数をそのまま貼り付ける場合に使用します。

    Examples:
        >>> print(f"{hr(0x887A0002):X}")  # -7785FFFE
    """
    return c_int32(code).value


def raise_hresult(hr: int) -> NoReturn:
    """COMエラーに対応する例外を発生します。
    Args:
        hr (int): COMエラーコード
    Raises:
        WinError: COMエラー。
    Examples:
        >>> from comtypes import hresult
        >>> raise_hresult(hresult.S_OK)
    """

    raise WinError(hr)


def check_hresult(hr: int) -> None:
    """COMエラーコードがエラーの場合に例外を発生します。
    Args:
        hr (int): COMエラーコード。0x80000000が含まれる場合はエラーです。
    Raises:
        WinError: COMエラー。
    Examples:
        >>> from comtypes import hresult
        >>> check_hresult(hresult.S_OK)
        >>> check_hresult(hresult.S_FALSE)
        >>> check_hresult(hresult.E_FAIL)
    """
    if hr < 0:
        raise WinError(hr)


if TYPE_CHECKING:

    IUnknownPointer = _Pointer[IUnknown]

    def query_interface[TIUnknown: IUnknown](o: Any, interface_type: type[TIUnknown]) -> _Pointer[TIUnknown]:
        """comtypes.IUnknown派生インターフェイスを変換して返します。
        Raises:
            TypeError: oがPOINTER(comtypes.IUnknown)またはPOINTER(comtypes.IUnknown派生クラス)ではない
        Returns:
            _type_: 変換後のIUnknown派生インターフェイスインスタンス。
        """
        ...

else:

    IUnknownPointer = POINTER(IUnknown)

    def query_interface[TIUnknown: IUnknown](o: Any, interface_type: type[TIUnknown]) -> IUnknownPointer:
        """comtypes.IUnknown派生インターフェイスを変換して返します。
        Raises:
            TypeError: oがPOINTER(comtypes.IUnknown)またはPOINTER(comtypes.IUnknown派生クラス)ではない
        Returns:
            _type_: 変換後のIUnknown派生インターフェイスインスタンス。
        """
        if not isinstance(o, POINTER(IUnknown)):
            raise TypeError
        return o.QueryInterface(interface_type) if o else interface_type()


def guid_from_define(a: int, b: int, c: int, d: int, e: int, f: int, g: int, h: int, i: int, j: int, k: int) -> GUID:
    """C++のGUID構造体形式でGUIDを作成します。

    Examples:
        >>> guid_from_define(0xFBF23B40, 0xE3F0, 0x101B, 0x84, 0x88, 0x00, 0xAA, 0x00, 0x3E, 0x56, 0xF8)
    """
    guid = GUID()
    guid.Data1 = a
    guid.Data2 = b
    guid.Data3 = c
    guid.Data4 = (d, e, f, g, h, i, j, k)
    return guid


@runtime_checkable
class IUnknownWrapper(Protocol):
    def __init__(self, o: Any) -> None: ...

    @property
    def wrapped_obj(self) -> c_void_p: ...
