"""SAFEARRAY機能を提供します。"""

from contextlib import contextmanager
from ctypes import (
    POINTER,
    ArgumentError,
    Array,
    Structure,
    byref,
    c_byte,
    c_char_p,
    c_double,
    c_float,
    c_int,
    c_int8,
    c_int16,
    c_int32,
    c_int64,
    c_size_t,
    c_ssize_t,
    c_uint,
    c_uint8,
    c_uint16,
    c_uint32,
    c_uint64,
    c_void_p,
    c_wchar_p,
)
from typing import Any, Iterator, Sequence

from comtypes import BSTR, GUID

from powc.datetime import FILETIME

from . import _oleaut32
from .core import ComResult, check_hresult, cr
from .variant import VARENUM


class _SAFEARRAYBOUND(Structure):
    _fields_ = (
        ("elements", c_uint32),
        ("lbound", c_int32),
    )

    __slots__ = ()


class SafeArrayPtr(c_void_p):
    """セーフ配列の管理機能を提供します。データは :code:`SAFEARRAY*` として管理します。"""

    __slots__ = ()

    @staticmethod
    def create_array(vt: VARENUM, elements: Sequence[int], lbounds: Sequence[int] | None = None) -> "SafeArrayPtr":
        """新しい多次元セーフ配列を作成します。

        Args:
            vt (VARENUM): 要素の型。
            elements (Sequence[int]): 各次元の要素数。
            lbounds (Sequence[int] | None, optional): 各次元のインデックス下限値。既定値は全て0です。

        Raises:
            ValueError: 要素と下限のサイズが一致しない。
        """
        if lbounds is None:
            pbounds = (_SAFEARRAYBOUND * len(elements))()
            for i in range(len(elements)):
                pbounds[i].elements = elements[i]
                pbounds[i].lbound = 0
        else:
            if len(elements) != len(lbounds):
                raise ValueError("要素と下限のサイズが一致しません。")
            pbounds = (_SAFEARRAYBOUND * len(elements))()
            for i in range(len(elements)):
                pbounds[i].elements = elements[i]
                pbounds[i].lbound = lbounds[i]
        return _SafeArrayCreate(vt, len(elements), pbounds)

    @staticmethod
    def create_vector(vt: VARENUM, elements: int, lbound: int = 0) -> "SafeArrayPtr":
        """新しい1次元セーフ配列を作成します。

        Args:
            vt (VARENUM): 要素の型。
            elements (int): 要素数。
            lbound (int, optional): 要素インデックスの下限値。既定値は0です。
        """
        return _SafeArrayCreateVector(vt, lbound, elements)

    def __del__(self):
        """セーフ配列を解放します。"""
        _SafeArrayDestroy(self)

    def clear_data_nothrow(self) -> ComResult[None]:
        """セーフ配列の各要素を解放します。"""
        return cr(_SafeArrayDestroyData(self), None)

    def clear_data(self) -> None:
        """セーフ配列の各要素を解放します。"""
        return self.clear_data_nothrow().value

    def lock_nothrow(self) -> ComResult[None]:
        """セーフ配列のロック数を加算します。使用後は :meth:`unlock` を呼び出してください"""
        return cr(_SafeArrayLock(self), None)

    def lock(self):
        """セーフ配列のロック数を加算します。使用後は :meth:`unlock` を呼び出してください"""
        return self.lock_nothrow().value

    def unlock_nothrow(self) -> ComResult[None]:
        """セーフ配列のロック数を減算します。通常は :meth:`lock` の後で呼び出します。"""
        return cr(_SafeArrayUnlock(self), None)

    def unlock(self) -> None:
        """セーフ配列のロック数を減算します。通常は :meth:`lock` の後で呼び出します。"""
        return self.unlock_nothrow().value

    @contextmanager
    def lock_scope(self) -> "Iterator[SafeArrayPtr]":
        """セーフ配列のロックスコープを作成します。通常は :code:`with` と組み合わせて使用します。"""
        lock_succeeded = False
        try:
            self.lock()
            lock_succeeded = True
            yield self
        finally:
            if lock_succeeded:
                self.unlock()

    @contextmanager
    def access_data(self) -> Iterator[Array[c_byte]]:
        """セーフ配列のメモリにアクセスします。セーフ配列は列挙終了までロックされます。"""
        needs_unaccess = False
        try:
            x = c_void_p()
            check_hresult(_SafeArrayAccessData(self, byref(x)))
            needs_unaccess = True
            if x.value is None:
                raise ValueError
            yield (c_byte * self.totalsize).from_address(x.value)
        finally:
            if needs_unaccess:
                check_hresult(_SafeArrayUnaccessData(self))

    @contextmanager
    def access_data_mv(self) -> Iterator[memoryview]:
        """セーフ配列のメモリにアクセスします。セーフ配列は列挙終了までロックされます。"""
        needs_unaccess = False
        try:
            x = c_void_p()
            check_hresult(_SafeArrayAccessData(self, byref(x)))
            needs_unaccess = True
            if x.value is None:
                raise ValueError
            yield memoryview((c_byte * self.totalsize).from_address(x.value)).cast("B")
        finally:
            if needs_unaccess:
                check_hresult(_SafeArrayUnaccessData(self))

    def clone_nothrow(self) -> "ComResult[SafeArrayPtr]":
        """セーフ配列の複製を取得します。"""

        x = SafeArrayPtr()
        return cr(_SafeArrayCopy(self, byref(x)), x)

    def clone(self) -> "SafeArrayPtr":
        """セーフ配列の複製を取得します。"""
        return self.clone_nothrow().value

    @property
    def dim(self) -> int:
        """セーフ配列の次元数を取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.dim)
                # 3
        """
        return _SafeArrayGetDim(self)

    @property
    def indices(self) -> tuple[int, ...]:
        """セーフ配列の次元数のインデックスを取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.indices)
                # (0, 1, 2)
        """
        return tuple(i for i in range(self.dim))

    @property
    def elemsize(self) -> int:
        """セーフ配列の要素のバイト数を取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.elemsize)
                # 4 (VT_I4のバイト数)
        """
        return _SafeArrayGetElemsize(self)

    def get_elemraw_at_nothrow(self, indexes: Sequence[int]) -> ComResult[Array[c_byte]]:
        """各次元のインデックスを指定して要素のバイナリ表現を取得します。"""
        x = (c_byte * self.elemsize)()
        a = (c_uint32 * len(indexes))(indexes)
        return cr(_SafeArrayGetElement(self, a, x), x)

    def get_elemraw_at(self, indexes: Sequence[int]) -> Array[c_byte]:
        return self.get_elemraw_at_nothrow(indexes).value

    get_elemraw_at.__doc__ = get_elemraw_at_nothrow.__doc__

    def get_elem_at(self, indexes: Sequence[int], convertsNotSupportedTypeToBytes: bool = True) -> Any:
        """各次元のインデックスを指定して要素を取得します。

        Args:
            indexes (Sequence[int]): 各次元のインデックス。
            convertsNotSupportedTypeToBytes (bool, optional):
                未対応の型をバイナリ表現に変換するか。偽の場合、未対応の型に対して例外を発生します。既定値は真です。
        """
        raw = self.get_elemraw_at(indexes)
        match self.vartype:
            case VARENUM.VT_EMPTY:
                return object()
            case VARENUM.VT_NULL:
                return None
            case VARENUM.VT_I2:
                return c_int16.from_buffer(raw).value
            case VARENUM.VT_I4:
                return c_int32.from_buffer(raw).value
            case VARENUM.VT_R4:
                return c_float.from_buffer(raw).value
            case VARENUM.VT_R8:
                return c_double.from_buffer(raw).value
            # case VARENUM.VT_CY:
            # case VARENUM.VT_DATE:
            case VARENUM.VT_BSTR:
                return BSTR.from_buffer(raw).value
            # case VARENUM.VT_DISPATCH:
            case VARENUM.VT_ERROR:
                return c_int32.from_buffer(raw).value
            case VARENUM.VT_BOOL:
                return c_int32.from_buffer(raw).value != 0
            # case VARENUM.VT_VARIANT:
            # case VARENUM.VT_UNKNOWN:
            # case VARENUM.VT_DECIMAL:
            case VARENUM.VT_I1:
                return c_int8.from_buffer(raw).value
            case VARENUM.VT_UI1:
                return c_uint8.from_buffer(raw).value
            case VARENUM.VT_UI2:
                return c_uint16.from_buffer(raw).value
            case VARENUM.VT_UI4:
                return c_uint32.from_buffer(raw).value
            case VARENUM.VT_I8:
                return c_int64.from_buffer(raw).value
            case VARENUM.VT_UI8:
                return c_uint64.from_buffer(raw).value
            case VARENUM.VT_INT:
                return c_int.from_buffer(raw).value
            case VARENUM.VT_UINT:
                return c_uint.from_buffer(raw).value
            # case VARENUM.VT_VOID:
            case VARENUM.VT_HRESULT:
                return c_int32.from_buffer(raw).value
            case VARENUM.VT_PTR:
                return c_void_p.from_buffer(raw).value
            # case VARENUM.VT_SAFEARRAY:
            # case VARENUM.VT_CARRAY:
            case VARENUM.VT_USERDEFINED:
                return raw
            case VARENUM.VT_LPSTR:
                # TODO check null
                return c_char_p.from_buffer(raw).value
            case VARENUM.VT_LPWSTR:
                # TODO check null
                return c_wchar_p.from_buffer(raw).value
            # case VARENUM.VT_RECORD:
            case VARENUM.VT_INT_PTR:
                return c_ssize_t.from_buffer(raw).value
            case VARENUM.VT_UINT_PTR:
                return c_size_t.from_buffer(raw).value
            case VARENUM.VT_FILETIME:
                return FILETIME.from_buffer(raw).value
            case VARENUM.VT_BLOB:
                return raw
            # case VARENUM.VT_STREAM:
            # case VARENUM.VT_STORAGE:
            # case VARENUM.VT_STREAMED_OBJECT:
            # case VARENUM.VT_STORED_OBJECT:
            # case VARENUM.VT_BLOB_OBJECT:
            # case VARENUM.VT_CF:
            case VARENUM.VT_CLSID:
                return GUID.from_buffer_copy(raw)
            # case VARENUM.VT_VERSIONED_STREAM:
            # case VARENUM.VT_BSTR_BLOB:
            # case VARENUM.VT_BYREF:
            case _:
                if not convertsNotSupportedTypeToBytes:
                    raise ValueError
                return raw

    def __getitem__(self, indices: int | Sequence[int], convertsNotSupportedTypeToBytes: bool = True) -> Any:
        if isinstance(indices, int):
            return self.get_elem_at([indices], convertsNotSupportedTypeToBytes)
        elif indices is Sequence[int]:
            return self.get_elem_at(indices, convertsNotSupportedTypeToBytes)
        else:
            raise ArgumentError

    def get_lbound_nothrow(self, dim: int) -> ComResult[int]:
        """指定した次元のインデックス下限を取得します。"""
        x = c_int32()
        return cr(_SafeArrayGetLBound(self, dim, byref(x)), x.value)

    def get_lbound(self, dim: int) -> int:
        return self.get_lbound_nothrow(dim).value

    get_lbound.__doc__ = get_lbound_nothrow.__doc__

    def get_ubound_nothrow(self, dim: int) -> ComResult[int]:
        """指定した次元のインデックス上限を取得します。"""
        x = c_int32()
        return cr(_SafeArrayGetUBound(self, dim, byref(x)), x.value)

    def get_ubound(self, dim: int) -> int:
        return self.get_ubound_nothrow(dim).value

    get_ubound.__doc__ = get_ubound_nothrow.__doc__

    @property
    def bounds(self) -> tuple[tuple[int, int], ...]:
        """セーフ配列の各次元の範囲を取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.bounds)
                # ((0, 2), (0, 2), (0, 4))
        """
        return tuple((self.get_lbound(1 + i), self.get_ubound(1 + i)) for i in range(self.dim))

    @property
    def totallen(self) -> int:
        """セーフ配列の全要素数を取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.totallen)
                # 11
        """
        return sum(u - l + 1 for l, u in self.bounds)  # noqa E741

    @property
    def totalsize(self) -> int:
        """セーフ配列の全要素のバイト数を取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.totalsize)
                # 44
        """
        return self.totallen * self.elemsize

    @property
    def vartype_nothrow(self) -> ComResult[VARENUM]:
        """セーフ配列の全要素の型を取得します。

        Examples:
            .. code-block:: python

                from powc.safearray import SafeArrayPtr
                from powc.variant import VARENUM

                p = SafeArrayPtr.create_array(VARENUM.VT_I4, (3, 3, 5))
                print(p.vartype)
                # 3 (VT_I4)
        """
        x = c_int16()
        return cr(_SafeArrayGetVartype(self, byref(x)), VARENUM(x.value))

    @property
    def vartype(self) -> VARENUM:
        return self.vartype_nothrow.value

    vartype.__doc__ = vartype_nothrow.__doc__

    #
    # 型別ユーティリティ
    #

    def to_bstrarray(self) -> tuple[str, ...]:
        """セーフ配列を :code:`BSTR` 型とみなした配列を作成します。"""
        with self.access_data() as data:
            return tuple((BSTR * self.totallen).from_buffer(data))

    def to_int32array(self) -> tuple[int, ...]:
        """セーフ配列を32ビット符号付き整数とみなした配列を作成します。"""
        with self.access_data() as data:
            return tuple((c_int32 * self.totallen).from_buffer(data))

    def to_uint32array(self) -> tuple[int, ...]:
        """セーフ配列を32ビット符号無し整数とみなした配列を作成します。"""
        with self.access_data() as data:
            return tuple((c_uint32 * self.totallen).from_buffer(data))

    @staticmethod
    def __copy_elems(destination: Array, source: Sequence) -> None:
        for i in range(len(source)):
            destination[i] = source[i]

    @staticmethod
    def create_vector_int8(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_int8` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_I1, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_int8 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_int16(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_int16` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_I2, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_int16 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_int32(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_int32` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_I4, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_int32 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_int64(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_int64` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_I8, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_int64 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_uint8(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_uint8` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_UI1, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_uint8 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_uint16(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_uint16` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_UI2, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_uint16 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_uint32(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_uint32` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_UI4, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_uint32 * len(source)).from_buffer(data), source)
        return v

    @staticmethod
    def create_vector_uint64(source: Sequence[int]) -> "SafeArrayPtr":
        """:code:`ctypes.c_uint64` の1次元セーフ配列を作成します。"""
        v = SafeArrayPtr.create_vector(VARENUM.VT_UI8, len(source))
        with v.access_data() as data:
            SafeArrayPtr.__copy_elems((c_uint64 * len(source)).from_buffer(data), source)
        return v


_SafeArrayCreate = _oleaut32.SafeArrayCreate
_SafeArrayCreate.restype = SafeArrayPtr
_SafeArrayCreate.argtypes = (c_int32, c_uint32, POINTER(_SAFEARRAYBOUND))

_SafeArrayDestroy = _oleaut32.SafeArrayDestroy
_SafeArrayDestroy.argtypes = (c_void_p,)

_SafeArrayDestroyData = _oleaut32.SafeArrayDestroyData
_SafeArrayDestroyData.argtypes = (c_void_p,)

_SafeArrayLock = _oleaut32.SafeArrayLock
_SafeArrayLock.argtypes = (c_void_p,)

_SafeArrayUnlock = _oleaut32.SafeArrayUnlock
_SafeArrayUnlock.argtypes = (c_void_p,)

_SafeArrayAccessData = _oleaut32.SafeArrayAccessData
_SafeArrayAccessData.argtypes = (c_void_p, POINTER(c_void_p))

_SafeArrayUnaccessData = _oleaut32.SafeArrayUnaccessData
_SafeArrayUnaccessData.argtypes = (c_void_p,)

_SafeArrayCopy = _oleaut32.SafeArrayCopy
_SafeArrayCopy.argtypes = (c_void_p, POINTER(SafeArrayPtr))

_SafeArrayCreateVector = _oleaut32.SafeArrayCreateVector
_SafeArrayCreateVector.restype = SafeArrayPtr
_SafeArrayCreateVector.argtypes = (c_uint16, c_int32, c_uint32)

_SafeArrayGetDim = _oleaut32.SafeArrayGetDim
_SafeArrayGetDim.restype = c_uint32
_SafeArrayGetDim.argtypes = (c_void_p,)

_SafeArrayGetElement = _oleaut32.SafeArrayGetElement
_SafeArrayGetElement.argtypes = (c_void_p, POINTER(c_uint32), POINTER(c_void_p))

_SafeArrayGetElemsize = _oleaut32.SafeArrayGetElemsize
_SafeArrayGetElemsize.restype = c_uint32
_SafeArrayGetElemsize.argtypes = (c_void_p,)

_SafeArrayGetLBound = _oleaut32.SafeArrayGetLBound
_SafeArrayGetLBound.argtypes = (c_void_p, c_uint32, POINTER(c_int32))

_SafeArrayGetUBound = _oleaut32.SafeArrayGetUBound
_SafeArrayGetUBound.argtypes = (c_void_p, c_uint32, POINTER(c_int32))

_SafeArrayGetVartype = _oleaut32.SafeArrayGetVartype
_SafeArrayGetVartype.argtypes = (c_void_p, POINTER(c_int16))
_SafeArrayGetVartype.argtypes = (c_void_p, POINTER(c_int16))
