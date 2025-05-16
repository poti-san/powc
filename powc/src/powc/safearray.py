"""SAFEARRAY機能を提供します。"""

from contextlib import contextmanager
from ctypes import (
    POINTER,
    Array,
    Structure,
    byref,
    c_byte,
    c_int16,
    c_int32,
    c_uint16,
    c_uint32,
    c_void_p,
)
from typing import Iterator, Sequence

from comtypes import BSTR

from powc.core import ComResult, check_hresult, cr

from . import _oleaut32
from .variant import VARENUM


class _SAFEARRAYBOUND(Structure):
    _fields_ = (
        ("elements", c_uint32),
        ("lbound", c_int32),
    )

    __slots__ = ()


class SafeArrayPtr(c_void_p):
    __slots__ = ()

    @staticmethod
    def create_array(vt: VARENUM, elements: Sequence[int], lbounds: Sequence[int] | None = None) -> "SafeArrayPtr":
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
        return SafeArrayPtr(_SafeArrayCreate(vt, len(elements), pbounds))

    @staticmethod
    def create_vector(vt: VARENUM, elements: int, lbound: int = 0) -> "SafeArrayPtr":
        return SafeArrayPtr(_SafeArrayCreateVector(vt, lbound, elements))

    def __del__(self):
        _SafeArrayDestroy(self)

    def clear_data_nothrow(self) -> ComResult[None]:
        return cr(_SafeArrayDestroyData(self), None)

    def clear_data(self) -> None:
        return self.clear_data_nothrow().value

    def lock_nothrow(self) -> ComResult[None]:
        return cr(_SafeArrayLock(self), None)

    def lock(self):
        return self.lock_nothrow().value

    def unlock_nothrow(self) -> ComResult[None]:
        return cr(_SafeArrayUnlock(self), None)

    def unlock(self) -> None:
        return self.unlock_nothrow().value

    @contextmanager
    def lock_scope(self) -> "Iterator[SafeArrayPtr]":
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
        x = SafeArrayPtr()
        return cr(_SafeArrayCopy(self, byref(x)), x)

    def clone(self) -> "SafeArrayPtr":
        return self.clone_nothrow().value

    @property
    def dim(self) -> int:
        return _SafeArrayGetDim(self)

    @property
    def indices(self) -> tuple[int, ...]:
        return tuple(1 + i for i in range(self.dim))

    @property
    def elemsize(self) -> int:
        return _SafeArrayGetElemsize(self)

    def get_elem_at_nothrow(self, indexes: Sequence[int]) -> ComResult[Array[c_byte]]:
        x = (c_byte * self.elemsize)()
        a = (c_uint32 * len(indexes))(indexes)
        return cr(_SafeArrayGetElement(self, a, x), x)

    def get_elem_at(self, indexes: Sequence[int]) -> Array[c_byte]:
        return self.get_elem_at_nothrow(indexes).value

    def get_lbound_nothrow(self, dim: int) -> ComResult[int]:
        x = c_int32()
        return cr(_SafeArrayGetLBound(self, dim, byref(x)), x.value)

    def get_lbound(self, dim: int) -> int:
        return self.get_lbound_nothrow(dim).value

    def get_ubound_nothrow(self, dim: int) -> ComResult[int]:
        x = c_int32()
        return cr(_SafeArrayGetUBound(self, dim, byref(x)), x.value)

    def get_ubound(self, dim: int) -> int:
        return self.get_ubound_nothrow(dim).value

    @property
    def bounds(self) -> tuple[tuple[int, int], ...]:
        return tuple((self.get_lbound(1 + i), self.get_ubound(1 + i)) for i in range(self.dim))

    @property
    def totallen(self) -> int:
        return sum(u - l for l, u in self.bounds)  # noqa E741

    @property
    def totalsize(self) -> int:
        return self.totallen * self.elemsize

    @property
    def vartype_nothrow(self) -> ComResult[VARENUM]:
        x = c_int16()
        return cr(_SafeArrayGetVartype(self, byref(x)), VARENUM(x.value))

    @property
    def vartype(self) -> VARENUM:
        return self.vartype_nothrow.value

    #
    # 型別ユーティリティ
    #

    def to_bstrarray(self) -> tuple[str, ...]:
        with self.access_data() as data:
            return tuple((BSTR * self.totallen).from_buffer(data))

    def to_int32array(self) -> tuple[int, ...]:
        with self.access_data() as data:
            return tuple((c_int32 * self.totallen).from_buffer(data))

    def to_uint32array(self) -> tuple[int, ...]:
        with self.access_data() as data:
            return tuple((c_uint32 * self.totallen).from_buffer(data))


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
