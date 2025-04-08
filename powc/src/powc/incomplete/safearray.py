"""
SAFEARRAY関係の機能を提供します。未完成です。
"""

# NOTE: 未完成

from ctypes import POINTER, Structure, c_int32, c_uint32, c_void_p
from typing import Sequence

from powc.core import check_hresult

from .. import _oleaut32
from ..variant import VARENUM


class SafeArrayBound(Structure):
    _fields_ = (
        ("elements", c_uint32),
        ("lbound", c_int32),
    )


class SafeArray:
    """このクラスは未完成です。"""

    __p: int

    __slots__ = ("__p",)

    def __init__(self, p):
        self.__p = p

    @staticmethod
    def create(vt: VARENUM, elements: Sequence[int]) -> "SafeArray":
        pbounds = (SafeArrayBound * len(elements))()
        for i in range(len(elements)):
            pbounds[i].elements = elements[i]
            pbounds[i].lbound = 0
        return SafeArray(_SafeArrayCreate(vt, len(elements), pbounds))

    @staticmethod
    def create_with_lbound(vt: VARENUM, bounds: Sequence[SafeArrayBound]) -> "SafeArray":
        pbounds = (SafeArrayBound * len(bounds))()
        for i in range(len(bounds)):
            pbounds[i] = bounds[i]
        return SafeArray(_SafeArrayCreate(vt, len(bounds), pbounds))

    def __del__(self):
        _SafeArrayDestroy(self.__p)

    def clear_data(self):
        check_hresult(_SafeArrayDestroyData(self.__p))

    def lock(self):
        check_hresult(_SafeArrayLock(self.__p))

    def unlock(self):
        check_hresult(_SafeArrayUnlock(self.__p))


_SafeArrayCreate = _oleaut32.SafeArrayCreate
_SafeArrayCreate.argtypes = (c_int32, c_uint32, POINTER(SafeArrayBound))
_SafeArrayCreate.restype = c_void_p

_SafeArrayDestroy = _oleaut32.SafeArrayDestroy
_SafeArrayDestroy.argtypes = (c_void_p,)
_SafeArrayDestroy.restype = c_int32

_SafeArrayDestroyData = _oleaut32.SafeArrayDestroyData
_SafeArrayDestroyData.argtypes = (c_void_p,)
_SafeArrayDestroyData.restype = c_int32

_SafeArrayLock = _oleaut32.SafeArrayLock
_SafeArrayLock.argtypes = (c_void_p,)
_SafeArrayLock.restype = c_int32

_SafeArrayUnlock = _oleaut32.SafeArrayUnlock
_SafeArrayUnlock.argtypes = (c_void_p,)
_SafeArrayUnlock.restype = c_int32
