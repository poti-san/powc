"""アイテムIDリスト。

主なクラスは :class:`ItemIDList` です。
"""

from ctypes import POINTER, byref, c_int32, c_void_p
from typing import Any

from comtypes import IUnknown

from powc.core import check_hresult, cotaskmem_free

from . import _shell32


class ItemIDList(c_void_p):
    """アイテムIDリスト。"""

    __slots__ = ()

    def __del__(self):
        cotaskmem_free(self)

    @staticmethod
    def from_object(o: Any) -> "ItemIDList":
        x = ItemIDList()
        check_hresult(_SHGetIDListFromObject(o, byref(x)))
        return x


_SHGetIDListFromObject = _shell32.SHGetIDListFromObject
_SHGetIDListFromObject.argtypes = (POINTER(IUnknown), POINTER(ItemIDList))
_SHGetIDListFromObject.restype = c_int32
_SHGetIDListFromObject.restype = c_int32
