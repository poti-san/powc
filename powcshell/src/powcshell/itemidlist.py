"""アイテムIDリスト。"""

from ctypes import POINTER, byref, c_int32, c_void_p, windll
from typing import Any

from comtypes import IUnknown

from powc.core import check_hresult, cotaskmemfree


class ItemIDList(c_void_p):
    """アイテムIDリスト。"""

    __slots__ = ()

    def __del__(self):
        cotaskmemfree(self)

    @staticmethod
    def from_object(o: Any) -> "ItemIDList":
        x = ItemIDList()
        check_hresult(_SHGetIDListFromObject(o, byref(x)))
        return x


_SHGetIDListFromObject = windll.shell32.SHGetIDListFromObject
_SHGetIDListFromObject.argtypes = (POINTER(IUnknown), POINTER(ItemIDList))
_SHGetIDListFromObject.restype = c_int32
