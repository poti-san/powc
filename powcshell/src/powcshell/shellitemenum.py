"""シェルアイテムの列挙。

主なクラスは :class:`EnumShellItems` です。 :class:`~.ShellItem` や :class:`~.ShellItem2` で使用されます。
"""

from ctypes import POINTER, _Pointer, byref, c_int32, c_uint32, c_void_p
from typing import TYPE_CHECKING, Any

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, check_hresult, cr, query_interface

from .shellitem import IShellItem

if TYPE_CHECKING:
    IShellItemPointer = _Pointer[IShellItem]
else:
    IShellItemPointer = POINTER(IShellItem)


class IEnumShellItems(IUnknown):
    """"""

    _iid_ = GUID("{70629033-e363-4a28-a567-0db78006e6d7}")

    __slots__ = ()


IEnumShellItems._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(POINTER(IShellItem)), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IEnumShellItems)),)),
]


class EnumShellItems:
    """IEnumShellItemsインターフェイスのラッパーです。IShellItemを列挙します。"""

    __o: Any  # IEnumShellItems

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumShellItems)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def clone_nothrow(self) -> "ComResult[EnumShellItems]":
        p = POINTER(IEnumShellItems)()
        return cr(self.__o.Clone(byref(p)), EnumShellItems(p))

    def clone(self) -> "EnumShellItems":
        return self.clone_nothrow().value

    def __iter__(self) -> "EnumShellItems":
        return self

    def __next__(self) -> IShellItemPointer:
        x = POINTER(IShellItem)()
        hr = self.__o.Next(1, byref(x), None)
        if hr == 1:
            raise StopIteration()
        check_hresult(hr)
        return x
        return x
