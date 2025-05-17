"""シェルアイテム配列。

主なクラスは :class:`ShellItemArray` です。
"""

from ctypes import POINTER, _Pointer, byref, c_int32, c_uint32, c_void_p
from enum import IntFlag
from typing import TYPE_CHECKING, Any, Iterator, Sequence, overload

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, check_hresult, cr, query_interface
from powcpropsys.propdesc import IPropertyDescriptionList, PropertyDescriptionList
from powcpropsys.propkey import PropertyKey
from powcpropsys.propstore import GetPropertyStoreFlag, IPropertyStore, PropertyStore

from . import _ole32, _shell32
from .shellitem import IShellItem, ShellItem
from .shellitem2 import ShellItem2
from .shellitemenum import EnumShellItems, IEnumShellItems


class ShellItemAttributeFlag(IntFlag):
    """SIATTRIBFLAGS"""

    AND = 0x1
    OR = 0x2
    APPCOMPAT = 0x3
    ALLITEMS = 0x4000


class IShellItemArray(IUnknown):
    """"""

    _iid_ = GUID("{b63ea76d-1f85-456f-a19c-48159efa858b}")
    _methods_ = [
        STDMETHOD(
            c_int32, "BindToHandler", (POINTER(IUnknown), POINTER(GUID), POINTER(GUID), POINTER(POINTER(IUnknown)))
        ),
        STDMETHOD(c_int32, "GetPropertyStore", (c_int32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(
            c_int32, "GetPropertyDescriptionList", (POINTER(PropertyKey), POINTER(GUID), POINTER(POINTER(IUnknown)))
        ),
        STDMETHOD(c_int32, "GetAttributes", (c_int32, c_int32, POINTER(c_int32))),
        STDMETHOD(c_int32, "GetCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetItemAt", (c_uint32, POINTER(POINTER(IShellItem)))),
        STDMETHOD(c_int32, "EnumItems", (POINTER(IEnumShellItems),)),
    ]

    __slots__ = ()


class ShellItemArray:
    """
    シェル項目配列。IShellItemArrayインターフェイスのラッパーです。
    """

    __o: Any  # POINTER(IShellItemArray)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IShellItemArray)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    if TYPE_CHECKING:

        def bind_tohandler_nothrow[TIUnknown](
            self, bhid: GUID, type: type[TIUnknown]
        ) -> ComResult[_Pointer[TIUnknown]]:  # type: ignore
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            ...

        def bind_tohandler[TIUnknown](self, bhid: GUID, type: type[TIUnknown]) -> _Pointer[TIUnknown]:  # type: ignore
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            ...

    else:

        def bind_tohandler_nothrow[TIUnknown](self, bhid: GUID, type: type[TIUnknown]):
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            p = POINTER(type)()
            return cr(self.__o.BindToHandler(None, bhid, p._iid_, byref(p)), p)

        def bind_tohandler[TIUnknown](self, bhid: GUID, type: type[TIUnknown]):
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            return self.bind_to_handler_nothrow(bhid, type).value

    def get_propstore_nothrow(self, flags: GetPropertyStoreFlag) -> ComResult[PropertyStore]:
        """プロパティストアを取得します。"""
        p = POINTER(IPropertyStore)()
        return cr(self.__o.GetPropertyStore(int(flags), IPropertyStore._iid_, byref(p)), PropertyStore(p))

    def get_propstore(self, flags: GetPropertyStoreFlag) -> PropertyStore:
        """プロパティストアを取得します。"""
        return self.get_propstore_nothrow(flags).value

    def get_propdescs_nothrow(self, key: PropertyKey) -> ComResult[PropertyDescriptionList]:
        """プロパティの説明を取得します。"""
        p = POINTER(IPropertyDescriptionList)()
        return cr(
            self.__o.GetPropertyDescriptionList(key, IPropertyDescriptionList._iid_, byref(p)),
            PropertyDescriptionList(p),
        )

    def get_propdescs(self, key: PropertyKey) -> PropertyDescriptionList:
        """プロパティの説明を取得します。"""
        return self.get_propdescs_nothrow(key).value

    def get_attrs_nothrow(self, flags: ShellItemAttributeFlag, mask: int = 0xFFFFFFFF) -> ComResult[int]:
        """項目の属性を指定した方法で合成して返します。"""
        x = c_int32()
        return cr(self.__o.GetAttributes(int(flags), mask, byref(x)), x.value)

    def get_attrs(self, flags: ShellItemAttributeFlag, mask: int = 0xFFFFFFFF) -> int:
        """項目の属性を指定した方法で合成して返します。"""
        return self.get_attrs_nothrow(flags, mask).value

    @property
    def attrs_and_nothrow(self) -> ComResult[int]:
        """項目属性の論理積を返します。"""
        return self.get_attrs_nothrow(ShellItemAttributeFlag.AND)

    @property
    def attrs_and(self) -> int:
        """項目属性の論理積を返します。"""
        return self.get_attrs(ShellItemAttributeFlag.AND)

    @property
    def attrs_or_nothrow(self) -> ComResult[int]:
        """項目属性の論理和を返します。"""
        return self.get_attrs_nothrow(ShellItemAttributeFlag.OR)

    @property
    def attrs_or(self) -> int:
        """項目属性の論理和を返します。"""
        return self.get_attrs(ShellItemAttributeFlag.OR)

    def __len__(self) -> int:
        """項目数を返します。"""
        x = c_uint32()
        check_hresult(self.__o.GetCount(byref(x)))
        return x.value

    def getitem_at_nothrow(self, index: int) -> ComResult[ShellItem2]:
        """指定した位置の項目を返します。"""
        p = POINTER(IShellItem)()
        return cr(self.__o.GetItemAt(index, byref(p)), ShellItem2(p))

    def getitem_at(self, index: int) -> ShellItem2:
        """指定した位置の項目を返します。"""
        return self.getitem_at_nothrow(index).value

    def getitem_v1_at_nothrow(self, index: int) -> ComResult[ShellItem]:
        """指定した位置の項目をIShellItemインターフェイスで返します。"""
        p = POINTER(IShellItem)()
        return cr(self.__o.GetItemAt(index, byref(p)), ShellItem(p))

    def getitem_v1_at(self, index: int) -> ShellItem:
        """指定した位置の項目をIShellItemインターフェイスで返します。"""
        return self.getitem_v1_at_nothrow(index).value

    @overload
    def __getitem__(self, key: int) -> ShellItem2:
        """指定した位置の項目を返します。"""
        ...

    @overload
    def __getitem__(self, key: slice) -> tuple[ShellItem2, ...]:
        """指定した範囲の項目タプルを返します。"""
        ...

    @overload
    def __getitem__(self, key: tuple[slice, ...]) -> tuple[ShellItem2, ...]:
        """指定した範囲の項目タプルを返します。"""
        ...

    def __getitem__(self, key) -> ShellItem2 | tuple[ShellItem2, ...]:
        """指定した位置・範囲の項目・項目タプルを返します。"""
        if isinstance(key, slice):
            return tuple(self.__getitem__(i) for i in range(*key.indices(len(self))))
        if isinstance(key, tuple):
            for subslice in key:
                if not isinstance(subslice, slice):
                    raise TypeError
            return tuple(item for item in (t for t in self.__getitem__(key)))
        else:
            return self.getitem_at(key)

    def __iter__(self):
        """項目のイテレータを返します。"""
        _getitem_at = self.getitem_at
        return (_getitem_at(i) for i in range(len(self)))

    def iter_items(self) -> Iterator[ShellItem2]:
        """項目のイテレータを返します。"""
        _getitem_at = self.getitem_at
        return (_getitem_at(i) for i in range(len(self)))

    def iter_items_v1(self) -> Iterator[ShellItem]:
        """項目をIShellItemインターフェイスで返すイテレータを返します。"""
        _getitem_v1_at = self.getitem_v1_at
        return (_getitem_v1_at(i) for i in range(len(self)))

    def enum_items_nothrow(self) -> ComResult[EnumShellItems]:
        """列挙用のIEnumShellItemsインターフェイスを返します。"""
        p = POINTER(IEnumShellItems)()
        return cr(self.__o.EnumItems(byref(p)), EnumShellItems(p))

    def enum_items(self) -> EnumShellItems:
        """列挙用のIEnumShellItemsインターフェイスを返します。"""
        return self.enum_items_nothrow().value

    @staticmethod
    def create_nothrow(parent_pidl: int, child_pidls: Sequence[int]) -> "ComResult[ShellItemArray]":
        p = (c_void_p * len(child_pidls))()
        for i in range(len(child_pidls)):
            p[i] = child_pidls[i]
        o = POINTER(IShellItemArray)()
        return cr(_SHCreateShellItemArray(parent_pidl, None, len(child_pidls), p, byref(o)), ShellItemArray(o))

    @staticmethod
    def create(parent_pidl: int, child_pidls: Sequence[int]) -> "ShellItemArray":
        return ShellItemArray.create_nothrow(parent_pidl, child_pidls).value

    @staticmethod
    def create_fromdataobj_nothrow(dataobj: c_void_p) -> "ComResult[ShellItemArray]":
        o = POINTER(IShellItemArray)()
        return cr(_SHCreateShellItemArrayFromDataObject(dataobj, IShellItemArray._iid_, byref(o)), ShellItemArray(o))

    @staticmethod
    def create_fromdataobj(dataobj: c_void_p) -> "ShellItemArray":
        return ShellItemArray.create_fromdataobj_nothrow(dataobj).value

    @staticmethod
    def create_fromidlists_nothrow(pidls: Sequence[int]) -> "ComResult[ShellItemArray]":
        p = (c_void_p * len(pidls))()
        for i in range(len(pidls)):
            p[i] = pidls[i]
        o = POINTER(IShellItemArray)()
        return cr(_SHCreateShellItemArrayFromIDLists(len(pidls), p, byref(o)), ShellItemArray(o))

    @staticmethod
    def create_fromidlists(pidls: Sequence[int]) -> "ShellItemArray":
        return ShellItemArray.create_fromidlists_nothrow(pidls).value

    @staticmethod
    def create_fromitem_nothrow(item: ShellItem) -> "ComResult[ShellItemArray]":
        o = POINTER(IShellItemArray)()
        return cr(
            _SHCreateShellItemArrayFromShellItem(item.wrapped_obj, IShellItemArray._iid_, byref(o)), ShellItemArray(o)
        )

    @staticmethod
    def create_fromitem(item: ShellItem) -> "ShellItemArray":
        return ShellItemArray.create_fromitem_nothrow(item).value

    @staticmethod
    def create_fromclipboard_nothrow() -> "ComResult[ShellItemArray]":
        dataobj = POINTER(IUnknown)()
        ret = cr(_OleGetClipboard(byref(dataobj)), None)
        if not ret:
            return cr(ret.hr, ShellItemArray(None))
        return ShellItemArray.create_fromdataobj_nothrow(dataobj)  # type: ignore

    @staticmethod
    def create_fromclipboard() -> "ShellItemArray":
        """クリップボードからシェル項目配列ShellItemArrayを作成して返します。

        Examples:
            >>> tuple(ShellItemArray.create_fromclipboard().iter_items())
        """
        return ShellItemArray.create_fromclipboard_nothrow().value

    @staticmethod
    def getitems_fromclipboard_nothrow() -> ComResult[tuple[ShellItem2, ...]]:
        array = ShellItemArray.create_fromclipboard_nothrow()
        if not array:
            return cr(array.hr, ())
        return cr(array.hr, tuple(array.value_unchecked.iter_items()))

    @staticmethod
    def getitems_fromclipboard() -> tuple[ShellItem2, ...]:
        """クリップボードからシェル項目配列tuple[ShellItem2, ...]を作成して返します。

        Examples:
            >>> ShellItemArray.getitems_fromclipboard()
        """
        return tuple(ShellItemArray.create_fromclipboard().iter_items())

    @staticmethod
    def getitems_v1_fromclipboard_nothrow() -> ComResult[tuple[ShellItem, ...]]:
        array = ShellItemArray.create_fromclipboard_nothrow()
        if not array:
            return cr(array.hr, ())
        return cr(array.hr, tuple(array.value_unchecked.iter_items_v1()))

    @staticmethod
    def getitems_v1_fromclipboard() -> tuple[ShellItem, ...]:
        """クリップボードからシェル項目配列tuple[ShellItem2, ...]を作成して返します。

        Examples:
            >>> ShellItemArray.getitems_v1_fromclipboard()
        """
        return tuple(ShellItemArray.create_fromclipboard().iter_items_v1())


_SHCreateShellItemArray = _shell32.SHCreateShellItemArray
_SHCreateShellItemArray.argtypes = (c_void_p, c_void_p, c_uint32, POINTER(c_void_p), POINTER(POINTER(IShellItemArray)))
_SHCreateShellItemArray.restype = c_int32

_SHCreateShellItemArrayFromDataObject = _shell32.SHCreateShellItemArrayFromDataObject
_SHCreateShellItemArrayFromDataObject.argtypes = (POINTER(IUnknown), POINTER(GUID), POINTER(POINTER(IUnknown)))
_SHCreateShellItemArrayFromDataObject.restype = c_int32

_SHCreateShellItemArrayFromIDLists = _shell32.SHCreateShellItemArrayFromIDLists
_SHCreateShellItemArrayFromIDLists.argtypes = (c_uint32, POINTER(c_void_p), POINTER(POINTER(IShellItemArray)))
_SHCreateShellItemArrayFromIDLists.restype = c_int32

_SHCreateShellItemArrayFromShellItem = _shell32.SHCreateShellItemArrayFromShellItem
_SHCreateShellItemArrayFromShellItem.argtypes = (
    POINTER(IShellItem),
    POINTER(GUID),
    POINTER(POINTER(IShellItemArray)),
)
_SHCreateShellItemArrayFromShellItem.restype = c_int32

_OleGetClipboard = _ole32.OleGetClipboard
_OleGetClipboard.argtypes = (POINTER(POINTER(IUnknown)),)
_OleGetClipboard.restype = c_int32
_OleGetClipboard.restype = c_int32
