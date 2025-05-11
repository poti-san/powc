"""プロパティにアクセスしやすいシェル項目。

主なクラスは :class:`ShellItem2` です。
"""

from ctypes import POINTER, byref, c_int32, c_int64, c_uint32, c_uint64, c_wchar_p
from typing import Any, Iterator, override

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, cotaskmem, cr, query_interface
from powc.datetime import FILETIME
from powcpropsys.propkey import PropertyKey
from powcpropsys.propstore import GetPropertyStoreFlag, IPropertyStore, PropertyStore
from powcpropsys.propsys import IPropertyDescriptionList, PropertyDescriptionList
from powcpropsys.propvariant import PropVariant

from .shellitem import (
    BindHandlerID,
    IShellItem,
    ShellItem,
    _SHCreateItemFromIDList,
    _SHCreateItemFromParsingName,
    _SHCreateItemInKnownFolder,
)
from .shellitemenum import EnumShellItems, IEnumShellItems


class IShellItem2(IShellItem):
    """"""

    _iid_ = GUID("{7e9fb0d3-919f-4307-ab2e-9b1860310c93}")
    _methods_ = [
        STDMETHOD(c_int32, "GetPropertyStore", (c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(
            c_int32,
            "GetPropertyStoreWithCreateObject",
            (c_uint32, POINTER(IUnknown), POINTER(GUID), POINTER(POINTER(IUnknown))),
        ),
        STDMETHOD(
            c_int32,
            "GetPropertyStoreForKeys",
            (POINTER(PropertyKey), c_uint32, c_int32, POINTER(GUID), POINTER(POINTER(IUnknown))),
        ),
        STDMETHOD(
            c_int32, "GetPropertyDescriptionList", (POINTER(PropertyKey), POINTER(GUID), POINTER(POINTER(IUnknown)))
        ),
        STDMETHOD(c_int32, "Update", (POINTER(IUnknown),)),  # IBindCtx
        STDMETHOD(c_int32, "GetProperty", (POINTER(PropertyKey), POINTER(PropVariant))),
        STDMETHOD(c_int32, "GetCLSID", (POINTER(PropertyKey), POINTER(GUID))),
        STDMETHOD(c_int32, "GetFileTime", (POINTER(PropertyKey), POINTER(c_int64))),
        STDMETHOD(c_int32, "GetInt32", (POINTER(PropertyKey), POINTER(c_int32))),
        STDMETHOD(c_int32, "GetString", (POINTER(PropertyKey), POINTER(c_wchar_p))),
        STDMETHOD(c_int32, "GetUInt32", (POINTER(PropertyKey), POINTER(c_uint32))),
        STDMETHOD(c_int32, "GetUInt64", (POINTER(PropertyKey), POINTER(c_uint64))),
        STDMETHOD(c_int32, "GetBool", (POINTER(PropertyKey), POINTER(c_int32))),
    ]

    __slots__ = ()


class ShellItem2(ShellItem):
    """シェル項目。IShellItem2インターフェイスのラッパーです。

    Note:
        \\_\\_iter\\_\\_、parent等でIShellItemが必要な場合は以下のコードでShellItemを作成してください。

        >>> from powc.core import queryinterface
        >>> item = ShellItem(queryinterface(item2.wrapped_obj, IShellItem))
    """

    __o: Any  # IShellItem2

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        """IShellItem2インターフェイスをラップして初期化します。

        Args:
            o (Any): IShellItem2インターフェイスに変換可能なIUnknown派生インターフェイスのインスタンス。
        """
        self.__o = query_interface(o, IShellItem2)
        super().__init__(o)

    def __repr__(self) -> str:
        name = self.name_desktopabsparsing_nothrow
        return f"ShellItem2({name.value_unchecked})" or "" if name else super().__repr__()

    def __str__(self) -> str:
        name = self.name_desktopabsparsing_nothrow
        return f"ShellItem2({name.value_unchecked})" or "" if name else super().__str__()

    @staticmethod
    def create_knownfolder_nothrow(folder_id: GUID, flags: int = 0) -> "ComResult[ShellItem2]":
        """既知フォルダのシェルアイテムを作成します。

        Args:
            folder_id (GUID): KnownFolderID定数。
            flags (int, optional): フラグ. Defaults to 0.
        """
        p = POINTER(IShellItem2)()
        return cr(_SHCreateItemInKnownFolder(folder_id, flags, None, IShellItem2._iid_, byref(p)), ShellItem2(p))

    @staticmethod
    def create_knownfolder(folder_id: GUID, flags: int = 0) -> "ShellItem2":
        return ShellItem2.create_knownfolder_nothrow(folder_id, flags).value

    create_knownfolder.__doc__ = create_knownfolder_nothrow.__doc__

    @staticmethod
    def create_knownfolder_item_nothrow(folder_id: GUID, itemname: str, flags: int = 0) -> "ComResult[ShellItem2]":
        """既知フォルダ内のシェル項目を作成します。

        Args:
            folder_id (GUID): KnownFolderID定数。
            itemname (str): 項目の名前。
            flags (int, optional): フラグ。
        """
        p = POINTER(IShellItem2)()
        return cr(_SHCreateItemInKnownFolder(folder_id, flags, itemname, IShellItem2._iid_, byref(p)), ShellItem2(p))

    @staticmethod
    def create_knownfolder_item(folder_id: GUID, itemname: str, flags: int = 0) -> "ShellItem2":
        return ShellItem2.create_knownfolder_item_nothrow(folder_id, itemname, flags).value

    create_knownfolder_item.__doc__ = create_knownfolder_item_nothrow.__doc__

    @staticmethod
    def create_parsingname_nothrow(name: str) -> "ComResult[ShellItem2]":
        """解析名からシェルアイテムを作成します。"""
        p = POINTER(IShellItem2)()
        return cr(_SHCreateItemFromParsingName(name, None, IShellItem2._iid_, byref(p)), ShellItem2(p))

    @staticmethod
    def create_parsingname(name: str) -> "ShellItem2":
        """解析名からシェルアイテムを作成します。"""
        return ShellItem2.create_parsingname_nothrow(name).value

    @staticmethod
    def create_from_idlist_nothrow(pidl: int) -> "ComResult[ShellItem2]":
        """アイテムIDリストからシェルアイテムを作成します。"""
        p = POINTER(IShellItem2)()
        return cr(_SHCreateItemFromIDList(pidl, IShellItem2._iid_, byref(p)), ShellItem2(p))

    @staticmethod
    def create_from_idlist(pidl: int) -> "ShellItem2":
        """アイテムIDリストからシェルアイテムを作成します。"""
        return ShellItem2.create_from_idlist_nothrow(pidl).value

    def get_propstore_nothrow(self, flags: GetPropertyStoreFlag) -> ComResult[PropertyStore]:
        """プロパティストアを取得します。"""
        p = POINTER(IPropertyStore)()
        return cr(self.__o.GetPropertyStore(int(flags), byref(IPropertyStore._iid_), byref(p)), PropertyStore(p))

    def get_propstore(self, flags: GetPropertyStoreFlag) -> PropertyStore:
        """プロパティストアを取得します。"""
        return self.get_propstore_nothrow(flags).value

    #         STDMETHOD(
    #             c_int32,
    #             "GetPropertyStoreWithCreateObject",
    #             (c_uint32, POINTER(IUnknown), POINTER(GUID), POINTER(IUnknown)),
    #         ),
    #         STDMETHOD(
    #             c_int32,
    #             "GetPropertyStoreForKeys",
    #             (POINTER(PropertyKey), c_uint32, GETPROPERTYSTOREFLAGS, POINTER(GUID), POINTER(IUnknown)),
    #         ),

    def get_propdescs_nothrow(self, key: PropertyKey) -> ComResult[PropertyDescriptionList]:
        """プロパティの説明を取得します。"""
        x = POINTER(IPropertyDescriptionList)()
        return cr(
            self.__o.GetPropertyDescriptionList(key, IPropertyDescriptionList._iid_, byref(x)),
            PropertyDescriptionList(x),
        )

    def get_propdescs(self, key: PropertyKey) -> PropertyDescriptionList:
        """プロパティの説明を取得します。"""
        return self.get_propdescs_nothrow(key).value

    @property
    def propstore_nothrow(self) -> ComResult[PropertyStore]:
        """既定のプロパティストアを取得します。"""
        return self.get_propstore_nothrow(GetPropertyStoreFlag.DEFAULT)

    @property
    def propstore(self) -> PropertyStore:
        """既定のプロパティストアを取得します。"""
        return self.propstore_nothrow.value

    @property
    def propstore_with_slowitem_nothrow(self) -> ComResult[PropertyStore]:
        """遅い項目を含むプロパティストアを取得します。"""
        return self.get_propstore_nothrow(GetPropertyStoreFlag.OPEN_SLOW_ITEM)

    @property
    def propstore_with_slowitem(self) -> PropertyStore:
        """遅い項目を含むプロパティストアを取得します。"""
        return self.propstore_with_slowitem_nothrow.value

    def get_prop_nothrow(self, key: PropertyKey) -> ComResult[PropVariant]:
        """キーに対応するプロパティを取得します。"""
        value = PropVariant()
        return cr(self.__o.GetProperty(byref(key), byref(value)), value)

    def get_prop(self, key: PropertyKey) -> PropVariant:
        """キーに対応するプロパティを取得します。"""
        return self.get_prop_nothrow(key).value

    def get_clsid_nothrow(self, key: PropertyKey) -> ComResult[GUID]:
        """キーに対応するCLSIDを取得します。"""
        value = GUID()
        return cr(self.__o.GetCLSID(byref(key), byref(value)), value)

    def get_clsid(self, key: PropertyKey) -> GUID:
        """キーに対応するCLSIDを取得します。"""
        return self.get_clsid_nothrow(key).value

    def get_filetime_nothrow(self, key: PropertyKey) -> ComResult[FILETIME]:
        """キーに対応するFILETIMEを取得します。"""
        value = FILETIME()
        return cr(self.__o.GetFileTime(byref(key), byref(value)), value)

    def get_filetime(self, key: PropertyKey) -> FILETIME:
        """キーに対応するFILETIMEを取得します。"""
        return self.get_filetime_nothrow(key).value

    def get_int32_nothrow(self, key: PropertyKey) -> ComResult[int]:
        """キーに対応する32ビット符号付き整数を取得します。"""
        value = c_int32()
        return cr(self.__o.GetInt32(byref(key), byref(value)), value.value)

    def get_int32(self, key: PropertyKey) -> int:
        """キーに対応する32ビット符号付き整数を取得します。"""
        return self.get_int32_nothrow(key).value

    def get_str_nothrow(self, key: PropertyKey) -> ComResult[str]:
        """キーに対応する文字列を取得します。"""
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetInt32(byref(key), byref(p)), p.value or "")

    def get_str(self, key: PropertyKey) -> str:
        """キーに対応する文字列を取得します。"""
        return self.get_str_nothrow(key).value

    def get_uint32_nothrow(self, key: PropertyKey) -> ComResult[int]:
        """キーに対応する32ビット符号無し整数を取得します。"""
        value = c_uint32()
        return cr(self.__o.GetUInt32(byref(key), byref(value)), value.value)

    def get_uint32(self, key: PropertyKey) -> int:
        """キーに対応する32ビット符号無し整数を取得します。"""
        return self.get_uint32_nothrow(key).value

    def get_uint64_nothrow(self, key: PropertyKey) -> ComResult[int]:
        """キーに対応する64ビット符号無し整数を取得します。"""
        value = c_uint64()
        return cr(self.__o.GetUInt64(byref(key), byref(value)), value.value)

    def get_uint64(self, key: PropertyKey) -> int:
        """キーに対応する64ビット符号無し整数を取得します。"""
        return self.get_uint64_nothrow(key).value

    def get_bool_nothrow(self, key: PropertyKey) -> ComResult[bool]:
        """キーに対応する真偽値を取得します。"""
        value = c_int32()
        return cr(self.__o.GetBool(byref(key), byref(value)), value.value != 0)

    def get_bool(self, key: PropertyKey) -> bool:
        """キーに対応する真偽値を取得します。"""
        return self.get_bool_nothrow(key).value

    @property
    def parent_nothrow(self) -> "ComResult[ShellItem2]":
        """親フォルダを取得します。"""
        p = POINTER(IShellItem2)()
        return cr(self.__o.GetParent(byref(p)), ShellItem2(p))

    @property
    def parent(self) -> "ShellItem2":
        """親フォルダを取得します。"""
        return self.parent_nothrow.value

    def iter_items(self) -> "Iterator[ShellItem2]":
        """フォルダ内の項目を列挙します。"""
        penum = EnumShellItems(self.bind_to_handler(BindHandlerID.ENUM_ITEMS, IEnumShellItems))
        return (ShellItem2(o.QueryInterface(IShellItem2)) for o in penum)  # type: ignore

    @property
    def items(self) -> "tuple[ShellItem2,...]":
        """フォルダ内の項目を列挙します。"""
        return tuple(self.iter_items())

    @property
    @override
    def linktarget_nothrow(self) -> "ComResult[ShellItem2]":
        """項目がシェルリンクの場合にリンク先項目を取得します。"""
        o = self.bind_to_handler_nothrow(BindHandlerID.LINK_TARGET_ITEM, IShellItem2)
        return cr(o.hr, ShellItem2(o.value_unchecked))

    @property
    @override
    def linktarget(self) -> "ShellItem2":
        return self.linktarget_nothrow.value
        return self.linktarget_nothrow.value
