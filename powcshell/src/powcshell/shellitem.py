"""シェル項目。

主なクラスは :class:`ShellItem` です。
"""

from ctypes import POINTER, _Pointer, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from enum import IntEnum, IntFlag
from typing import TYPE_CHECKING, Any, Iterator

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, cotaskmem, cr, query_interface
from powc.stream import ComStream, IStream

from . import _shell32
from .itemidlist import ItemIDList
from .shellexec import ShellExecute, ShellExecuteOption, ShowCommand


class ShellItemCompareHint(IntEnum):
    DISPLAY = 0
    ALLFIELDS = 0x80000000
    CANONICAL = 0x10000000
    TEST_FILESYSPATH_IF_NOT_EQUAL = 0x20000000


class ShellItemAttribute(IntFlag):
    CANCOPY = 0x1
    CANMOVE = 0x2
    CANLINK = 0x4
    STORAGE = 0x00000008
    CANRENAME = 0x00000010
    CANDELETE = 0x00000020
    HASPROPSHEET = 0x00000040
    DROPTARGET = 0x00000100
    PLACEHOLDER = 0x00000800
    SYSTEM = 0x00001000
    ENCRYPTED = 0x00002000
    ISSLOW = 0x00004000
    GHOSTED = 0x00008000
    LINK = 0x00010000
    SHARE = 0x00020000
    READONLY = 0x00040000
    HIDDEN = 0x00080000
    FILESYSANCESTOR = 0x10000000
    FOLDER = 0x20000000
    FILESYSTEM = 0x40000000
    HASSUBFOLDER = 0x80000000
    VALIDATE = 0x01000000
    REMOVABLE = 0x02000000
    COMPRESSED = 0x04000000
    BROWSABLE = 0x08000000
    NONENUMERATED = 0x00100000
    NEWCONTENT = 0x00200000
    CANMONIKER = 0x00400000
    HASSTORAGE = 0x00400000
    STREAM = 0x00400000
    STORAGEANCESTOR = 0x00800000

    # CAPABILITYMASK = 0x00000177
    # DISPLAYATTRMASK = 0x000FC000
    # CONTENTSMASK = 0x80000000
    # STORAGECAPMASK = 0x70C50008
    # PKEYSFGAOMASK = 0x81044000


class BindHandlerID:
    SF_OBJECT = GUID("{3981e224-f559-11d3-8e3a-00c04f6837d5}")
    SF_UI_OBJECT = GUID("{3981e225-f559-11d3-8e3a-00c04f6837d5}")
    SF_VIEW_OBJECT = GUID("{3981e226-f559-11d3-8e3a-00c04f6837d5}")
    STORAGE = GUID("{3981e227-f559-11d3-8e3a-00c04f6837d5}")
    STREAM = GUID("{1CEBB3AB-7C10-499a-A417-92CA16C4CB83}")
    RANDOM_ACCESS_STREAM = GUID("{f16fc93b-77ae-4cfe-bda7-a866eea6878d}")
    LINK_TARGET_ITEM = GUID("{3981e228-f559-11d3-8e3a-00c04f6837d5}")
    STORAGE_ENUM = GUID("{4621A4E3-F0D6-4773-8A9C-46E77B174840}")
    TRANSFER = GUID("{5D080304-FE2C-48fc-84CE-CF620B0F3C53}")
    PROPERTY_STORE = GUID("{0384e1a4-1523-439c-a4c8-ab911052f586}")
    THUMBNAIL_HANDLER = GUID("{7b2e650a-8e20-4f4a-b09e-6597afc72fb0}")
    ENUM_ITEMS = GUID("{94f60519-2850-4924-aa5a-d15e84868039}")
    DATA_OBJECT = GUID("{B8C0BD9F-ED24-455c-83E6-D5390C4FE8C4}")
    ASSOC_ARRAY = GUID("{bea9ef17-82f1-4f60-9284-4f8db75c3be9}")
    FILTER = GUID("{38d08778-f557-4690-9ebf-ba54706ad8f7}")
    ENUM_ASSOC_HANDLERS = GUID("{b8ab0b9c-c2ec-4f7a-918d-314900e6280a}")
    STORAGE_ITEM = GUID("{404e2109-77d2-4699-a5a0-4fdf10db9837}")
    FILE_PLACEHOLDER = GUID("{8677DCEB-AAE0-4005-8D3D-547FA852F825}")


class IShellItem(IUnknown):
    """"""

    _iid_ = GUID("{43826d1e-e718-42ee-bc55-a1e261c37bfe}")

    __slots__ = ()


IShellItem._methods_ = [
    STDMETHOD(
        c_int32,
        "BindToHandler",
        (
            POINTER(IUnknown),
            POINTER(GUID),
            POINTER(GUID),
            POINTER(POINTER(IUnknown)),
        ),
    ),
    STDMETHOD(c_int32, "GetParent", (POINTER(POINTER(IShellItem)),)),
    STDMETHOD(c_int32, "GetDisplayName", (c_int32, POINTER(c_wchar_p))),
    STDMETHOD(c_int32, "GetAttributes", (c_uint32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "Compare", (POINTER(IShellItem), c_int32, POINTER(c_int32))),
]

# IShellItemを参照するので循環参照回避のためにここで宣言します。
from .shellitemenum import EnumShellItems, IEnumShellItems  # isort: skip  # noqa: E402


class ShellItem:
    """シェル項目。IShellItemインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IShellItem)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        """IShellItemインターフェイスをラップして初期化します。

        Args:
            o (Any): IShellItemインターフェイスに変換可能なIUnknown派生インターフェイスのインスタンス。
        """
        self.__o = query_interface(o, IShellItem)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __repr__(self) -> str:
        name = self.name_desktopabsparsing_nothrow
        return f"ShellItem({name.value_unchecked})" or "" if name else super().__repr__()

    def __str__(self) -> str:
        name = self.name_desktopabsparsing_nothrow
        return f"ShellItem({name.value_unchecked})" or "" if name else super().__str__()

    # TODO: flags
    @staticmethod
    def create_knownfolder_nothrow(folder_id: GUID, flags: int = 0) -> "ComResult[ShellItem]":
        """既知フォルダのシェルアイテムを作成します。

        Args:
            folder_id (GUID): KnownFolderID定数の値。
            flags (int, optional): _description_. Defaults to 0.

        Returns:
            ComResult[ShellItem]: _description_
        """
        global _SHCreateItemInKnownFolder

        p = POINTER(IShellItem)()
        return cr(_SHCreateItemInKnownFolder(folder_id, flags, None, IShellItem._iid_, byref(p)), ShellItem(p))

    @staticmethod
    def create_knownfolder(folder_id: GUID, flags: int = 0) -> "ShellItem":
        return ShellItem.create_knownfolder_nothrow(folder_id, flags).value

    create_knownfolder.__doc__ = create_knownfolder_nothrow.__doc__

    @staticmethod
    def create_knownfolder_item_nothrow(folder_id: GUID, itemname: str, flags: int = 0) -> "ComResult[ShellItem]":
        """既知フォルダ内のシェルアイテムを作成します。

        Args:
            folder_id (GUID): KnownFolderID定数の値。
            itemname (str): _description_
            flags (int, optional): _description_. Defaults to 0.

        Returns:
            ComResult[ShellItem]: _description_
        """
        p = POINTER(IShellItem)()
        return cr(_SHCreateItemInKnownFolder(folder_id, flags, itemname, IShellItem._iid_, byref(p)), ShellItem(p))

    @staticmethod
    def create_knownfolder_item(folder_id: GUID, itemname: str, flags: int = 0) -> "ShellItem":
        return ShellItem.create_knownfolder_item_nothrow(folder_id, itemname, flags).value

    create_knownfolder_item.__doc__ = create_knownfolder_item_nothrow.__doc__

    @staticmethod
    def create_parsingname_nothrow(name: str) -> "ComResult[ShellItem]":
        """解析名からシェルアイテムを作成します。"""
        p = POINTER(IShellItem)()
        return cr(_SHCreateItemFromParsingName(name, None, IShellItem._iid_, byref(p)), ShellItem(p))

    @staticmethod
    def create_parsingname(name: str) -> "ShellItem":
        """解析名からシェルアイテムを作成します。"""
        return ShellItem.create_parsingname_nothrow(name).value

    @staticmethod
    def create_from_idlist_nothrow(pidl: int) -> "ComResult[ShellItem]":
        """アイテムIDリストからシェルアイテムを作成します。"""
        p = POINTER(IShellItem)()
        return cr(_SHCreateItemFromIDList(pidl, IShellItem._iid_, byref(p)), ShellItem(p))

    @staticmethod
    def create_from_idlist(pidl: int) -> "ShellItem":
        """アイテムIDリストからシェルアイテムを作成します。"""
        return ShellItem.create_from_idlist_nothrow(pidl).value

    if TYPE_CHECKING:

        def bind_to_handler_nothrow[TIUnknown](
            self, bhid: GUID, type: type[TIUnknown]
        ) -> ComResult[_Pointer[TIUnknown]]:  # type: ignore
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            ...

        def bind_to_handler[TIUnknown](self, bhid: GUID, type: type[TIUnknown]) -> _Pointer[TIUnknown]:  # type: ignore
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            ...

    else:

        def bind_to_handler_nothrow[TIUnknown](self, bhid: GUID, type: type[TIUnknown]):
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            p = POINTER(type)()
            return cr(self.__o.BindToHandler(None, bhid, p._iid_, byref(p)), p)

        def bind_to_handler[TIUnknown](self, bhid: GUID, type: type[TIUnknown]):
            """バインドハンドラIDで指定されたハンドラを取得します。"""
            return self.bind_to_handler_nothrow(bhid, type).value

    @property
    def parent_nothrow(self) -> "ComResult[ShellItem]":
        """親フォルダを取得します。"""
        p = POINTER(IShellItem)()
        return cr(self.__o.GetParent(byref(p)), ShellItem(p))

    @property
    def parent(self) -> "ShellItem":
        """親フォルダを取得します。"""
        return self.parent_nothrow.value

    def __get_displayname_nothrow(self, name: int) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetDisplayName(name, byref(p)), p.value or "")

    @property
    def name_normaldisplay_nothrow(self) -> ComResult[str]:
        """標準の表示名を取得します。"""
        return self.__get_displayname_nothrow(0)

    @property
    def name_normaldisplay(self) -> str:
        """標準の表示名を取得します。"""
        return self.__get_displayname_nothrow(0).value

    @property
    def name_parentrelparsing_nothrow(self) -> ComResult[str]:
        """親フォルダを基準とした解析名を取得します。"""
        return self.__get_displayname_nothrow(0x80018001)

    @property
    def name_parentrelparsing(self) -> str:
        """親フォルダを基準とした解析名を取得します。"""
        return self.__get_displayname_nothrow(0x80018001).value

    @property
    def name_desktopabsparsing_nothrow(self) -> ComResult[str]:
        """デスクトップを基準とした解析名を取得します。"""
        return self.__get_displayname_nothrow(0x80028000)

    @property
    def name_desktopabsparsing(self) -> str:
        """デスクトップを基準とした解析名を取得します。"""
        return self.__get_displayname_nothrow(0x80028000).value

    @property
    def name_parentrelediting_nothrow(self) -> ComResult[str]:
        """親フォルダを基準とした編集名を取得します。"""
        return self.__get_displayname_nothrow(0x80031001)

    @property
    def name_parentrelediting(self) -> str:
        """親フォルダを基準とした編集名を取得します。"""
        return self.__get_displayname_nothrow(0x80031001).value

    @property
    def name_desktopabsediting_nothrow(self) -> ComResult[str]:
        """デスクトップを基準とした編集名を取得します。"""
        return self.__get_displayname_nothrow(0x8004C000)

    @property
    def name_desktopabsediting(self) -> str:
        """デスクトップを基準とした編集名を取得します。"""
        return self.__get_displayname_nothrow(0x8004C000).value

    @property
    def name_fspath_nothrow(self) -> ComResult[str]:
        """ファイルシステムパスを取得します。"""
        return self.__get_displayname_nothrow(0x80058000)

    @property
    def name_fspath(self) -> str:
        """ファイルシステムパスを取得します。"""
        return self.__get_displayname_nothrow(0x80058000).value

    @property
    def name_url_nothrow(self) -> ComResult[str]:
        """URLを取得します。"""
        return self.__get_displayname_nothrow(0x80068000)

    @property
    def name_url(self) -> str:
        """URLを取得します。"""
        return self.__get_displayname_nothrow(0x80068000).value

    @property
    def name_parentreladdressbar_nothrow(self) -> ComResult[str]:
        """アドレスバーに表示されるフレンドリ名を取得します。"""
        return self.__get_displayname_nothrow(0x8007C001)

    @property
    def name_parentreladdressbar(self) -> str:
        """アドレスバーに表示されるフレンドリ名を取得します。"""
        return self.__get_displayname_nothrow(0x8007C001).value

    @property
    def name_parentrel_nothrow(self) -> ComResult[str]:
        """相対名を取得します。"""
        return self.__get_displayname_nothrow(0x80080001)

    @property
    def name_parentrel(self) -> str:
        """相対名を取得します。"""
        return self.__get_displayname_nothrow(0x80080001).value

    @property
    def name_parentrelforui_nothrow(self) -> ComResult[str]:
        """UI用の相対名を取得します。"""
        return self.__get_displayname_nothrow(0x80094001)

    @property
    def name_parentrelforui(self) -> str:
        """UI用の相対名を取得します。"""
        return self.__get_displayname_nothrow(0x80094001).value

    @property
    def attributes_nothrow(self) -> ComResult[ShellItemAttribute]:
        """項目属性を取得します。"""
        x = c_uint32()
        return cr(self.__o.GetAttributes(0xFFFFFFFF, byref(x)), ShellItemAttribute(x.value))

    @property
    def attributes(self) -> ShellItemAttribute:
        """項目属性を取得します。"""
        return self.attributes_nothrow.value

    def compare_nothrow(self, other: "ShellItem", hint: ShellItemCompareHint) -> ComResult[int]:
        """項目を比較します。"""
        x = c_int32()
        return cr(self.__o.compare(other.__o, int(hint)), x.value)

    def compare(self, other: "ShellItem", hint: ShellItemCompareHint) -> int:
        """項目を比較します。"""
        return self.compare_nothrow(other, hint).value

    def iter_items(self) -> "Iterator[ShellItem]":
        """フォルダ内の項目を列挙します。"""
        p = self.bind_to_handler(BindHandlerID.ENUM_ITEMS.value, IEnumShellItems)
        penum = EnumShellItems(p)
        return (ShellItem(o) for o in penum)

    @property
    def items(self) -> "tuple[ShellItem, ...]":
        """フォルダ内の項目を列挙します。"""
        return tuple(self.iter_items())

    def iter_storageitems(self) -> "Iterator[ShellItem]":
        """フォルダ内のストレージ項目を列挙します。"""
        p = self.bind_to_handler(BindHandlerID.STORAGE_ENUM.value, IEnumShellItems)
        penum = EnumShellItems(p)
        return (ShellItem(o) for o in penum)

    @property
    def storageitems(self) -> "tuple[ShellItem, ...]":
        """フォルダ内のストレージ項目を列挙します。"""
        return tuple(self.iter_storageitems())

    def open_stream_nothrow(self) -> ComResult[ComStream]:
        """ファイルのストリームを開きます。"""
        o = self.bind_to_handler_nothrow(BindHandlerID.STREAM, IStream)
        return cr(o.hr, ComStream(o.value_unchecked))

    def open_stream(self) -> ComStream:
        """ファイルのストリームを開きます。"""
        return self.open_stream_nothrow().value

    @property
    def linktarget_nothrow(self) -> "ComResult[ShellItem]":
        """項目がシェルリンクの場合にリンク先項目を取得します。"""
        o = self.bind_to_handler_nothrow(BindHandlerID.LINK_TARGET_ITEM, IShellItem)
        return cr(o.hr, ShellItem(o.value_unchecked))

    @property
    def linktarget(self) -> "ShellItem":
        return self.linktarget_nothrow.value

    def get_itemid(self) -> ItemIDList:
        """アイテムIDリストを取得します。"""
        return ItemIDList.from_object(self.__o)

    def execute_fs(
        self,
        verb: str,
        invokes: bool = True,
        params: str | None = None,
        dir: str | None = None,
        showcmd: ShowCommand = ShowCommand.SHOWNORMAL,
        hotkey: int | None = None,
        monitor_handle: int | None = None,
        options: ShellExecuteOption = ShellExecuteOption.DEFAULT,
    ) -> None:
        """ファイルシステム操作を実行します。"""

        idlist = self.get_itemid()
        ShellExecute.execute_pidl(
            idlist.value or 0, verb, invokes, params, dir, showcmd, hotkey, monitor_handle, options
        )


_SHCreateItemInKnownFolder = _shell32.SHCreateItemInKnownFolder
_SHCreateItemInKnownFolder.argtypes = (
    POINTER(GUID),
    c_uint32,
    c_wchar_p,
    POINTER(GUID),
    POINTER(
        POINTER(IUnknown),
    ),
)
_SHCreateItemInKnownFolder.restype = c_int32

_SHCreateItemFromParsingName = _shell32.SHCreateItemFromParsingName
_SHCreateItemFromParsingName.argtypes = (c_wchar_p, POINTER(IUnknown), POINTER(GUID), POINTER(POINTER(IUnknown)))
_SHCreateItemFromParsingName.restype = c_int32

_SHCreateItemFromIDList = _shell32.SHCreateItemFromIDList
_SHCreateItemFromIDList.argtypes = (c_void_p, POINTER(GUID), POINTER(POINTER(IUnknown)))
_SHCreateItemFromIDList.restype = c_int32
_SHCreateItemFromIDList.restype = c_int32
