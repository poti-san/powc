"""シェルライブラリ。

主なクラスは :class:`ShellLibrary` です。
"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from enum import IntEnum, IntFlag
from os import PathLike
from typing import Any

from comtypes import GUID, STDMETHOD, CoCreateInstance, IUnknown

from powc.core import ComResult, cotaskmem, cr, query_interface
from powc.stream import StorageMode

from .knownfolderid import KnownFolderID
from .shellitem import IShellItem, ShellItem
from .shellitem2 import IShellItem2, ShellItem2
from .shellitemarray import IShellItemArray, ShellItemArray


class LibraryFolderFilter(IntEnum):
    """LIBRARYFOLDERFILTER"""

    FORCE_FILESYSTEM = 1
    STORAGE_ITEMS = 2
    ALL_ITEMS = 3


class LibraryOptionFlag(IntFlag):
    """LIBRARYOPTIONFLAGS"""

    DEFAULT = 0
    PINNED_TO_NAVPANE = 0x1
    ALL = 0x1


class DefaultSaveFolderType(IntEnum):
    """DEFAULTSAVEFOLDERTYPE"""

    DETECT = 1
    PRIVATE = 2
    PUBLIC = 3


class LibrarySaveFlag(IntFlag):
    """LIBRARYSAVEFLAGS"""

    FAIL_IF_THERE = 0
    OVERRIDE_EXISTING = 0x1
    MAKE_UNIQUE_NAME = 0x2


class IShellLibrary(IUnknown):
    _iid_ = GUID("{11a66efa-382e-451a-9234-1e0e12ef3085}")
    _methods_ = [
        STDMETHOD(c_int32, "LoadLibraryFromItem", (POINTER(IShellItem), c_uint32)),
        STDMETHOD(c_int32, "LoadLibraryFromKnownFolder", (POINTER(GUID), c_uint32)),
        STDMETHOD(c_int32, "AddFolder", (POINTER(IShellItem),)),
        STDMETHOD(c_int32, "RemoveFolder", (POINTER(IShellItem),)),
        STDMETHOD(c_int32, "GetFolders", (c_int32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "ResolveFolder", (POINTER(IShellItem), c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "GetDefaultSaveFolder", (c_int32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "SetDefaultSaveFolder", (c_int32, POINTER(IShellItem))),
        STDMETHOD(c_int32, "GetOptions", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "SetOptions", (c_int32, c_int32)),
        STDMETHOD(c_int32, "GetFolderType", (POINTER(GUID),)),
        STDMETHOD(c_int32, "SetFolderType", (POINTER(GUID),)),
        STDMETHOD(c_int32, "GetIcon", (POINTER(c_wchar_p),)),
        STDMETHOD(c_int32, "SetIcon", (c_wchar_p,)),
        STDMETHOD(c_int32, "Commit", ()),
        STDMETHOD(c_int32, "Save", (POINTER(IShellItem), c_wchar_p, c_int32, POINTER(POINTER(IShellItem)))),
        STDMETHOD(c_int32, "SaveInKnownFolder", (POINTER(GUID), c_wchar_p, c_int32, POINTER(POINTER(IShellItem)))),
    ]
    __slots__ = ()


class ShellLibrary:
    """シェルライブラリ。IShellLibraryインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IShellLibrary)

    def __init__(self, o: Any) -> None:
        """IShellLibraryインターフェイスをラップして初期化します。

        Args:
            o (Any): IShellLibraryインターフェイスに変換可能なIUnknown派生インターフェイスのインスタンス。
        """
        self.__o = query_interface(o, IShellLibrary)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "ShellLibrary":
        CLSID_ShellLibrary = GUID("{d9b3211d-e57f-4426-aaef-30a806add397}")
        return ShellLibrary(CoCreateInstance(CLSID_ShellLibrary, IShellLibrary))

    @staticmethod
    def load_fromitem(item: ShellItem, mode: StorageMode | int) -> "ShellLibrary":
        lib = ShellLibrary.create()
        lib.load_libfromitem(item, mode)
        return lib

    @staticmethod
    def load_fromknownfolder(folder_id: KnownFolderID, mode: StorageMode | int) -> "ShellLibrary":
        lib = ShellLibrary.create()
        lib.load_libfromknownfolder(folder_id, mode)
        return lib

    @staticmethod
    def load_fromparsingname(name: str | PathLike, mode: StorageMode | int) -> "ShellLibrary":
        return ShellLibrary.load_fromitem(ShellItem2.create_parsingname(str(name)), mode)

    def load_libfromitem_nothrow(self, item: ShellItem, mode: StorageMode | int) -> ComResult[None]:
        return cr(self.__o.LoadLibraryFromItem(item.wrapped_obj, int(mode)), None)

    def load_libfromitem(self, item: ShellItem, mode: StorageMode | int) -> None:
        return self.load_libfromitem_nothrow(item, mode).value

    def load_libfromknownfolder_nothrow(self, folder_id: KnownFolderID, mode: StorageMode | int) -> ComResult[None]:
        return cr(self.__o.LoadLibraryFromKnownFolder(folder_id, int(mode)), None)

    def load_libfromknownfolder(self, folder_id: KnownFolderID, mode: StorageMode | int) -> None:
        return self.load_libfromknownfolder_nothrow(folder_id, mode).value

    def add_folder_nothrow(self, item: ShellItem) -> ComResult[None]:
        return cr(self.__o.AddFolder(item.wrapped_obj), None)

    def add_folder(self, item: ShellItem) -> None:
        return self.add_folder_nothrow(item).value

    def remove_folder_nothrow(self, item: ShellItem) -> ComResult[None]:
        return cr(self.__o.RemoveFolder(item.wrapped_obj), None)

    def remove_folder(self, item: ShellItem) -> None:
        return self.remove_folder_nothrow(item).value

    def get_folders_nothrow(self, filter: LibraryFolderFilter | int) -> ComResult[ShellItemArray]:
        """ライブラリに含まれるサブフォルダのセットを取得します。

        戻り値のwrapped_interfaceはCShellItemArrayの実装するインターフェイス（IObjectCollection、IObjectArray）に変換可能です。
        """
        x = POINTER(IShellItemArray)()
        return cr(self.__o.GetFolders(int(filter), IShellItemArray._iid_, byref(x)), ShellItemArray(x))

    def get_folders(self, filter: LibraryFolderFilter | int) -> ShellItemArray:
        return self.get_folders_nothrow(filter).value

    get_folders.__doc__ = get_folders_nothrow.__doc__

    @property
    def folders_filesys_nothrow(self) -> ComResult[ShellItemArray]:
        return self.get_folders_nothrow(LibraryFolderFilter.FORCE_FILESYSTEM)

    @property
    def folders_storage_nothrow(self) -> ComResult[ShellItemArray]:
        return self.get_folders_nothrow(LibraryFolderFilter.STORAGE_ITEMS)

    @property
    def folders_all_nothrow(self) -> ComResult[ShellItemArray]:
        return self.get_folders_nothrow(LibraryFolderFilter.ALL_ITEMS)

    @property
    def folders_filesys(self) -> ShellItemArray:
        return self.get_folders(LibraryFolderFilter.FORCE_FILESYSTEM)

    @property
    def folders_storage(self) -> ShellItemArray:
        return self.get_folders(LibraryFolderFilter.STORAGE_ITEMS)

    @property
    def folders_all(self) -> ShellItemArray:
        return self.get_folders(LibraryFolderFilter.ALL_ITEMS)

    def resolve_folder_nothrow(self, folder_to_resolve: ShellItem, timeout: int) -> ComResult[ShellItem2]:
        x = POINTER(IShellItem2)()
        return cr(
            self.__o.ResolveFolder(folder_to_resolve.wrapped_obj, timeout, IShellItem2._iid_, byref(x)), ShellItem2(x)
        )

    def resolve_folder(self, folder_to_resolve: ShellItem, timeout: int) -> ShellItem2:
        return self.resolve_folder_nothrow(folder_to_resolve, timeout).value

    def get_defaultsavefolder_nothrow(self, folder_type: DefaultSaveFolderType) -> ComResult[ShellItem2]:
        x = POINTER(IShellItem2)()
        return cr(self.__o.GetDefaultSaveFolder(int(folder_type), IShellItem2._iid_, byref(x)), ShellItem2(x))

    def get_defaultsavefolder(self, folder_type: DefaultSaveFolderType) -> ShellItem2:
        return self.get_defaultsavefolder_nothrow(folder_type).value

    @property
    def defaultsavefolder_detect_nothrow(self) -> ComResult[ShellItem2]:
        return self.get_defaultsavefolder_nothrow(DefaultSaveFolderType.DETECT)

    @property
    def defaultsavefolder_public_nothrow(self) -> ComResult[ShellItem2]:
        return self.get_defaultsavefolder_nothrow(DefaultSaveFolderType.PUBLIC)

    @property
    def defaultsavefolder_private_nothrow(self) -> ComResult[ShellItem2]:
        return self.get_defaultsavefolder_nothrow(DefaultSaveFolderType.PRIVATE)

    @property
    def defaultsavefolder_detect(self) -> ShellItem2:
        return self.get_defaultsavefolder_nothrow(DefaultSaveFolderType.DETECT).value

    @property
    def defaultsavefolder_public(self) -> ShellItem2:
        return self.get_defaultsavefolder_nothrow(DefaultSaveFolderType.PUBLIC).value

    @property
    def defaultsavefolder_private(self) -> ShellItem2:
        return self.get_defaultsavefolder_nothrow(DefaultSaveFolderType.PRIVATE).value

    def set_defaultsavefolder_nothrow(self, folder_type: DefaultSaveFolderType, item: ShellItem) -> ComResult[None]:
        return cr(self.__o.SetDefaultSaveFolder(int(folder_type), item.wrapped_obj), None)

    def set_defaultsavefolder(self, folder_type: DefaultSaveFolderType, item: ShellItem) -> None:
        return self.set_defaultsavefolder_nothrow(folder_type, item).value

    def set_defaultsavefolder_detect_nothrow(self, item: ShellItem) -> ComResult[None]:
        return self.set_defaultsavefolder_nothrow(DefaultSaveFolderType.DETECT, item)

    def set_defaultsavefolder_public_nothrow(self, item: ShellItem) -> ComResult[None]:
        return self.set_defaultsavefolder_nothrow(DefaultSaveFolderType.PUBLIC, item)

    def set_defaultsavefolder_private_nothrow(self, item: ShellItem) -> ComResult[None]:
        return self.set_defaultsavefolder_nothrow(DefaultSaveFolderType.PRIVATE, item)

    @defaultsavefolder_detect.setter
    def defaultsavefolder_detect(self, item: ShellItem) -> None:
        return self.set_defaultsavefolder_detect_nothrow(item).value

    @defaultsavefolder_public.setter
    def defaultsavefolder_public(self, item: ShellItem) -> None:
        return self.set_defaultsavefolder_public_nothrow(item).value

    @defaultsavefolder_private.setter
    def defaultsavefolder_private(self, item: ShellItem) -> None:
        return self.set_defaultsavefolder_private_nothrow(item).value

    @property
    def options_nothrow(self) -> ComResult[LibraryOptionFlag]:
        x = c_int32()
        return cr(self.__o.GetOptions(byref(x)), LibraryOptionFlag(x.value))

    @property
    def options(self) -> LibraryOptionFlag:
        return self.options_nothrow.value

    def set_options_nothrow(self, mask: LibraryOptionFlag | int, options: LibraryOptionFlag | int) -> ComResult[None]:
        return self.__o.SetOptions(int(mask), int(options))

    def set_options(self, mask: LibraryOptionFlag | int, options: LibraryOptionFlag | int) -> None:
        return self.set_options_nothrow(mask, options).value

    @property
    def pinnedtonavbar_nothrow(self) -> ComResult[bool]:
        ret = self.options_nothrow
        if not ret:
            return cr(ret.hr, False)
        return cr(ret.hr, (ret.value_unchecked & LibraryOptionFlag.PINNED_TO_NAVPANE) != 0)

    @property
    def pinnedtonavbar(self) -> bool:
        return self.pinnedtonavbar_nothrow.value

    def set_pinnedtonavbar_nothrow(self, value: bool) -> ComResult[None]:
        f = LibraryOptionFlag.PINNED_TO_NAVPANE if value else 0
        return self.set_options_nothrow(LibraryOptionFlag.PINNED_TO_NAVPANE, f)

    @pinnedtonavbar.setter
    def pinnedtonavbar(self, value: bool) -> None:
        return self.set_pinnedtonavbar_nothrow(value).value

    @property
    def foldertype_nothrow(self) -> ComResult[GUID]:
        """フォルダの種類（FolderTypeID定数）を取得します。"""
        x = GUID()
        return cr(self.__o.GetFolderType(byref(x)), x)

    @property
    def foldertype(self) -> GUID:
        """フォルダの種類（FolderTypeID定数）を取得または設定します。"""
        return self.foldertype_nothrow.value

    def set_foldertype_nothrow(self, value: GUID) -> ComResult[None]:
        """フォルダの種類（FolderTypeID定数）を設定します。"""
        return cr(self.__o.SetFolderType(value), None)

    @foldertype.setter
    def foldertype(self, value: GUID) -> None:
        return self.set_foldertype_nothrow(value).value

    @property
    def icon_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as x:
            return cr(self.__o.GetIcon(byref(x)), x.value or "")

    @property
    def icon(self) -> str:
        return self.icon_nothrow.value

    def set_icon_nothrow(self, value: str) -> ComResult[None]:
        return cr(self.__o.SetIcon(value), None)

    @icon.setter
    def icon(self, value: str) -> None:
        return self.set_icon_nothrow(value).value

    def commit_nothrow(self) -> ComResult[None]:
        return cr(self.__o.Commit(), None)

    def commit(self) -> None:
        return self.commit_nothrow().value

    def save_nothrow(
        self, folder_to_save_in: ShellItem | None, lib_name: str, flags: LibrarySaveFlag | int
    ) -> ComResult[ShellItem2]:
        """ライブラリを保存します。既定の場所（FOLDERID_Libraries）に保存する場合はfolder_to_save_inをNoneにします。"""

        x = POINTER(IShellItem2)()
        return cr(
            self.__o.Save(folder_to_save_in.wrapped_obj if folder_to_save_in else None, lib_name, int(flags), byref(x)),
            ShellItem2(x),
        )

    def save(self, folder_to_save_in: ShellItem | None, lib_name: str, flags: LibrarySaveFlag | int) -> ShellItem2:
        """ライブラリを保存します。既定の場所（FOLDERID_Libraries）に保存する場合はfolder_to_save_inをNoneにします。"""
        return self.save_nothrow(folder_to_save_in, lib_name, flags).value

    def save_inknownfolder_nothrow(
        self, folder_id_to_save_in: KnownFolderID, lib_name: str, flags: LibrarySaveFlag | int
    ) -> ComResult[ShellItem2]:

        x = POINTER(IShellItem2)()
        return cr(
            self.__o.SaveInKnownFolder(folder_id_to_save_in, lib_name, int(flags), byref(x)),
            ShellItem2(x),
        )

    def save_inknownfolder(
        self, folder_id_to_save_in: KnownFolderID, lib_name: str, flags: LibrarySaveFlag | int
    ) -> ShellItem2:
        return self.save_inknownfolder_nothrow(folder_id_to_save_in, lib_name, flags).value

    def add_folderpath_nothrow(self, path: str | PathLike) -> ComResult[None]:
        item = ShellItem2.create_parsingname_nothrow(str(path))
        if not item:
            return cr(item.hr, None)
        return self.add_folder_nothrow(item.value_unchecked)

    def add_folderpath(self, path: str | PathLike) -> None:
        return self.add_folderpath_nothrow(str(path)).value

    def remove_folderpath_nothrow(self, path: str | PathLike) -> ComResult[None]:
        item = ShellItem2.create_parsingname_nothrow(str(path))
        if not item:
            return cr(item.hr, None)
        return self.remove_folder_nothrow(item.value_unchecked)

    def remove_folderpath(self, path: str | PathLike) -> None:
        return self.remove_folderpath_nothrow(str(path)).value

    def save_infolderpath_nothrow(self, path: str | PathLike, lib_name: str, flags: LibrarySaveFlag) -> ComResult[str]:
        item = ShellItem2.create_parsingname_nothrow(str(path))
        if not item:
            return cr(item.hr, "")
        ret = self.save_nothrow(item.value_unchecked, lib_name, flags)
        if not ret:
            return cr(ret.hr, "")
        return item.value_unchecked.name_desktopabsparsing_nothrow

    def save_infolderpath(self, path: str | PathLike, lib_name: str, flags: LibrarySaveFlag) -> str:
        return self.save_infolderpath_nothrow(str(path), lib_name, flags).value
