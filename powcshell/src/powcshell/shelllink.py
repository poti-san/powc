"""ショートカット。

主なクラスは :class:`ShellLink` です。"""

from ctypes import (
    POINTER,
    byref,
    c_int32,
    c_uint16,
    c_uint32,
    c_void_p,
    c_wchar,
    c_wchar_p,
)
from enum import IntFlag
from os import PathLike
from typing import Any, Final

from comtypes import GUID, STDMETHOD, CoCreateInstance, IUnknown

from powc.core import ComResult, check_hresult, cr, query_interface
from powc.persist import PersistFile, PersistStream
from powc.stream import ComStream, StorageMode

from .itemidlist import ItemIDList


class IShellLinkW(IUnknown):
    """"""

    _iid_ = GUID("{000214F9-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "GetPath", (c_wchar_p, c_int32, c_void_p, c_uint32)),
        STDMETHOD(c_int32, "GetIDList", (POINTER(c_void_p),)),
        STDMETHOD(c_int32, "SetIDList", (c_void_p,)),
        STDMETHOD(c_int32, "GetDescription", (c_wchar_p, c_int32)),
        STDMETHOD(c_int32, "SetDescription", (c_wchar_p,)),
        STDMETHOD(c_int32, "GetWorkingDirectory", (c_wchar_p, c_int32)),
        STDMETHOD(c_int32, "SetWorkingDirectory", (c_wchar_p,)),
        STDMETHOD(c_int32, "GetArguments", (c_wchar_p, c_int32)),
        STDMETHOD(c_int32, "SetArguments", (c_wchar_p,)),
        STDMETHOD(c_int32, "GetHotkey", (POINTER(c_uint16),)),
        STDMETHOD(c_int32, "SetHotkey", (c_uint16,)),
        STDMETHOD(c_int32, "GetShowCmd", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "SetShowCmd", (c_int32,)),
        STDMETHOD(c_int32, "GetIconLocation", (c_wchar_p, c_int32, POINTER(c_int32))),
        STDMETHOD(c_int32, "SetIconLocation", (c_wchar_p, c_int32)),
        STDMETHOD(c_int32, "SetRelativePath", (c_wchar_p, c_uint32)),
        STDMETHOD(c_int32, "Resolve", (c_void_p, c_uint32)),
        STDMETHOD(c_int32, "SetPath", (c_wchar_p,)),
    ]
    __slots__ = ()


class ShellLinkGetPath(IntFlag):
    DEFAULT = 0
    SHORTPATH = 0x1
    UNCPRIORITY = 0x2
    RAWPATH = 0x4
    RELATIVEPRIORITY = 0x8


class ShellLinkResolve(IntFlag):
    NONE = 0
    NO_UI = 0x1
    ANY_MATCH = 0x2
    UPDATE = 0x4
    NOUPDATE = 0x8
    NOSEARCH = 0x10
    NOTRACK = 0x20
    NOLINKINFO = 0x40
    INVOKE_MSI = 0x80
    NO_UI_WITH_MSG_PUMP = 0x101
    OFFER_DELETE_WITHOUT_FILE = 0x200
    KNOWNFOLDER = 0x400
    MACHINE_IN_LOCAL_TARGET = 0x800
    UPDATE_MACHINE_AND_SID = 0x1000
    NO_OBJECT_ID = 0x2000


class ShellLink:
    """ショートカット(.lnk)。IShellLinkWのラッパーです。

    Examples:
        空のショートカット作成

        >>> ShellLink.create()

        .lnkファイルを開いてショートカット作成

        >>> ShellLink.create_from_file(path)

        ストリームからショートカット作成

        >>> ShellLink.create_from_stream(stream)
    """

    __MAX_PATH: Final[int] = 260
    __INFOTIPSIZE: Final[int] = 1024

    __slots__ = ("__o", "__perfile", "__perstream")
    __o: Any  # POINTER(IShellLinkW)
    __perfile: PersistFile
    __perstream: PersistStream

    def __init__(self, o: Any):
        self.__o = query_interface(o, IShellLinkW)
        self.__perfile = PersistFile(o)
        self.__perstream = PersistStream(o)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def persist_file(self) -> PersistFile:
        return self.__perfile

    @property
    def persist_stream(self) -> PersistStream:
        return self.__perstream

    def __repr__(self) -> str:
        s = self.persist_file.curfile_nothrow
        return f'ShellLink("{s.value_unchecked}")' if s else super().__repr__()

    def __str__(self) -> str:
        s = self.persist_file.curfile_nothrow
        return f'ShellLink("{s.value_unchecked}")' if s else super().__str__()

    @staticmethod
    def create() -> "ShellLink":
        CLSID_ShellLink = GUID("{00021401-0000-0000-C000-000000000046}")
        return ShellLink(CoCreateInstance(CLSID_ShellLink, IShellLinkW))

    @staticmethod
    def create_from_file(path: str | PathLike, mode: StorageMode = StorageMode.READ) -> "ShellLink":
        link = ShellLink.create()
        link.persist_file.load(str(path), mode)
        return link

    @staticmethod
    def create_from_stream(stream: ComStream) -> "ShellLink":
        link = ShellLink.create()
        link.persist_stream.load(stream)
        return link

    def get_path_nothrow(self, flags: ShellLinkGetPath) -> ComResult[str]:
        buf = (c_wchar * ShellLink.__MAX_PATH)()
        return cr(self.__o.GetPath(buf, len(buf), None, int(flags)), buf.value)

    def get_path(self, flags: ShellLinkGetPath) -> str:
        return self.get_path_nothrow(flags).value

    def set_path_nothrow(self, path: str | PathLike) -> ComResult[None]:
        return cr(self.__o.SetPath(str(path)), None)

    def set_path(self, path: str | PathLike) -> None:
        return self.set_path_nothrow(str(path)).value

    @property
    def path_nothrow(self) -> ComResult[str]:
        return self.get_path_nothrow(ShellLinkGetPath.DEFAULT)

    @property
    def path(self) -> str:
        return self.path_nothrow.value

    @path.setter
    def path(self, path: str) -> None:
        self.set_path(path)

    @property
    def shortpath_nothrow(self) -> ComResult[str]:
        return self.get_path_nothrow(ShellLinkGetPath.SHORTPATH)

    @property
    def shortpath(self) -> str:
        return self.shortpath_nothrow.value

    @property
    def rawpath_nothrow(self) -> ComResult[str]:
        return self.get_path_nothrow(ShellLinkGetPath.RAWPATH)

    @property
    def rawpath(self) -> str:
        return self.rawpath_nothrow.value

    @property
    def idlist_nothrow(self) -> ComResult[ItemIDList]:
        x = ItemIDList()
        return cr(self.__o.GetIDList(byref(x)), x)

    def set_idlist_nothrow(self, pidl: ItemIDList) -> None:
        self.__o.SetIDList(pidl)

    @property
    def idlist(self) -> ItemIDList:
        return self.idlist_nothrow.value

    @idlist.setter
    def idlist(self, pidl: ItemIDList) -> None:
        check_hresult(self.__o.SetIDList(pidl))

    @property
    def description_nothrow(self) -> ComResult[str]:
        buf = (c_wchar * ShellLink.__INFOTIPSIZE)()
        return cr(self.__o.GetDescription(buf, len(buf)), buf.value)

    def set_description_nothrow(self, value: str) -> ComResult[None]:
        return cr(self.__o.SetDescription(value), None)

    @property
    def description(self) -> str:
        return self.description_nothrow.value

    @description.setter
    def description(self, value: str) -> None:
        check_hresult(self.__o.SetDescription(value))

    @property
    def workdir_nothrow(self) -> ComResult[str]:
        buf = (c_wchar * ShellLink.__MAX_PATH)()
        return cr(self.__o.GetWorkingDirectory(buf, len(buf)), buf.value)

    @property
    def workdir(self) -> str:
        return self.workdir_nothrow.value

    def set_workdir_nothrow(self, value: str) -> ComResult[None]:
        return cr(self.__o.SetPath(value), None)

    @workdir.setter
    def workdir(self, value: str) -> None:
        return self.set_workdir_nothrow(value).value

    @property
    def arguments_nothrow(self) -> ComResult[str]:
        buf = (c_wchar * ShellLink.__INFOTIPSIZE)()
        return cr(self.__o.GetArguments(buf, len(buf)), buf.value)

    @property
    def arguments(self) -> str:
        return self.arguments_nothrow.value

    def set_arguments_nothrow(self, args: str) -> ComResult[None]:
        return cr(self.__o.SetArguments(args), None)

    @arguments.setter
    def arguments(self, args: str) -> None:
        return self.set_arguments_nothrow(args).value

    @property
    def hotkey_nothrow(self) -> ComResult[int]:
        x = c_uint16()
        return cr(self.__o.GetHotkey(byref(x)), x.value)

    @property
    def hotkey(self) -> int:
        return self.hotkey_nothrow.value

    def set_hotkey_nothrow(self, hotkey: int) -> ComResult[None]:
        return cr(self.__o.SetHotkey(hotkey), None)

    @hotkey.setter
    def hotkey(self, hotkey: int) -> None:
        return self.set_hotkey_nothrow(hotkey).value

    @property
    def showcmd_nothrow(self) -> ComResult[int]:
        x = c_int32()
        return cr(self.__o.GetShowCmd(byref(x)), x.value)

    @property
    def showcmd(self) -> int:
        return self.showcmd_nothrow.value

    def set_showcmd_nothrow(self, showcmd: int) -> ComResult[None]:
        return cr(self.__o.GetShowCmd(showcmd), None)

    @showcmd.setter
    def showcmd(self, showcmd: int) -> None:
        return self.set_showcmd_nothrow(showcmd).value

    @property
    def iconlocation_nothrow(self) -> ComResult[tuple[str, int]]:
        buf = (c_wchar * ShellLink.__MAX_PATH)()
        x = c_int32()
        return cr(self.__o.GetIconLocation(buf, len(buf), byref(x)), (buf.value, x.value))

    @property
    def iconlocation(self) -> tuple[str, int]:
        return self.iconlocation_nothrow.value

    def set_iconlocation_nothrow(self, info: tuple[str, int]) -> ComResult[None]:
        return cr(self.__o.SetIconLocation(info[0], info[1]), None)

    @iconlocation.setter
    def iconlocation(self, info: tuple[str, int]) -> None:
        return self.set_iconlocation_nothrow(info).value

    def set_relpath_nothrow(self, path: str) -> ComResult[None]:
        return cr(self.__o.SetRelativePath(path, 0), None)

    def set_relpath(self, path: str) -> None:
        return self.set_relpath_nothrow(path).value

    def resolve_nothrow(self, flags: ShellLinkResolve, window_handle: int = 0) -> ComResult[None]:
        return cr(self.__o.Resolve(window_handle, int(flags)), None)

    def resolve(self, flags: ShellLinkResolve, window_handle: int = 0) -> None:
        return self.resolve_nothrow(flags, window_handle).value
