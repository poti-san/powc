"""インターネットショートカット。

主なクラスは :class:`UniformResourceLocator` です。
"""

from ctypes import (
    POINTER,
    Structure,
    byref,
    c_int32,
    c_uint32,
    c_void_p,
    c_wchar_p,
    sizeof,
)
from enum import IntFlag
from typing import Any

from comtypes import GUID, STDMETHOD, CoCreateInstance, IUnknown

from powc.core import ComResult, cotaskmem, cr, guid_from_define, query_interface
from powc.persist import PersistFile, PersistStream
from powc.stream import ComStream, StorageMode


class IUrlInvokeCommandFlag(IntFlag):
    """IURL_INVOKECOMMAND_FLAGS"""

    DDEWAIT = 0x0004
    ASYNCOK = 0x0008
    LOG_USAGE = 0x0010


class IUrlSetUrlFlag(IntFlag):
    """IURL_SETURL_FLAGS"""

    GUESS_PROTOCOL = 0x0001
    USE_DEFAULT_PROTOCOL = 0x0002


class _URLINVOKECOMMANDINFOW(Structure):
    _fields_ = (
        ("dwcbSize", c_uint32),
        ("dwFlags", c_uint32),  # IURL_INVOKECOMMAND_FLAGS
        ("hwndParent", c_void_p),  # if IURL_INVOKECOMMAND_FL_ALLOW_UI
        ("pcszVerb", c_wchar_p),  # if not IURL_INVOKECOMMAND_FL_USE_DEFAULT_VERB
    )
    _slots_ = ("dwcbSize", "dwFlags", "hwndParent", "pcszVerb")


class IUniformResourceLocatorW(IUnknown):
    """"""

    _iid_ = GUID("{cabb0da0-da57-11cf-9974-0020afd79762}")
    _methods_ = [
        STDMETHOD(c_int32, "SetURL", (c_wchar_p, c_uint32)),
        STDMETHOD(c_int32, "GetURL", (POINTER(c_wchar_p),)),
        STDMETHOD(c_int32, "InvokeCommand", (POINTER(_URLINVOKECOMMANDINFOW),)),
    ]
    __slots__ = ()


class UniformResourceLocator:
    """インターネットショートカット。IUniformResourceLocatorWのラッパーです。

    Examples:
        空のインターネットショートカット作成

        >>> UniformResourceLocator.create()

        .lnkファイルを開いてインターネットショートカット作成

        >>> UniformResourceLocator.create_from_file(path)

        ストリームからインターネットショートカット作成

        >>> UniformResourceLocator.create_from_stream(stream)
    """

    __IURL_INVOKECOMMAND_FL_ALLOW_UI = 0x0001
    __IURL_INVOKECOMMAND_FL_USE_DEFAULT_VERB = 0x0002

    __o: Any  # POINTER(IUniformResourceLocatorW)
    __perfile: PersistFile
    __perstream: PersistStream

    __slots__ = ("__o", "__perfile", "__perstream")

    def __init__(self, o: Any):
        self.__o = query_interface(o, IUniformResourceLocatorW)
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

    @staticmethod
    def create() -> "UniformResourceLocator":
        CLSID_InternetShortcut = guid_from_define(
            0xFBF23B40, 0xE3F0, 0x101B, 0x84, 0x88, 0x00, 0xAA, 0x00, 0x3E, 0x56, 0xF8
        )
        return UniformResourceLocator(CoCreateInstance(CLSID_InternetShortcut, IUniformResourceLocatorW))

    @staticmethod
    def create_fromfile(path: str, mode: StorageMode = StorageMode.READ) -> "UniformResourceLocator":
        link = UniformResourceLocator.create()
        link.persist_file.load(path, mode)
        return link

    @staticmethod
    def create_fromstream(stream: ComStream) -> "UniformResourceLocator":
        link = UniformResourceLocator.create()
        link.persist_stream.load(stream)
        return link

    @property
    def url_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetURL(byref(p)), p.value or "")

    @property
    def url(self) -> str:
        return self.url_nothrow.value

    def set_url_nothrow(self, url: str, flags: IUrlSetUrlFlag) -> ComResult[None]:
        return cr(self.__o.SetURL(url, int(flags)), None)

    def set_url(self, url: str, flags, IUrlSetUrlFlag) -> None:
        return self.set_url_nothrow(url, flags).value

    def invoke_nothrow(
        self, verb: str, flags: IUrlInvokeCommandFlag, window_handle: int | None = None
    ) -> ComResult[None]:
        ic = _URLINVOKECOMMANDINFOW()
        ic.dwcbSize = sizeof(ic)
        ic.flags = flags | (self.__IURL_INVOKECOMMAND_FL_ALLOW_UI if window_handle else 0)
        ic.hwndParent = window_handle
        ic.pcszVerb = verb
        return cr(self.__o.InvokeCommand(), None)

    def invoke(self, cmd: str, flags: IUrlInvokeCommandFlag, window_handle: int) -> None:
        return self.invoke_nothrow(cmd, flags, window_handle).value

    def invoke_default_nothrow(self, flags: IUrlInvokeCommandFlag, window_handle: int | None = None) -> ComResult[None]:
        ic = _URLINVOKECOMMANDINFOW()
        ic.dwcbSize = sizeof(ic)
        ic.flags = (
            flags
            | (self.__IURL_INVOKECOMMAND_FL_ALLOW_UI if window_handle else 0)
            | self.__IURL_INVOKECOMMAND_FL_USE_DEFAULT_VERB
        )
        ic.hwndParent = window_handle
        return cr(self.__o.InvokeCommand(), None)

    def invoke_default(self, flags: IUrlInvokeCommandFlag, window_handle: int) -> None:
        return self.invoke_default_nothrow(flags, window_handle).value
        return self.invoke_default_nothrow(flags, window_handle).value
