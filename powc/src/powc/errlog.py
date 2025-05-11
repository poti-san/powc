"""COMのエラーログ。"""

from ctypes import POINTER, Structure, c_int32, c_uint16, c_uint32, c_void_p, c_wchar_p
from typing import Any

from comtypes import BSTR, GUID, STDMETHOD, IUnknown

from .core import ComResult, cr, query_interface


class ComExceptionInfo(Structure):
    """EXCEPINFO"""

    __slots__ = ()
    _fields_ = (
        ("code", c_uint16),
        ("reserved", c_uint16),
        ("source", BSTR),
        ("description", BSTR),
        ("helpfile", BSTR),
        ("helpcontext", c_uint32),
        ("reserved2", c_void_p),
        ("pfn_deferred_fillin", c_void_p),
        ("scode", c_int32),
    )


class IErrorLog(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{3127CA40-446E-11CE-8135-00AA004BB851}")
    _methods_ = [
        STDMETHOD(c_int32, "AddError", (c_wchar_p, POINTER(ComExceptionInfo))),
    ]


class ErrorLog:
    """エラーログ。IErrorLogインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IErrorLog)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IErrorLog)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def add_error_nothrow(self, propname: str, info: ComExceptionInfo) -> ComResult[None]:
        return cr(self.__o.AddError(propname, info), None)

    def add_error(self, propname: str, info: ComExceptionInfo) -> None:
        return self.add_error_nothrow(propname, info).value

    def add_error_code_nothrow(
        self,
        propname: str,
        code: int,
        source: str,
        description: str | None = None,
        helpfile: str | None = None,
        helpcontext: int = 0,
    ) -> ComResult[None]:
        info = ComExceptionInfo()
        info.source = source
        info.description = description
        info.helpfile = helpfile
        info.helpcontext = helpcontext
        info.scode = code
        return cr(self.__o.AddError(propname, info), None)

    def add_error_code(
        self,
        propname: str,
        code: int,
        source: str,
        description: str | None = None,
        helpfile: str | None = None,
        helpcontext: int = 0,
    ) -> None:
        return self.add_error_code_nothrow(propname, code, source, description, helpfile, helpcontext).value
