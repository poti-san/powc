"""プロパティバッグ。"""

from ctypes import (
    POINTER,
    Structure,
    byref,
    c_int16,
    c_int32,
    c_uint16,
    c_uint32,
    c_void_p,
    c_wchar_p,
)
from enum import IntEnum
from typing import Any, Sequence

from comtypes import GUID, STDMETHOD, IUnknown

from .core import ComResult, IUnknownPointer, cr, query_interface
from .errlog import ErrorLog, IErrorLog
from .variant import VARENUM, Variant


class PropertyBag2Type(IntEnum):
    """tagPROPBAG2_TYPE"""

    UNDEFINED = 0
    DATA = 1
    URL = 2
    OBJECT = 3
    STREAM = 4
    STORAGE = 5
    MONIKER = 6


class PropertyBag2Entry(Structure):
    """PROPBAG2"""

    __slots__ = ()
    _fields_ = (
        ("type", c_uint32),
        ("_vt", c_int16),
        ("cftype", c_uint16),
        ("hint", c_uint32),
        ("name", c_wchar_p),
        ("clsid", GUID),
    )

    @property
    def vt(self) -> VARENUM:
        return VARENUM(self._vt)

    @vt.setter
    def vt(self, value: VARENUM) -> None:
        self._vt = int(value)


class IPropertyBag2(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{22F55882-280B-11d0-A8A9-00A0C90C2004}")
    _methods_ = [
        STDMETHOD(
            c_int32,
            "Read",
            (c_uint32, POINTER(PropertyBag2Entry), POINTER(IErrorLog), POINTER(Variant), POINTER(c_int32)),
        ),
        STDMETHOD(c_int32, "Write", (c_uint32, POINTER(PropertyBag2Entry), POINTER(Variant))),
        STDMETHOD(c_int32, "CountProperties", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetPropertyInfo", (c_uint32, c_uint32, POINTER(PropertyBag2Entry), POINTER(c_uint32))),
        STDMETHOD(c_int32, "LoadObject", (c_wchar_p, c_uint32, POINTER(IUnknown), POINTER(IErrorLog))),
    ]


class PropertyBag2:
    """IPropertyBag2インターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IPropertyBag2)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyBag2)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    # TODO: TEST
    def read_nothrow(
        self, props: Sequence[PropertyBag2Entry], errlog: ErrorLog | None = None
    ) -> ComResult[tuple[tuple[Variant, ...], tuple[int, ...]]]:
        _props = (PropertyBag2Entry * len(props))(props)
        _values = (Variant * len(props))()
        _hrs = (c_int32 * len(props))()
        return cr(self.__o.Read(len(props), _props, _values, _hrs), (tuple(_values), tuple(_hrs)))

    # TODO: TEST
    def read(
        self, props: Sequence[PropertyBag2Entry], errlog: ErrorLog | None = None
    ) -> tuple[tuple[Variant, ...], tuple[int, ...]]:
        return self.read_nothrow(props, errlog).value

    def write_nothrow(self, props: Sequence[PropertyBag2Entry], values: Sequence[Variant]) -> ComResult[None]:
        _props = (PropertyBag2Entry * len(props))(props)
        _values = (Variant * len(props))(values)
        return cr(self.__o.Write(len(props), _props, _values), None)

    def write(self, props: Sequence[PropertyBag2Entry], values: Sequence[Variant]) -> None:
        return self.write_nothrow(props, values).value

    @property
    def propcount_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.CountProperties(byref(x)), x.value)

    @property
    def propcount(self) -> int:
        return self.propcount_nothrow.value

    def get_propinfos_nothrow(self, index: int, length: int) -> ComResult[tuple[int, tuple[PropertyBag2Entry, ...]]]:
        x1 = (PropertyBag2Entry * length)()
        x2 = c_uint32()
        return cr(self.__o.GetPropertyInfo(index, length, x1, byref(x2)), (x2.value, tuple(x1)))

    def get_propinfos(self, index: int, length: int) -> tuple[int, tuple[PropertyBag2Entry, ...]]:
        return self.get_propinfos_nothrow(index, length).value

    def loadobj_nothrow(
        self, propname: str, hint: int, outer: IUnknownPointer | None = None, errlog: ErrorLog | None = None
    ) -> ComResult[None]:
        return cr(self.__o.LoadObject(propname, hint, outer, errlog.wrapped_obj if errlog else None), None)

    def loadobj(
        self, propname: str, hint: int, outer: IUnknownPointer | None = None, errlog: ErrorLog | None = None
    ) -> None:
        return self.loadobj_nothrow(propname, hint, outer, errlog).value
