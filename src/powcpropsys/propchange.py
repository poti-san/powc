from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p
from enum import IntFlag
from typing import Any

from comtypes import GUID, STDMETHOD, IUnknown
from powc.core import ComResult, cr, query_interface

from . import _propsys
from .propkey import PropertyKey
from .propvariant import PropVariant


class IObjectWithPropertyKey(IUnknown):
    _iid_ = GUID("{fc0ca0a7-c316-4fd2-9031-3e628e6d4f23}")
    _methods_ = [
        STDMETHOD(c_int32, "SetPropertyKey", (POINTER(GUID),)),
        STDMETHOD(c_int32, "GetPropertyKey", (POINTER(GUID),)),
    ]
    __slots__ = ()


class PropertyChangeAction(IntFlag):
    """PKA_FLAGS"""

    SET = 0
    APPEND = 1
    DELETE = 2


class IPropertyChange(IObjectWithPropertyKey):
    _iid_ = GUID("{f917bc8a-1bba-4478-a245-1bde03eb9431}")
    _methods_ = [STDMETHOD(c_int32, "ApplyToPropVariant", (POINTER(PropVariant), POINTER(PropVariant)))]
    __slots__ = ()


class PropertyChange:
    """プロパティ変更情報。IPropertyChangeのラッパーです。"""

    __o: Any  # POINTER(IPropertyChange)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyChange)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_nothrow(
        flags: PropertyChangeAction, key: PropertyKey, value: PropVariant
    ) -> "ComResult[PropertyChange]":
        x = POINTER(IPropertyChange)()
        return cr(
            _PSCreateSimplePropertyChange(int(flags), key, value, IPropertyChange._iid_, byref(x)), PropertyChange(x)
        )

    @staticmethod
    def create(flags: PropertyChangeAction, key: PropertyKey, value: PropVariant) -> "PropertyChange":
        return PropertyChange.create_nothrow(flags, key, value).value

    @property
    def propkey_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetPropertyKey(byref(x)), x)

    @property
    def propkey(self) -> GUID:
        return self.propkey_nothrow.value

    def set_propkey_nothrow(self, value: GUID) -> ComResult[None]:
        return cr(self.__o.SetPropertyKey(value), None)

    @propkey.setter
    def propkey(self, value: GUID) -> None:
        return self.set_propkey_nothrow(value).value

    def apply_to_propvariant_nothrow(self, value: PropVariant) -> ComResult[PropVariant]:
        pv = PropVariant()
        return cr(self.__o.ApplyToPropVariant(value, pv), pv)

    def apply_to_propvariant(self, value: PropVariant) -> PropVariant:
        return self.apply_to_propvariant_nothrow(value).value


class IPropertyChangeArray(IUnknown):
    _iid_ = GUID("{380f5cad-1b5e-42f2-805d-637fd392d31e}")
    _methods_ = [
        STDMETHOD(c_int32, "GetCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetAt", (c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "InsertAt", (c_uint32, POINTER(IPropertyChange))),
        STDMETHOD(c_int32, "Append", (POINTER(IPropertyChange),)),
        STDMETHOD(c_int32, "AppendOrReplace", (POINTER(IPropertyChange),)),
        STDMETHOD(c_int32, "RemoveAt", (c_uint32,)),
        STDMETHOD(c_int32, "IsKeyInArray", (POINTER(PropertyKey),)),
    ]
    __slots__ = ()


class PropertyChangeArray:
    """プロパティ変更情報配列。IPropertyChangeArrayのラッパーです。"""

    __o: Any  # POINTER(IPropertyChangeArray)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyChangeArray)

    @staticmethod
    def create_nothrow() -> "ComResult[PropertyChangeArray]":
        x = POINTER(IPropertyChangeArray)()
        return cr(
            _PSCreatePropertyChangeArray(None, None, None, 0, IPropertyChangeArray._iid_, byref(x)),
            PropertyChangeArray(x),
        )

    @staticmethod
    def create() -> "PropertyChangeArray":
        return PropertyChangeArray.create_nothrow().value

    # TODO:createの初期値対応版

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def count_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetCount(byref(x)), x.value)

    @property
    def count(self) -> int:
        return self.count_nothrow.value

    def get_at_nothrow(self, index: int) -> ComResult[PropertyChange]:
        x = POINTER(IPropertyChange)()
        return cr(self.__o.GetAt(index, IPropertyChange._iid_, byref(x)), PropertyChange(x))

    def get_at(self, index: int) -> PropertyChange:
        return self.get_at_nothrow(index).value

    def insert_at_nothrow(self, index: int, item: PropertyChange) -> ComResult[None]:
        return cr(self.__o.InsertAt(index, item.wrapped_obj), None)

    def insert_at(self, index: int, item: PropertyChange) -> None:
        return self.insert_at_nothrow(index, item).value

    def append_nothrow(self, item: PropertyChange) -> ComResult[None]:
        return cr(self.__o.Append(item.wrapped_obj), None)

    def append(self, item: PropertyChange) -> None:
        return self.append_nothrow(item).value

    def append_or_replace_nothrow(self, item: PropertyChange) -> ComResult[None]:
        return cr(self.__o.AppendOrReplace(item.wrapped_obj), None)

    def append_or_replace_at(self, item: PropertyChange) -> None:
        return self.append_or_replace_nothrow(item).value

    def remove_at_nothrow(self, index: int) -> ComResult[None]:
        return cr(self.__o.RemoveAt(index), None)

    def remove_at(self, index: int) -> None:
        return self.remove_at_nothrow(index).value

    def is_key_in_array_nothrow(self, key: PropertyKey) -> ComResult[bool]:
        hr = self.__o.IsKeyInArray(key)
        return cr(hr, hr == 0)

    def is_key_in_array(self, key: PropertyKey) -> bool:
        return self.is_key_in_array_nothrow(key).value


_PSCreatePropertyChangeArray = _propsys.PSCreatePropertyChangeArray
_PSCreatePropertyChangeArray.restype = c_int32
_PSCreatePropertyChangeArray.argtypes = (
    POINTER(PropertyKey),
    POINTER(c_int32),
    POINTER(PropVariant),
    c_uint32,
    POINTER(GUID),
    POINTER(POINTER(IUnknown)),
)

_PSCreateSimplePropertyChange = _propsys.PSCreateSimplePropertyChange
_PSCreateSimplePropertyChange.restype = c_int32
_PSCreateSimplePropertyChange.argtypes = (
    c_int32,
    POINTER(PropertyKey),
    POINTER(PropVariant),
    POINTER(GUID),
    POINTER(POINTER(IUnknown)),
)
