"""プロパティシステム。"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from typing import Any

from comtypes import GUID, STDMETHOD, IUnknown
from powc.core import ComResult, cotaskmem, cr, query_interface

from . import _propsys
from .propdesc import (
    IPropertyDescription,
    IPropertyDescriptionList,
    PropDescEnumFilter,
    PropDescFormatFlags,
    PropertyDescription,
    PropertyDescriptionList,
)
from .propkey import PropertyKey
from .propvariant import PropVariant


class IPropertySystem(IUnknown):
    """"""

    _iid_ = GUID("{ca724e8a-c3e6-442b-88a4-6fb0db8035a3}")
    _methods_ = [
        STDMETHOD(c_int32, "GetPropertyDescription", (POINTER(PropertyKey), POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "GetPropertyDescriptionByName", (c_wchar_p, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(
            c_int32, "GetPropertyDescriptionListFromString", (c_wchar_p, POINTER(GUID), POINTER(POINTER(IUnknown)))
        ),
        STDMETHOD(c_int32, "EnumeratePropertyDescriptions", (c_int32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(
            c_int32, "FormatForDisplay", (POINTER(PropertyKey), POINTER(PropVariant), c_int32, c_wchar_p, c_uint32)
        ),
        STDMETHOD(
            c_int32, "FormatForDisplayAlloc", (POINTER(PropertyKey), POINTER(PropVariant), c_int32, POINTER(c_wchar_p))
        ),
        STDMETHOD(c_int32, "RegisterPropertySchema", (c_wchar_p,)),
        STDMETHOD(c_int32, "UnregisterPropertySchema", (c_wchar_p,)),
        STDMETHOD(c_int32, "RefreshPropertySchema", ()),
    ]
    __slots__ = ()


class PropertySystem:
    """プロパティシステム。IPropertySystemインターフェイスのラッパーです。

    Examples:
        >>> propsys = PropertySystem.create()
    """

    __o: Any  # IPropertySystem

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertySystem)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_nothrow() -> "ComResult[PropertySystem]":
        p = POINTER(IPropertySystem)()
        global _PSGetPropertySystem
        return cr(_PSGetPropertySystem(IPropertySystem._iid_, byref(p)), PropertySystem(p))

    @staticmethod
    def create() -> "PropertySystem":
        return PropertySystem.create_nothrow().value

    def __repr__(self) -> str:
        return "PropertySystem"

    def __str__(self) -> str:
        return "PropertySystem"

    def get_propdesc_nothrow(self, key: PropertyKey) -> ComResult[PropertyDescription]:
        x = POINTER(IPropertyDescription)()
        return cr(
            self.__o.GetPropertyDescription(byref(key), IPropertyDescription._iid_, byref(x)),
            PropertyDescription(x),
        )

    def get_propdesc(self, key: PropertyKey) -> PropertyDescription:
        return self.get_propdesc_nothrow(key).value

    def get_propdesc_by_name_nothrow(self, name: str) -> ComResult[PropertyDescription]:
        x = POINTER(IPropertyDescription)()
        return cr(
            self.__o.GetPropertyDescriptionByName(name, IPropertyDescription._iid_, byref(x)),
            PropertyDescription(x),
        )

    def get_propdesc_by_name(self, name: str) -> PropertyDescription:
        return self.get_propdesc_by_name_nothrow(name).value

    def get_propdescs_from_string_nothrow(self, proplist: str) -> ComResult[PropertyDescriptionList]:
        x = POINTER(IPropertyDescriptionList)()
        return cr(
            self.__o.GetPropertyDescriptionListFromString(proplist, IPropertyDescriptionList._iid_, byref(x)),
            PropertyDescriptionList(x),
        )

    def get_propdescs_from_string(self, proplist: str) -> PropertyDescriptionList:
        return self.get_propdescs_from_string_nothrow(proplist).value

    def get_propdescs_nothrow(self, filter: PropDescEnumFilter) -> ComResult[PropertyDescriptionList]:
        x = POINTER(IPropertyDescriptionList)()
        return cr(
            self.__o.EnumeratePropertyDescriptions(int(filter), IPropertyDescriptionList._iid_, byref(x)),
            PropertyDescriptionList(x),
        )

    def get_propdescs(self, filter: PropDescEnumFilter) -> PropertyDescriptionList:
        return self.get_propdescs_nothrow(filter).value

    @property
    def propdescs_all_nothrow(self) -> ComResult[PropertyDescriptionList]:
        return self.get_propdescs_nothrow(PropDescEnumFilter.ALL)

    @property
    def propdescs_all(self) -> PropertyDescriptionList:
        return self.get_propdescs(PropDescEnumFilter.ALL)

    @property
    def propdescs_system_nothrow(self) -> ComResult[PropertyDescriptionList]:
        return self.get_propdescs_nothrow(PropDescEnumFilter.SYSTEM)

    @property
    def propdescs_system(self) -> PropertyDescriptionList:
        return self.get_propdescs(PropDescEnumFilter.SYSTEM)

    def formatfordisplay_nothrow(
        self, key: PropertyKey, value: PropVariant, format: PropDescFormatFlags
    ) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.FormatForDisplayAlloc(byref(key), byref(value), int(format), byref(p)), p.value or "")

    def formatfordisplay(self, key: PropertyKey, value: PropVariant, format: PropDescFormatFlags) -> str:
        return self.formatfordisplay_nothrow(key, value, format).value

    def register_propscheme_nothrow(self, path: str) -> ComResult[None]:
        return cr(self.__o.RegisterPropertySchema(path), None)

    def register_propscheme(self, path: str) -> None:
        return self.register_propscheme_nothrow(path).value

    def unregister_propscheme_nothrow(self, path: str) -> ComResult[None]:
        return cr(self.__o.UnregisterPropertySchema(path), None)

    def unregister_propscheme(self, path: str) -> None:
        return self.unregister_propscheme_nothrow(path).value

    def refresh_propscheme_nothrow(self) -> ComResult[None]:
        return cr(self.__o.RefreshPropertySchema(), None)

    def refresh_propscheme(self) -> None:
        return self.refresh_propscheme_nothrow().value

    # ユーティリティメソッド
    def get_propkeys_all(self) -> tuple[PropertyKey, ...]:
        return tuple(propdesc.propkey for propdesc in self.propdescs_all)

    def get_propkeys_system(self) -> tuple[PropertyKey, ...]:
        return tuple(propdesc.propkey for propdesc in self.propdescs_system)


_PSGetPropertySystem = _propsys.PSGetPropertySystem
_PSGetPropertySystem.argtypes = (POINTER(GUID), POINTER(POINTER(IUnknown)))
_PSGetPropertySystem.restype = c_int32
