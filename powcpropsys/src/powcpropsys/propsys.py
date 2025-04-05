"""プロパティシステム。"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p, windll
from enum import IntEnum, IntFlag
from typing import Any

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, cotaskmem, cr, queryinterface

from .propdesc import (
    IPropertyDescription,
    IPropertyDescriptionList,
    PropertyDescription,
    PropertyDescriptionList,
)
from .propkey import PropertyKey
from .propvariant import PropVariant


class PropDescEnumFilter(IntEnum):
    """PROPDESC_ENUMFILTER"""

    ALL = 0
    SYSTEM = 1
    NONSYSTEM = 2
    VIEWABLE = 3
    QUERYABLE = 4
    INFULLTEXTQUERY = 5
    COLUMN = 6


class PropDescTypeFlags(IntFlag):
    """PROPDESC_TYPE_FLAGS"""

    DEFAULT = 0
    MULTIPLEVALUES = 0x1
    IS_INNATE = 0x2
    IS_GROUP = 0x4
    CAN_GROUPBY = 0x8
    CAN_STACKBY = 0x10
    IS_TREEPROPERTY = 0x20
    INCLUDE_INFULLTEXTQUERY = 0x40
    IS_VIEWABLE = 0x80
    IS_QUERYABLE = 0x100
    CANBE_PURGED = 0x200
    SEARCH_RAWVALUE = 0x400
    DONT_COERCE_EMPTYSTRINGS = 0x800
    ALWAYS_INSUPPLEMENTAL_STORE = 0x1000
    IS_SYSTEM_PROPERTY = 0x80000000
    MASK_ALL = 0x80001FFF


class PropDescViewFlags(IntFlag):
    """PROPDESC_VIEW_FLAGS"""

    DEFAULT = 0
    CENTER_ALIGN = 0x1
    RIGHT_ALIGN = 0x2
    BEGIN_NEWGROUP = 0x4
    FILL_AREA = 0x8
    SORT_DESCENDING = 0x10
    SHOW_ONLYIFPRESENT = 0x20
    SHOW_BYDEFAULT = 0x40
    SHOW_INPRIMARYLIST = 0x80
    SHOW_INSECONDARYLIST = 0x100
    HIDE_LABEL = 0x200
    HIDDEN = 0x800
    CANWRAP = 0x1000
    MASK_ALL = 0x1BFF


class PropDescDisplayType(IntEnum):
    """PROPDESC_DISPLAYTYPE"""

    STRING = 0
    NUMBER = 1
    BOOLEAN = 2
    DATETIME = 3
    ENUMERATED = 4


class PropDescGroupingRange(IntEnum):
    """PROPDESC_GROUPING_RANGE"""

    DISCRETE = 0
    ALPHANUMERIC = 1
    SIZE = 2
    DYNAMIC = 3
    DATE = 4
    PERCENT = 5
    ENUMERATED = 6


class PropDescFormatFlags(IntFlag):
    """PROPDESC_FORMAT_FLAGS"""

    DEFAULT = 0
    PREFIXNAME = 0x1
    FILENAME = 0x2
    ALWAYSKB = 0x4
    RESERVED_RIGHTTOLEFT = 0x8
    SHORTTIME = 0x10
    LONGTIME = 0x20
    HIDETIME = 0x40
    SHORTDATE = 0x80
    LONGDATE = 0x100
    HIDEDATE = 0x200
    RELATIVEDATE = 0x400
    USEEDITINVITATION = 0x800
    READONLY = 0x1000
    NOAUTOREADINGORDER = 0x2000


class PropDescSortDescription(IntEnum):
    """PROPDESC_SORTDESCRIPTION"""

    GENERAL = 0
    A_Z = 1
    LOWEST_HIGHEST = 2
    SMALLEST_BIGGEST = 3
    OLDEST_NEWEST = 4


class PropDescRelativeDescriptionType(IntEnum):
    """PROPDESC_RELATIVEDESCRIPTION_TYPE"""

    GENERAL = 0
    DATE = 1
    SIZE = 2
    COUNT = 3
    REVISION = 4
    LENGTH = 5
    DURATION = 6
    SPEED = 7
    RATE = 8
    RATING = 9
    PRIORITY = 10


class PropDescAggregationType(IntEnum):
    """PROPDESC_AGGREGATION_TYPE"""

    DEFAULT = 0
    FIRST = 1
    SUM = 2
    AVERAGE = 3
    DATERANGE = 4
    UNION = 5
    MAX = 6
    MIN = 7


class PropDescConditionType(IntEnum):
    """PROPDESC_CONDITION_TYPE"""

    NONE = 0
    STRING = 1
    SIZE = 2
    DATETIME = 3
    BOOLEAN = 4
    NUMBER = 5


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
    """プロパティシステム。IPropertySystemインターフェイスのラッパー。

    Examples:
        >>> propsys = PropertySystem.create()
    """

    __o: Any  # IPropertySystem

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = queryinterface(o, IPropertySystem)

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


_PSGetPropertySystem = windll.propsys.PSGetPropertySystem
_PSGetPropertySystem.argtypes = (POINTER(GUID), POINTER(POINTER(IUnknown)))
_PSGetPropertySystem.restype = c_int32
