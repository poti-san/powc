"""
プロパティ情報。
"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from dataclasses import dataclass
from enum import IntEnum, IntFlag
from typing import Any, Iterator, overload

from comtypes import GUID, STDMETHOD, IUnknown
from powc.core import ComResult, check_hresult, cotaskmem, cr, query_interface
from powc.variant import VARENUM

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


class ShellColumnState(IntFlag):
    """SHCOLSTATEF"""

    DEFAULT = 0
    TYPE_STR = 0x1
    TYPE_INT = 0x2
    TYPE_DATE = 0x3
    TYPEMASK = 0xF
    ONBYDEFAULT = 0x10
    SLOW = 0x20
    EXTENDED = 0x40
    SECONDARYUI = 0x80
    HIDDEN = 0x100
    PREFER_VARCMP = 0x200
    PREFER_FMTCMP = 0x400
    NOSORTBYFOLDERNESS = 0x800
    VIEWONLY = 0x10000
    BATCHREAD = 0x20000
    NO_GROUPBY = 0x40000
    FIXED_WIDTH = 0x1000
    NODPISCALE = 0x2000
    FIXED_RATIO = 0x4000
    DISPLAYMASK = 0xF000


class ConditionOperation(IntEnum):
    """CONDITION_OPERATION"""

    IMPLICIT = 0
    EQUAL = 1
    NOTEQUAL = 2
    LESSTHAN = 3
    GREATERTHAN = 4
    LESSTHANOREQUAL = 5
    GREATERTHANOREQUAL = 6
    VALUE_STARTSWITH = 7
    VALUE_ENDSWITH = 8
    VALUE_CONTAINS = 9
    VALUE_NOTCONTAINS = 10
    DOSWILDCARDS = 11
    WORD_EQUAL = 12
    WORD_STARTSWITH = 13
    APPLICATION_SPECIFIC = 14


class PropEnumType(IntEnum):
    DISCRETEVALUE = 0
    RANGEDVALUE = 1
    DEFAULTVALUE = 2
    ENDRANGE = 3


class IPropertyEnumType(IUnknown):
    _iid_ = GUID("{11e1fbf9-2d56-4a6b-8db3-7cd193a471f2}")
    _methods_ = [
        STDMETHOD(c_int32, "GetEnumType", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetValue", (POINTER(PropVariant),)),
        STDMETHOD(c_int32, "GetRangeMinValue", (POINTER(PropVariant),)),
        STDMETHOD(c_int32, "GetRangeSetValue", (POINTER(PropVariant),)),
        STDMETHOD(c_int32, "GetDisplayText", (POINTER(c_wchar_p),)),
    ]
    __slots__ = ()


class PropertyEnumType:
    """プロパティの列挙情報。IPropertyEnumTypeのラッパーです。"""

    __o: Any  # POINTER(IPropertyEnumType)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyEnumType)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def enumtype_nothrow(self) -> ComResult[PropEnumType]:
        x = c_int32()
        return cr(self.__o.GetEnumType(byref(x)), PropEnumType(x.value))

    @property
    def enumtype(self) -> PropEnumType:
        return self.enumtype_nothrow.value

    @property
    def value_nothrow(self) -> ComResult[PropVariant]:
        x = PropVariant()
        return cr(self.__o.GetValue(byref(x)), x)

    @property
    def value(self) -> PropVariant:
        return self.value_nothrow.value

    @property
    def range_min_nothrow(self) -> ComResult[PropVariant]:
        x = PropVariant()
        return cr(self.__o.GetRangeMinValue(byref(x)), x)

    @property
    def range_min(self) -> PropVariant:
        return self.range_min_nothrow.value

    @property
    def range_set_nothrow(self) -> ComResult[PropVariant]:
        x = PropVariant()
        return cr(self.__o.GetRangeSetValue(byref(x)), x)

    @property
    def range_set(self) -> PropVariant:
        return self.range_set_nothrow.value

    @property
    def displaytext_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetDisplayText(byref(p)), p.value or "")

    @property
    def displaytext(self) -> str:
        return self.displaytext_nothrow.value


class IPropertyEnumTypeList(IUnknown):
    _iid_ = GUID("{a99400f4-3d84-4557-94ba-1242fb2cc9a6}")
    _methods_ = [
        STDMETHOD(c_int32, "GetCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetAt", (c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        # Unused
        STDMETHOD(c_int32, "GetconditionAt", (c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "FindMatchingIndex", (POINTER(PropVariant), POINTER(c_uint32))),
    ]
    __slots__ = ()


class PropertyEnumTypeList:
    """プロパティの列挙情報リスト。PropertyEnumTypeListのラッパーです。"""

    __o: Any  # POINTER(IPropertyEnumTypeList)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyEnumTypeList)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __repr__(self) -> str:
        return "PropertyEnumTypeList"

    def __str__(self) -> str:
        return "PropertyEnumTypeList"

    def __len__(self) -> int:
        x = c_uint32()
        check_hresult(self.__o.GetCount(byref(x)))
        return x.value

    def getat(self, index: int) -> PropertyEnumType:
        """
        インデックスを指定して項目を取得します。
        キーが整数固定なので__getitem__より高速です。
        """
        x = POINTER(IPropertyEnumType)()
        check_hresult(self.__o.GetAt(int(index), IPropertyEnumType._iid_, byref(x)))
        return PropertyEnumType(x)

    @overload
    def __getitem__(self, key: int) -> PropertyEnumType: ...
    @overload
    def __getitem__(self, key: slice) -> tuple[PropertyEnumType, ...]: ...
    @overload
    def __getitem__(self, key: tuple[slice, ...]) -> tuple[PropertyEnumType, ...]: ...

    def __getitem__(self, key) -> PropertyEnumType | tuple[PropertyEnumType, ...]:
        if isinstance(key, slice):
            return tuple(self.__getitem__(i) for i in range(*key.indices(len(self))))
        if isinstance(key, tuple):
            for subslice in key:
                if not isinstance(subslice, slice):
                    raise TypeError
            return tuple(item for item in (t for t in self.__getitem__(key)))
        else:
            return self.getat(key)

    def __iter__(self) -> Iterator[PropertyEnumType]:
        return (self.getat(i) for i in range(len(self)))

    @property
    def items(self) -> tuple[PropertyEnumType, ...]:
        return tuple(iter(self))

    def find_matching_index_nothrow(self, value: PropVariant) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.FindMatchingIndex(byref(value), byref(x)), x.value)

    def find_matching_index(self, value: PropVariant) -> int:
        return self.find_matching_index_nothrow(value).value


class IPropertyDescription(IUnknown):
    """"""

    _iid_ = GUID("{6f79d558-3e96-4549-a1d1-7d75d2288814}")
    _methods_ = [
        STDMETHOD(c_int32, "GetPropertyKey", (POINTER(PropertyKey),)),
        STDMETHOD(c_int32, "GetCanonicalName", (POINTER(c_wchar_p),)),
        STDMETHOD(c_int32, "GetPropertyType", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetDisplayName", (POINTER(c_wchar_p),)),
        STDMETHOD(c_int32, "GetEditInvitation", (POINTER(c_wchar_p),)),
        STDMETHOD(c_int32, "GetTypeFlags", (c_int32, POINTER(c_int32))),
        STDMETHOD(c_int32, "GetViewFlags", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetDefaultColumnWidth", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetDisplayType", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetColumnState", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetGroupingRange", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetRelativeDescriptionType", (POINTER(c_int32),)),
        STDMETHOD(
            c_int32,
            "GetRelativeDescription",
            (POINTER(PropVariant), POINTER(PropVariant), POINTER(c_wchar_p), POINTER(c_wchar_p)),
        ),
        STDMETHOD(c_int32, "GetSortDescription", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetSortDescriptionLabel", (c_int32, POINTER(c_wchar_p))),
        STDMETHOD(c_int32, "GetAggregationType", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetConditionType", (POINTER(c_int32), POINTER(c_int32))),
        STDMETHOD(c_int32, "GetEnumTypeList", (POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "CoerceToCanonicalValue", (POINTER(PropVariant),)),
        STDMETHOD(c_int32, "FormatForDisplay", (POINTER(PropVariant), c_int32, POINTER(c_wchar_p))),
        STDMETHOD(c_int32, "IsValueCanonical", (POINTER(PropVariant),)),
    ]
    __slots__ = ()


class IPropertyDescriptionList(IUnknown):
    """"""

    _iid_ = GUID("{1f9fc1d0-c39b-4b26-817f-011967d3440e}")
    _methods_ = [
        STDMETHOD(c_int32, "GetCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetAt", (c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
    ]
    __slots__ = ()


class PropertyDescription:
    """プロパティシステムのプロパティの説明。IPropertyDescriptionのラッパーです。"""

    __o: Any  # IPropertyDescription

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyDescription)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __repr__(self) -> str:
        cname = self.canonicalname_nothrow
        return f'PropertyDescription("{cname.value_unchecked if cname else repr(self.propkey)}")'

    @property
    def propkey_nothrow(self) -> ComResult[PropertyKey]:
        x = PropertyKey()
        return cr(self.__o.GetPropertyKey(byref(x)), x)

    @property
    def propkey(self) -> PropertyKey:
        return self.propkey_nothrow.value

    @property
    def canonicalname_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetCanonicalName(byref(p)), p.value or "")

    @property
    def canonicalname(self) -> str:
        return self.canonicalname_nothrow.value

    @property
    def proptype_nothrow(self) -> ComResult[VARENUM]:
        x = c_int32()
        return cr(self.__o.GetPropertyType(byref(x)), VARENUM(x.value))

    @property
    def proptype(self) -> VARENUM:
        return self.proptype_nothrow.value

    @property
    def displayname_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetDisplayName(byref(p)), p.value or "")

    @property
    def displayname(self) -> str:
        return self.displayname_nothrow.value

    @property
    def edit_invitation_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetEditInvitation(byref(p)), p.value or "")

    @property
    def edit_invitation(self) -> str:
        return self.edit_invitation_nothrow.value

    @property
    def typeflags_nothrow(self) -> ComResult[PropDescTypeFlags]:
        x = c_int32()
        return cr(self.__o.GetTypeFlags(byref(x)), PropDescTypeFlags(x.value))

    @property
    def typeflags(self) -> PropDescTypeFlags:
        return self.typeflags_nothrow.value

    @property
    def viewflags_nothrow(self) -> ComResult[PropDescViewFlags]:
        x = c_int32()
        return cr(self.__o.GetViewFlags(byref(x)), PropDescViewFlags(x.value))

    @property
    def viewflags(self) -> PropDescViewFlags:
        return self.viewflags_nothrow.value

    @property
    def default_columnwidth_nothrow(self) -> ComResult[int]:
        x = c_int32()
        return cr(self.__o.GetDefaultColumnWidth(byref(x)), x.value)

    @property
    def default_columnwidth(self) -> int:
        return self.default_columnwidth_nothrow.value

    @property
    def displaytype_nothrow(self) -> ComResult[PropDescDisplayType]:
        x = c_int32()
        return cr(self.__o.GetDisplayType(byref(x)), PropDescDisplayType(x.value))

    @property
    def displaytype(self) -> PropDescDisplayType:
        return self.displaytype_nothrow.value

    @property
    def columnstate_nothrow(self) -> ComResult[ShellColumnState]:
        x = c_int32()
        return cr(self.__o.GetColumnState(byref(x)), ShellColumnState(x.value))

    @property
    def columnstate(self) -> ShellColumnState:
        return self.columnstate_nothrow.value

    @property
    def groupingrange_nothrow(self) -> ComResult[PropDescGroupingRange]:
        x = c_int32()
        return cr(self.__o.GetGroupingRange(byref(x)), PropDescGroupingRange(x.value))

    @property
    def groupingrange(self) -> PropDescGroupingRange:
        return self.groupingrange_nothrow.value

    @property
    def reldesctype_nothrow(self) -> ComResult[PropDescRelativeDescriptionType]:
        x = c_int32()
        return cr(self.__o.GetRelativeDescriptionType(byref(x)), PropDescRelativeDescriptionType(x.value))

    @property
    def reldesctype(self) -> PropDescRelativeDescriptionType:
        return self.reldesctype_nothrow.value

    @dataclass(frozen=True)
    class RelativeDescription:
        desc1: str
        desc2: str

    def get_relativedesc_nothrow(
        self, value1: PropVariant, value2: PropVariant
    ) -> "ComResult[PropertyDescription.RelativeDescription]":

        with cotaskmem(c_wchar_p()) as p1, cotaskmem(c_wchar_p()) as p2:
            return cr(
                self.__o.GetRelativeDescription(value1, value2, byref(p1), byref(p2)),
                PropertyDescription.RelativeDescription(p1.value or "", p2.value or ""),
            )

    def get_relativedesc(self, value1: PropVariant, value2: PropVariant) -> "PropertyDescription.RelativeDescription":
        return self.get_relativedesc_nothrow(value1, value2).value

    @property
    def sortdesc_nothrow(self) -> ComResult[PropDescSortDescription]:
        x = c_int32()
        return cr(self.__o.GetSortDescription(byref(x)), PropDescSortDescription(x.value))

    @property
    def sortdesc(self) -> PropDescSortDescription:
        return self.sortdesc_nothrow.value

    def get_sortdesclabel_nothrow(self, desc: bool) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetSortDescriptionLabel(1 if desc else 0, byref(p)), p.value or "")

    def get_sortdesclabel(self, desc: bool) -> str:
        return self.get_sortdesclabel_nothrow(desc).value

    @property
    def sortdesclabel_ascending_nothrow(self) -> ComResult[str]:
        return self.get_sortdesclabel_nothrow(False)

    @property
    def sortdesclabel_descending_nothrow(self) -> ComResult[str]:
        return self.get_sortdesclabel_nothrow(True)

    @property
    def sortdesclabel_ascending(self) -> str:
        return self.get_sortdesclabel_nothrow(False).value

    @property
    def sortdesclabel_descending(self) -> str:
        return self.get_sortdesclabel_nothrow(True).value

    @property
    def aggregationtype_nothrow(self) -> ComResult[PropDescAggregationType]:
        x = c_int32()
        return cr(self.__o.GetAggregationType(byref(x)), PropDescAggregationType(x.value))

    @property
    def aggregationtype(self) -> PropDescAggregationType:
        return self.aggregationtype_nothrow.value

    @dataclass
    class ConditionType:
        conditiontype: PropDescConditionType
        default_operation: ConditionOperation

    @property
    def conditiontype_nothrow(self) -> "ComResult[PropertyDescription.ConditionType]":
        x1 = c_int32()
        x2 = c_int32()
        return cr(
            self.__o.GetConditionType(byref(x1), byref(x2)),
            PropertyDescription.ConditionType(PropDescConditionType(x1.value), ConditionOperation(x2.value)),
        )

    @property
    def conditiontype(self) -> "PropertyDescription.ConditionType":
        return self.conditiontype_nothrow.value

    @property
    def enumtypelist_nothrow(self) -> ComResult[PropertyEnumTypeList]:
        p = POINTER(IPropertyEnumTypeList)()
        return cr(self.__o.GetEnumTypeList(IPropertyEnumTypeList._iid_, byref(p)), PropertyEnumTypeList(p))

    @property
    def enumtypelist(self) -> PropertyEnumTypeList:
        return self.enumtypelist_nothrow.value

    def corece_to_canonicalvalue_nothrow(self, value: PropVariant) -> ComResult[None]:
        return self.__o.FormatForDisplay(value)

    def corece_to_canonicalvalue(self, value: PropVariant) -> None:
        return self.corece_to_canonicalvalue_nothrow(value).value

    def format_for_display_nothrow(self, value: PropVariant, flags: PropDescFormatFlags) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return self.__o.FormatForDisplay(value, int(flags), byref(p))

    def format_for_display(self, value: PropVariant, flags: PropDescFormatFlags) -> str:
        return self.format_for_display_nothrow(value, flags).value

    def is_value_canonical_nothrow(self, value: PropVariant) -> ComResult[bool]:
        ret = self.__o.IsValueCanonical(value)
        return cr(ret.hr, ret.hr == 0)

    def is_value_canonical(self, value: PropVariant) -> bool:
        return self.is_value_canonical_nothrow(value).value


class PropertyDescriptionList:
    """プロパティシステムのプロパティの説明リスト。IPropertyDescriptionListのラッパーです。"""

    __o: Any  # IPropertyDescriptionList

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IPropertyDescriptionList)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __repr__(self) -> str:
        return "PropertyDescriptionList"

    def __str__(self) -> str:
        return "PropertyDescriptionList"

    def __len__(self) -> int:
        x = c_uint32()
        check_hresult(self.__o.GetCount(byref(x)))
        return x.value

    def getat(self, index: int) -> PropertyDescription:
        """
        インデックスを指定して項目を取得します。
        キーが整数固定なので__getitem__より高速です。
        """
        x = POINTER(IPropertyDescription)()
        check_hresult(self.__o.GetAt(int(index), IPropertyDescription._iid_, byref(x)))
        return PropertyDescription(x)

    @overload
    def __getitem__(self, key: int) -> PropertyDescription: ...
    @overload
    def __getitem__(self, key: slice) -> tuple[PropertyDescription, ...]: ...
    @overload
    def __getitem__(self, key: tuple[slice, ...]) -> tuple[PropertyDescription, ...]: ...

    def __getitem__(self, key) -> PropertyDescription | tuple[PropertyDescription, ...]:
        if isinstance(key, slice):
            return tuple(self.__getitem__(i) for i in range(*key.indices(len(self))))
        if isinstance(key, tuple):
            for subslice in key:
                if not isinstance(subslice, slice):
                    raise TypeError
            return tuple(item for item in (t for t in self.__getitem__(key)))
        else:
            return self.getat(key)

    def __iter__(self) -> Iterator[PropertyDescription]:
        return (self.getat(i) for i in range(len(self)))

    @property
    def items(self) -> tuple[PropertyDescription, ...]:
        return tuple(iter(self))
