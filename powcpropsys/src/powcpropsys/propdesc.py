"""
プロパティ情報。
"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from typing import Any, Iterator, overload

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, check_hresult, cotaskmem, cr, queryinterface

from .propkey import PropertyKey
from .propvariant import PropVariant


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
    """プロパティシステムのプロパティの説明。IPropertyDescriptionのラッパー。"""

    __o: Any  # IPropertyDescription

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = queryinterface(o, IPropertyDescription)

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


# TODO
# STDMETHOD(c_int32, "GetPropertyType", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetDisplayName", (POINTER(c_wchar_p),)),
# STDMETHOD(c_int32, "GetEditInvitation", (POINTER(c_wchar_p),)),
# STDMETHOD(c_int32, "GetTypeFlags", (c_int32, POINTER(c_int32))),
# STDMETHOD(c_int32, "GetViewFlags", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetDefaultColumnWidth", (POINTER(c_uint32),)),
# STDMETHOD(c_int32, "GetDisplayType", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetColumnState", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetGroupingRange", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetRelativeDescriptionType", (POINTER(c_int32),)),
# STDMETHOD(
#     c_int32,
#     "GetRelativeDescription",
#     (POINTER(PropVariant), POINTER(PropVariant), POINTER(c_wchar_p), POINTER(c_wchar_p)),
# ),
# STDMETHOD(c_int32, "GetSortDescription", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetSortDescriptionLabel", (c_int32, POINTER(c_wchar_p))),
# STDMETHOD(c_int32, "GetAggregationType", (POINTER(c_int32),)),
# STDMETHOD(c_int32, "GetConditionType", (POINTER(c_int32), POINTER(c_int32))),
# STDMETHOD(c_int32, "GetEnumTypeList", (POINTER(GUID), POINTER(POINTER(IUnknown)))),
# STDMETHOD(c_int32, "CoerceToCanonicalValue", (POINTER(PropVariant),)),
# STDMETHOD(c_int32, "FormatForDisplay", (POINTER(PropVariant), c_int32, POINTER(c_wchar_p))),
# STDMETHOD(c_int32, "IsValueCanonical", (POINTER(PropVariant),)),


class PropertyDescriptionList:
    """プロパティシステムのプロパティの説明リスト。IPropertyDescriptionListのラッパー。"""

    __o: Any  # IPropertyDescriptionList

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = queryinterface(o, IPropertyDescriptionList)

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
