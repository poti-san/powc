from ctypes import byref, c_wchar
from typing import Any, Iterator

from comtypes.hresult import S_OK
from powc.core import ComResult, cr, query_interface

from .. import _dwrite
from .types import *


class DWriteLocalizedStrings:
    """DirectWriteローカライズ済み文字列。IDWriteLocalizedStringsインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDWriteLocalizedStrings)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDWriteLocalizedStrings)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def count_nothrow(self) -> ComResult[int]:
        return cr(S_OK, self.__o.GetCount())

    @property
    def count(self) -> int:
        return self.__o.GetCount()

    def find_localename_nothrow(self, localename: str) -> ComResult[tuple[int, bool]]:
        x1 = c_uint32()
        x2 = c_int32()
        return cr(self.__o.FindLocaleName(localename, byref(x1), byref(x2)), (x1.value, x2.value != 0))

    def find_localename(self, localename: str) -> tuple[int, bool]:
        return self.find_localename_nothrow(localename).value

    def get_localenamelen_nothrow(self, index: int) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetLocaleNameLength(index, byref(x)), x.value)

    def get_localenamelen(self, index: int) -> int:
        return self.get_localenamelen_nothrow(index).value

    def get_localename_nothrow(self, index: int) -> ComResult[str]:
        len = c_uint32()
        hr = self.__o.GetLocaleNameLength(index, byref(len))
        if hr != 0:
            return cr(hr, "")
        len.value += 1
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetLocaleName(index, buf, len), buf.value)

    def get_localename(self, index: int) -> str:
        return self.get_localename_nothrow(index).value

    def get_stringlen_nothrow(self, index: int) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetStringLength(index, byref(x)), x.value)

    def get_stringlen(self, index: int) -> int:
        return self.get_stringlen_nothrow(index).value

    def get_string_nothrow(self, index: int) -> ComResult[str]:
        len = c_uint32()
        hr = self.__o.GetStringLength(index, byref(len))
        if hr != 0:
            return cr(hr, "")
        len.value += 1
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetString(index, buf, len), buf.value)

    def get_string(self, index: int) -> str:
        return self.get_string_nothrow(index).value

    @property
    def localename_iter(self) -> Iterator[str]:
        return (self.get_localename(i) for i in range(self.count))

    @property
    def localnames(self) -> tuple[str, ...]:
        return tuple(self.localename_iter)

    @property
    def string_iter(self) -> Iterator[str]:
        return (self.get_string(i) for i in range(self.count))

    @property
    def strings(self) -> tuple[str, ...]:
        return tuple(self.string_iter)


class DWriteFont:
    """DirectWriteフォント。IDWriteFontインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDWriteFont)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDWriteFont)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def fontfamily_nothrow(self) -> "ComResult[DWriteFontFamily]":
        x = POINTER(IDWriteFontFamily)()
        return cr(self.__o.GetFontFamily(byref(x)), DWriteFontFamily(x))

    @property
    def fontfamily(self) -> "DWriteFontFamily":
        return self.fontfamily_nothrow.value

    @property
    def weight_nothrow(self) -> ComResult[DWriteFontWeight]:
        return cr(S_OK, DWriteFontWeight(self.__o.GetWeight()))

    @property
    def weight(self) -> DWriteFontWeight:
        return DWriteFontWeight(self.__o.GetWeight())

    @property
    def stretch_nothrow(self) -> ComResult[DWriteFontStretch]:
        return cr(S_OK, DWriteFontStretch(self.__o.GetStretch()))

    @property
    def stretch(self) -> DWriteFontStretch:
        return DWriteFontStretch(self.__o.GetStretch())

    @property
    def style_nothrow(self) -> ComResult[DWriteFontStyle]:
        return cr(S_OK, DWriteFontStyle(self.__o.GetStretch()))

    @property
    def style(self) -> DWriteFontStyle:
        return DWriteFontStyle(self.__o.GetStretch())

    @property
    def is_symbolfont_nothrow(self) -> ComResult[bool]:
        return cr(S_OK, self.__o.IsSymbolFont() != 0)

    @property
    def is_symbolfont(self) -> bool:
        return self.__o.IsSymbolFont() != 0

    @property
    def facenames_nothrow(self) -> ComResult[DWriteLocalizedStrings]:
        x = POINTER(IDWriteLocalizedStrings)()
        return cr(self.__o.GetFaceNames(byref(x)), DWriteLocalizedStrings(x))

    @property
    def facenames(self) -> DWriteLocalizedStrings:
        return self.facenames_nothrow.value


# TODO
# IDWriteFont._methods_ = [
#     STDMETHOD(
#         c_int32, "GetInformationalStrings", (c_int32, POINTER(POINTER(IDWriteLocalizedStrings)), POINTER(c_int32))
#     ),
#     STDMETHOD(c_int32, "GetSimulations", ()),
#     STDMETHOD(c_int32, "GetMetrics", (POINTER(DWRITE_FONT_METRICS),)),
#     STDMETHOD(c_int32, "HasCharacter", (c_uint32, POINTER(c_int32))),
#     STDMETHOD(c_int32, "CreateFontFace", (POINTER(POINTER(IDWriteFontFace)),)),
# ]


class DWriteFontFace:
    """DirectWriteフォントフェイス。IDWriteFontFaceインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDWriteFontFace)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDWriteFontFace)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o


# TODO
# IDWriteFontFace._methods_ = [
#     STDMETHOD(c_int32, "GetType", ()),  # DWRITE_FONT_FACE_TYPE
#     STDMETHOD(c_int32, "GetFiles", (POINTER(c_uint32), POINTER(POINTER(IDWriteFontFile)))),
#     STDMETHOD(c_uint32, "GetIndex", ()),
#     STDMETHOD(c_int32, "GetSimulations", ()),  # DWRITE_FONT_SIMULATIONS
#     STDMETHOD(c_int32, "IsSymbolFont", ()),
#     STDMETHOD(None, "GetMetrics", (POINTER(DWRITE_FONT_METRICS),)),
#     STDMETHOD(c_uint16, "GetGlyphCount", ()),
#     STDMETHOD(c_int32, "GetDesignGlyphMetrics", (POINTER(c_uint16), c_uint32, POINTER(DWRITE_GLYPH_METRICS), c_int32)),
#     STDMETHOD(c_int32, "GetGlyphIndices", (POINTER(c_uint32), c_uint32, POINTER(c_uint16))),
#     STDMETHOD(
#         c_int32,
#         "TryGetFontTable",
#         (c_uint32, POINTER(c_void_p), POINTER(c_uint32), POINTER(c_void_p), POINTER(c_int32)),
#     ),
#     STDMETHOD(None, "ReleaseFontTable", (c_void_p)),
#     STDMETHOD(
#         c_int32,
#         "GetGlyphRunOutline",
#         (
#             c_float,
#             POINTER(c_float),
#             POINTER(DWRITE_GLYPH_OFFSET),
#             c_uint32,
#             c_int32,
#             c_int32,
#             POINTER(IDWriteGeometrySink),
#         ),
#     ),
#     STDMETHOD(
#         c_int32,
#         "GetRecommendedRenderingMode",
#         (c_float, c_float, c_int32, POINTER(IDWriteRenderingParams), POINTER(c_int32)),
#     ),
#     STDMETHOD(
#         c_int32, "GetGdiCompatibleMetrics", (c_float, c_float, POINTER(DWRITE_MATRIX), POINTER(DWRITE_FONT_METRICS))
#     ),
#     STDMETHOD(
#         c_int32,
#         "GetGdiCompatibleGlyphMetrics",
#         (
#             c_float,
#             c_float,
#             POINTER(DWRITE_MATRIX),
#             c_int32,
#             POINTER(c_uint16),
#             c_uint32,
#             POINTER(DWRITE_GLYPH_METRICS),
#             c_int32,
#         ),
#     ),
# ]


class DWriteFontCollection:
    """DirectWriteフォントコレクション。IDWriteFontCollectionインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDWriteFontCollection)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDWriteFontCollection)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def fontfamily_count_nothrow(self) -> ComResult[int]:
        return cr(S_OK, self.__o.GetFontFamilyCount())

    @property
    def fontfamily_count(self) -> int:
        return self.__o.GetFontFamilyCount()

    def get_fontfamily_nothrow(self, index: int) -> "ComResult[DWriteFontFamily]":
        x = POINTER(IDWriteFontFamily)()
        return cr(self.__o.GetFontFamily(index, byref(x)), DWriteFontFamily(x))

    def get_fontfamily(self, index: int) -> "DWriteFontFamily":
        return self.get_fontfamily_nothrow(index).value

    @property
    def fontfamily_iter(self) -> "Iterator[DWriteFontFamily]":
        get_fontfamily = self.get_fontfamily
        return (get_fontfamily(i) for i in range(self.fontfamily_count))

    @property
    def fontfamilies(self) -> "tuple[DWriteFontFamily, ...]":
        return tuple(self.fontfamily_iter)

    def find_familyname_nothrow(self, familyname: str) -> ComResult[tuple[int, bool]]:
        x1 = c_uint32()
        x2 = c_int32()
        return cr(self.__o.FindFamilyName(familyname, byref(x1), byref(x2)), (x1.value, x2.value != 0))

    def find_familyname(self, familyname: str) -> tuple[int, bool]:
        return self.find_familyname_nothrow(familyname).value

    def get_font_from_fontface_nothrow(self, fontface: DWriteFontFace) -> "ComResult[DWriteFont]":
        x = POINTER(IDWriteFont)()
        return cr(self.__o.GetFontFromFontFace(fontface.wrapped_obj, byref(x)), DWriteFont(x))

    def get_font_from_fontface(self, fontface: DWriteFontFace) -> "DWriteFont":
        return self.get_font_from_fontface_nothrow(fontface).value


class DWriteFontList:
    """DirectWriteフォントリスト（順序付きコレクション）。IDWriteFontListインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDWriteFontList)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDWriteFontList)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def as_collection_nothrow(self) -> ComResult[DWriteFontCollection]:
        x = POINTER(IDWriteFontCollection)()
        return cr(self.__o.GetFontCollection(byref(x)), DWriteFontCollection(x))

    def as_collection(self) -> DWriteFontCollection:
        return self.as_collection_nothrow().value

    @property
    def count_nothrow(self) -> ComResult[int]:
        return cr(S_OK, self.__o.GetFontCount())

    @property
    def count(self) -> int:
        return self.__o.GetFontCount()

    def get_font_nothrow(self, index: int) -> ComResult[DWriteFont]:
        x = POINTER(IDWriteFont)()
        return cr(self.__o.GetFont(index, byref(x)), DWriteFont(x))

    def get_font(self, index: int) -> DWriteFont:
        return self.get_font_nothrow(index).value

    @property
    def font_iter(self) -> Iterator[DWriteFont]:
        get_font = self.get_font
        return (get_font(i) for i in range(self.count))

    @property
    def fonts(self) -> tuple[DWriteFont, ...]:
        return tuple(self.font_iter)


class DWriteFontFamily(DWriteFontList):
    """DirectWriteフォントファミリ。IDWriteFontFamilyインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDWriteFontFamily)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IDWriteFontFamily)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def familynames_nothrow(self) -> ComResult[DWriteLocalizedStrings]:
        x = POINTER(IDWriteLocalizedStrings)()
        return cr(self.__o.GetFamilyNames(byref(x)), DWriteLocalizedStrings(x))

    @property
    def familynames(self) -> DWriteLocalizedStrings:
        return self.familynames_nothrow.value


_DWriteCreateFactory = _dwrite.DWriteCreateFactory
_DWriteCreateFactory.argtypes = (c_int32, POINTER(GUID), POINTER(POINTER(IUnknown)))


class DWriteFactory:
    """DirectWriteファクトリー。IDWriteFactoryインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IIDWriteFactory)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDWriteFactory)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_nothrow(factorytype: DWriteFactoryType) -> "ComResult[DWriteFactory]":
        x = POINTER(IDWriteFactory)()
        return cr(_DWriteCreateFactory(int(factorytype), IDWriteFactory._iid_, byref(x)), DWriteFactory(x))

    @staticmethod
    def create(factorytype: DWriteFactoryType) -> "DWriteFactory":
        return DWriteFactory.create_nothrow(factorytype).value

    @property
    def systemfontcollection_nothrow(self) -> ComResult[DWriteFontCollection]:
        x = POINTER(IDWriteFontCollection)()
        return cr(self.__o.GetSystemFontCollection(byref(x), 0), DWriteFontCollection(x))

    @property
    def systemfontcollection(self) -> DWriteFontCollection:
        return self.systemfontcollection_nothrow.value


#     STDMETHOD(c_int32, "", (POINTER(IDWriteFontCollection), c_int32)),
#     STDMETHOD(
#         c_int32,
#         "CreateCustomFontCollection",
#         (POINTER(IDWriteFontCollectionLoader), c_void_p, c_uint32, POINTER(POINTER(IDWriteFontCollection))),
#     ),
#     STDMETHOD(c_int32, "RegisterFontCollectionLoader", (POINTER(IDWriteFontCollectionLoader),)),
#     STDMETHOD(c_int32, "UnregisterFontCollectionLoader", (POINTER(IDWriteFontCollectionLoader),)),
#     STDMETHOD(c_int32, "CreateFontFileReference", (c_wchar_p, POINTER(FILETIME), POINTER(POINTER(IDWriteFontFile)))),
#     STDMETHOD(
#         c_int32,
#         "CreateCustomFontFileReference",
#         (c_void_p, c_uint32, POINTER(IDWriteFontFileLoader), POINTER(POINTER(IDWriteFontFile))),
#     ),
#     STDMETHOD(
#         c_int32,
#         "CreateFontFace",
#         (
#             c_int32,
#             c_uint32,
#             POINTER(POINTER(IDWriteFontFile)),
#             c_uint32,
#             c_int32,
#             POINTER(POINTER(IDWriteFontFace)),
#         ),
#     ),
#     STDMETHOD(c_int32, "CreateRenderingParams", (POINTER(POINTER(IDWriteRenderingParams)),)),
#     STDMETHOD(
#         c_int32,
#         "CreateMonitorRenderingParams",
#         (
#             c_void_p,
#             POINTER(POINTER(IDWriteRenderingParams)),
#         ),
#     ),
#     STDMETHOD(
#         c_int32,
#         "CreateCustomRenderingParams",
#         (c_float, c_float, c_float, c_int32, c_int32, POINTER(POINTER(IDWriteRenderingParams))),
#     ),
#     STDMETHOD(c_int32, "RegisterFontFileLoader", (POINTER(IDWriteFontFileLoader),)),
#     STDMETHOD(c_int32, "UnRegisterFontFileLoader", (POINTER(IDWriteFontFileLoader),)),
#     STDMETHOD(
#         c_int32,
#         "CreateTextFormat",
#         (
#             c_wchar_p,
#             POINTER(IDWriteFontCollection),
#             c_int32,
#             c_int32,
#             c_int32,
#             c_float,
#             c_wchar_p,
#             POINTER(POINTER(IDWriteTextFormat)),
#         ),
#     ),
#     STDMETHOD(c_int32, "CreateTypography", (POINTER(POINTER(IDWriteTypography)),)),
#     STDMETHOD(c_int32, "GetGdiInterop", (POINTER(POINTER(IDWriteGdiInterop)),)),
#     STDMETHOD(
#         c_int32,
#         "CreateTextLayout",
#         (c_wchar_p, c_uint32, POINTER(IDWriteTextFormat), c_float, c_float, POINTER(IDWriteTextLayout)),
#     ),
#     STDMETHOD(
#         c_int32,
#         "CreateGdiCompatibleTextLayout",
#         (
#             c_wchar_p,
#             c_uint32,
#             POINTER(IDWriteTextFormat),
#             c_float,
#             c_float,
#             c_float,
#             POINTER(DWRITE_MATRIX),
#             c_int32,
#             POINTER(POINTER(IDWriteTextLayout)),
#         ),
#     ),
#     STDMETHOD(
#         c_int32, "CreateEllipsisTrimmingSign", (POINTER(IDWriteTextFormat), POINTER(POINTER(IDWriteInlineObject)))
#     ),
#     STDMETHOD(c_int32, "CreateTextAnalyzer", (POINTER(POINTER(IDWriteTextAnalyzer)),)),
#     STDMETHOD(
#         c_int32,
#         "CreateNumberSubstitution",
#         (c_int32, c_wchar_p, c_int32, POINTER(POINTER(IDWriteNumberSubstitution))),
#     ),
#     STDMETHOD(
#         c_int32,
#         "CreateGlyphRunAnalysis",
#         (
#             POINTER(c_int32),
#             c_float,
#             POINTER(DWRITE_MATRIX),
#             c_int32,
#             c_int32,
#             c_float,
#             c_float,
#             POINTER(POINTER(IDWriteGlyphRunAnalysis)),
#         ),
#     ),
# ]
