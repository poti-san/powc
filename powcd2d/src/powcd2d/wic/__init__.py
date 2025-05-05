# https://learn.microsoft.com/ja-jp/windows/win32/wic/-wic-about-windows-imaging-codec

# 単一ファイルにまとめるには大きすぎるので複数ファイルに分離しています。
# star importsに伴う警告は抑制しています。

from ctypes import (
    POINTER,
    Array,
    byref,
    c_byte,
    c_int32,
    c_uint32,
    c_void_p,
    c_wchar,
    cast,
    sizeof,
)
from typing import Any, Iterator, Sequence, override

from comtypes import GUID, CoCreateInstance
from powc.comobj import (
    ComStringEnumerator,
    IEnumString,
    IEnumUnknown,
    IUnknownEnumerator,
)
from powc.core import ComResult, cr, query_interface
from powc.propbag import PropertyBag2
from powc.stream import ComStream
from powcpropsys.propvariant import PropVariant

from .constant import *
from .types import *


class WICPalette:
    """WICパレット。IWICPaletteインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICPalette)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICPalette)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def initilaize_predefined_nothrow(
        self, palette_type: WICBitmapPaletteType | int, adds_transparent_color: bool
    ) -> ComResult[None]:
        return self.__o.InitializePredefined(int(palette_type), 1 if adds_transparent_color else 0, None)

    def initilaize_predefined(self, palette_type: WICBitmapPaletteType | int, adds_transparent_color: bool) -> None:
        return self.initilaize_predefined_nothrow(palette_type, adds_transparent_color).value

    def initialize_custom_nothrow(self, colors: Sequence[WICColor]) -> ComResult[None]:
        x = (c_uint32 * len(colors))(colors)
        return self.__o.InitializeCustom(x, len(colors))

    def initialize_custom(self, colors: Sequence[WICColor]) -> None:
        return self.initialize_custom_nothrow(colors).value

    def initialize_from_bitmap_nothrow(
        self, surface: "WICBitmap", count: int, adds_transparent_color: bool
    ) -> ComResult[None]:
        return self.__o.InitializeFromBitmap(surface.wrapped_obj, count, 1 if adds_transparent_color else 0)

    def initialize_from_bitmap(self, surface: "WICBitmap", count: int, adds_transparent_color: bool) -> None:
        return self.initialize_from_bitmap_nothrow(surface, count, adds_transparent_color).value

    def initialize_from_palette_nothrow(self, palette: "WICPalette") -> ComResult[None]:
        return cr(self.__o.InitializeFromPalette(byref(palette.wrapped_obj)), None)

    def initialize_from_palette(self, palette: "WICPalette") -> None:
        return self.initialize_from_palette_nothrow(palette).value

    @property
    def type_nothrow(self) -> ComResult[WICBitmapPaletteType]:
        x = c_int32()
        return cr(self.__o.GetType(byref(x)), WICBitmapPaletteType(x.value))

    @property
    def type(self) -> WICBitmapPaletteType:
        return self.type_nothrow.value

    @property
    def colorcount_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetType(byref(x)), x.value)

    @property
    def colorcount(self) -> int:
        return self.colorcount_nothrow.value

    @property
    def colors_nothrow(self) -> ComResult[tuple[WICColor, ...]]:
        count = c_uint32()
        self.__o.GetColors(0, None, byref(count))
        buf = (WICColor * count.value)()
        return cr(self.__o.GetColors(count.value, buf, byref(count)), tuple(buf))

    @property
    def is_blackwhite_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.IsBlackWhite(byref(x)), x.value != 0)

    @property
    def is_blackwhite(self) -> bool:
        return self.is_blackwhite_nothrow.value

    @property
    def is_grayscale_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.IsGrayscale(byref(x)), x.value != 0)

    @property
    def is_grayscale(self) -> bool:
        return self.is_grayscale_nothrow.value

    @property
    def has_alpha_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.HasAlpha(byref(x)), x.value != 0)

    @property
    def has_alpha(self) -> bool:
        return self.has_alpha_nothrow.value


class WICBitmapSource:
    """ビットマップソース。IWICBitmapSourceインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapSource)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICBitmapSource)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def size_nothrow(self) -> ComResult[tuple[int, int]]:
        x1 = c_uint32()
        x2 = c_uint32()
        return cr(self.__o.GetSize(byref(x1), byref(x2)), (x1.value, x2.value))

    @property
    def size(self) -> tuple[int, int]:
        return self.size_nothrow.value

    @property
    def pixelformat_nothrow(self) -> ComResult[WICPixelFormatGUID]:
        x = WICPixelFormatGUID()
        return cr(self.__o.GetPixelFormat(byref(x)), x)

    @property
    def pixelformat(self) -> WICPixelFormatGUID:
        return self.pixelformat_nothrow.value

    @property
    def resolution_nothrow(self) -> ComResult[tuple[float, float]]:
        x1 = c_double()
        x2 = c_double()
        return cr(self.__o.GetResolution(byref(x1), byref(x2)), (x1.value, x2.value))

    @property
    def resolution(self) -> tuple[float, float]:
        return self.resolution_nothrow.value

    def copy_palette_nothrow(self, dest_palette: WICPalette) -> ComResult[None]:
        return cr(self.__o.CopyPalette(dest_palette.wrapped_obj), None)

    def copy_palette(self, dest_palette: WICPalette) -> None:
        return self.copy_palette_nothrow(dest_palette).value

    def copy_pixels_nothrow(self, rc: WICRect, stride: int, buffer_size: int) -> ComResult[Array[c_byte]]:
        x = (c_byte * buffer_size)()
        return cr(self.__o.CopyPixels(rc, stride, buffer_size, x), x)

    def copy_pixels(self, rc: WICRect, stride: int, buffer_size: int) -> Array[c_byte]:
        return self.copy_pixels_nothrow(rc, stride, buffer_size).value


# class IWICFormatConverter(IWICBitmapSource):
#     __slots__ = ()
#     _iid_ = GUID("{00000301-a8f2-4877-ba0a-fd2b6645fb94}")
#     _methods_ = [
#         STDMETHOD(
#             c_int32,
#             "Initialize",
#             (POINTER(IWICBitmapSource), POINTER(WICPixelFormatGUID), c_int32, POINTER(IWICPalette), c_double, c_int32),
#         ),
#         STDMETHOD(c_int32, "CanConvert", (POINTER(WICPixelFormatGUID), POINTER(WICPixelFormatGUID), POINTER(c_int32))),
#     ]


# class IWICPlanarFormatConverter(IWICBitmapSource):
#     __slots__ = ()
#     _iid_ = GUID("{BEBEE9CB-83B0-4DCC-8132-B0AAA55EAC96}")
#     _methods_ = [
#         STDMETHOD(
#             c_int32,
#             "Initialize",
#             (
#                 POINTER(POINTER(IWICBitmapSource)),
#                 c_uint32,
#                 POINTER(WICPixelFormatGUID),
#                 c_int32,
#                 POINTER(IWICPalette),
#                 c_double,
#                 c_int32,
#             ),
#         ),
#         STDMETHOD(c_int32, "CanConvert", (POINTER(WICPixelFormatGUID), c_uint32, POINTER(GUID), POINTER(c_int32))),
#     ]


# class IWICBitmapScaler(IWICBitmapSource):
#     __slots__ = ()
#     _iid_ = GUID("{00000302-a8f2-4877-ba0a-fd2b6645fb94}")
#     _methods_ = [
#         STDMETHOD(c_int32, "Initialize", (POINTER(IWICBitmapSource), c_uint32, c_uint32, c_int32)),
#     ]


# class IWICBitmapClipper(IWICBitmapSource):
#     __slots__ = ()
#     _iid_ = GUID("{E4FBCF03-223D-4e81-9333-D635556DD1B5}")
#     _methods_ = [
#         STDMETHOD(c_int32, "Initialize", (POINTER(IWICBitmapSource), POINTER(WICRect))),
#     ]


# class IWICBitmapFlipRotator(IWICBitmapSource):
#     __slots__ = ()
#     _fields_ = GUID("{5009834F-2D6A-41ce-9E1B-17C5AFF7A782}")
#     _methods_ = [
#         STDMETHOD(c_int32, "Initialize", (POINTER(IWICBitmapSource), c_int32)),
#     ]


class WICBitmapLock:
    """固定されたビットマップ領域。IWICBitmapLockインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapLock)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICBitmapLock)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def size_nothrow(self) -> ComResult[tuple[int, int]]:
        x1 = c_uint32()
        x2 = c_uint32()
        return cr(self.__o.GetSize(byref(x1), byref(x2)), (x1.value, x2.value))

    @property
    def size(self) -> tuple[int, int]:
        return self.size_nothrow.value

    @property
    def stride_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetStride(byref(x)), x.value)

    @property
    def stride(self) -> int:
        return self.stride_nothrow.value

    @property
    def datapointer_nothrow(self) -> ComResult[tuple[int, int]]:
        x1 = c_uint32()
        x2 = POINTER(c_byte)()
        return cr(self.__o.GetDataPointer(byref(x1), byref(x2)), (x1.value, cast(x2, c_void_p).value or 0))

    @property
    def datapointer(self) -> tuple[int, int]:
        return self.datapointer_nothrow.value

    @property
    def pixelformat_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetPixelFormat(byref(x)), x)

    @property
    def pixelformat(self) -> GUID:
        return self.pixelformat_nothrow.value


class WICBitmap(WICBitmapSource):
    """ビットマップ。IWICBitmapインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmap)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IWICBitmap)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def lock_nothrow(self, rc: WICRect, flags: WICBitmapLockFlag | int) -> ComResult[WICBitmapLock]:
        x = POINTER(IWICBitmapLock)()
        return cr(self.__o.Lock(rc, int(flags), byref(x)), WICBitmapLock(x))

    def lock(self, rc: WICRect, flags: WICBitmapLockFlag | int) -> WICBitmapLock:
        return self.lock_nothrow(rc, flags).value

    def set_palette_nothrow(self, palette: WICPalette) -> ComResult[None]:
        return cr(self.__o.SetPalette(palette.wrapped_obj), None)

    def set_palette(self, palette: WICPalette) -> None:
        return self.set_palette_nothrow(palette).value

    @property
    def resolution_nothrow(self) -> ComResult[tuple[float, float]]:
        return super().resolution_nothrow

    @property
    def resolution(self) -> tuple[float, float]:
        return super().resolution

    def set_resolution_nothrow(self, value: tuple[float, float]) -> ComResult[None]:
        return self.__o.SetResolution(value[0], value[1])

    @resolution.setter
    def resolution(self, value: tuple[float, float]) -> None:
        return self.set_resolution_nothrow(value).value


# class IWICColorContext(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{3C613A02-34B2-44ea-9A7C-45AEA9C6FD6D}")
#     _methods_ = [
#         STDMETHOD(c_int32, "InitializeFromFilename", (c_wchar_p,)),
#         STDMETHOD(c_int32, "InitializeFromMemory", (POINTER(c_byte), c_uint32)),
#         STDMETHOD(c_int32, "InitializeFromExifColorSpace", (c_uint32,)),
#         STDMETHOD(c_int32, "GetType", (c_int32,)),  # WICColorContextType
#         STDMETHOD(c_int32, "GetProfileBytes", (c_uint32, POINTER(c_byte), POINTER(c_uint32))),
#         STDMETHOD(c_int32, "GetExifColorSpace", (POINTER(c_uint32),)),
#     ]


# class IWICColorTransform(IWICBitmapSource):
#     __slots__ = ()
#     _iid_ = GUID("{B66F034F-D0E2-40ab-B436-6DE39E321A94}")
#     _methods_ = [
#         STDMETHOD(
#             c_int32,
#             "Initialize",
#             (
#                 POINTER(IWICBitmapSource),
#                 POINTER(IWICColorContext),
#                 POINTER(IWICColorContext),
#                 POINTER(WICPixelFormatGUID),
#             ),
#         ),
#     ]


# class IWICFastMetadataEncoder(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{B84E2C09-78C9-4AC4-8BD3-524AE1663A2F}")


# class IWICStream(IStream):
#     __slots__ = ()
#     _iid_ = GUID("{135FF860-22B7-4ddf-B0F6-218F4F299A43}")
#     _methods_ = [
#         STDMETHOD(c_int32, "InitializeFromIStream", (POINTER(IStream),)),
#         STDMETHOD(c_int32, "InitializeFromFilename", (c_wchar_p, c_uint32)),
#         STDMETHOD(c_int32, "InitializeFromMemory", (c_void_p, c_uint32)),
#         STDMETHOD(c_int32, "InitializeFromIStreamRegion", (POINTER(IStream), c_uint64, c_uint64)),
#     ]


# class IWICEnumMetadataItem(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{DC2BB46D-3F07-481E-8625-220C4AEDBB33}")


# IWICEnumMetadataItem._methods_ = [
#     STDMETHOD(c_int32, "Next", (POINTER(PropVariant), POINTER(PropVariant), POINTER(PropVariant), POINTER(c_uint32))),
#     STDMETHOD(c_int32, "Skip", (c_uint32,)),
#     STDMETHOD(c_int32, "Reset", ()),
#     STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IWICEnumMetadataItem)),)),
# ]


# class IWICMetadataQueryReader(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{30989668-E1C9-4597-B395-458EEDB808DF}")
#     _methods_ = [
#         STDMETHOD(c_int32, "GetContainerFormat", (POINTER(GUID),)),
#         STDMETHOD(c_int32, "GetLocation", (c_uint32, c_wchar_p, POINTER(c_uint32))),
#         STDMETHOD(c_int32, "GetMetadataByName", (c_wchar_p, POINTER(PropVariant))),
#         STDMETHOD(c_int32, "GetEnumerator", (POINTER(IEnumString),)),
#     ]


# class IWICMetadataQueryWriter(IWICMetadataQueryReader):
#     __slots__ = ()
#     _iid_ = GUID("{A721791A-0DEF-4d06-BD91-2118BF1DB10B}")
#     _methods_ = [
#         STDMETHOD(c_int32, "SetMetadataByName", (c_wchar_p, POINTER(PropVariant))),
#         STDMETHOD(c_int32, "RemoveMetadataByName", (c_wchar_p,)),
#     ]


# IWICFastMetadataEncoder._methods_ = [
#     STDMETHOD(c_int32, "Commit", ()),
#     STDMETHOD(c_int32, "GetMetadataQueryWriter", (POINTER(IWICMetadataQueryWriter),)),
# ]


class WICBitmapFrameDecode(WICBitmapSource):
    """デコーダーの個々のイメージフレーム。IWICBitmapFrameDecodeインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapFrameDecode)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IWICBitmapFrameDecode)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def get_metadataqueryreader_nothtow(self) -> "ComResult[WICMetadataQueryReader]":
        x = POINTER(IWICMetadataQueryReader)()
        return cr(self.__o.GetMetadataQueryReader(byref(x)), WICMetadataQueryReader(x))

    def get_metadataqueryreader(self) -> "WICMetadataQueryReader":
        return self.get_metadataqueryreader_nothtow().value


# TODO
#         STDMETHOD(c_int32, "GetColorContexts", (c_uint32, POINTER(POINTER(IWICColorContext)), POINTER(c_uint32))),
#         STDMETHOD(c_int32, "GetThumbnail", (POINTER(POINTER(IWICBitmapSource)),)),
#     ]


class WICBitmapFrameEncode:
    """エンコーダーの個々のイメージフレーム。IWICBitmapFrameEncodeインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapFrameEncode)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICBitmapFrameEncode)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    # TODO: TEST
    def initialize_nothrow(self, propbag: PropertyBag2 | None) -> ComResult[None]:
        return cr(self.__o.Initialize(propbag.wrapped_obj if propbag else None), None)

    # TODO: TEST
    def initialize(self, propbag: PropertyBag2 | None) -> None:
        return self.initialize_nothrow(propbag).value


# TODO
#         STDMETHOD(c_int32, "SetSize", (c_uint32, c_uint32)),
#         STDMETHOD(c_int32, "SetResolution", (c_double, c_double)),
#         STDMETHOD(c_int32, "SetPixelFormat", (POINTER(WICPixelFormatGUID),)),
#         STDMETHOD(c_int32, "SetColorContexts", (c_uint32, POINTER(IWICColorContext))),
#         STDMETHOD(c_int32, "SetPalette", (POINTER(IWICPalette),)),
#         STDMETHOD(c_int32, "SetThumbnail", (POINTER(IWICBitmapSource),)),
#         STDMETHOD(c_int32, "WritePixels", (c_uint32, c_uint32, c_uint32, POINTER(c_byte))),
#         STDMETHOD(c_int32, "WriteSource", (POINTER(IWICBitmapSource), POINTER(WICRect))),
#         STDMETHOD(c_int32, "Commit", ()),
#         STDMETHOD(c_int32, "GetMetadataQueryWriter", (POINTER(IWICMetadataQueryWriter),)),


# class IWICPlanarBitmapFrameEncode(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{F928B7B8-2221-40C1-B72E-7E82F1974D1A}")
#     _methods_ = [
#         STDMETHOD(c_int32, "WritePixels", (c_uint32, POINTER(WICBitmapPlane), c_uint32)),
#         STDMETHOD(c_int32, "WriteSource", (POINTER(POINTER(IWICBitmapSource)), c_uint32, POINTER(WICRect))),
#     ]


# class IWICImageEncoder(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{04C75BF8-3CE1-473B-ACC5-3CC4F5E94999}")
#     _methods_ = [
#         STDMETHOD(
#             c_int32, "WriteFrame", (POINTER(ID2D1Image), POINTER(IWICBitmapFrameEncode), POINTER(WICImageParameters))
#         ),
#         STDMETHOD(
#             c_int32,
#             "WriteFrameThumbnail",
#             (POINTER(ID2D1Image), POINTER(IWICBitmapFrameEncode), POINTER(WICImageParameters)),
#         ),
#         STDMETHOD(
#             c_int32,
#             "WriteThumbnail",
#             (POINTER(ID2D1Image), POINTER(IWICBitmapEncoder), POINTER(WICImageParameters)),
#         ),
#     ]


# class IWICBitmapDecoder(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{9EDDE9E7-8DEE-47ea-99DF-E6FAF2ED44BF}")


# class IWICBitmapSourceTransform(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{3B16811B-6A43-4ec9-B713-3D5A0C13B940}")
#     _methods_ = [
#         STDMETHOD(
#             c_int32,
#             "CopyPixels",
#             (POINTER(WICRect), c_uint32, c_uint32, POINTER(WICPixelFormatGUID), c_int32, c_uint32, c_uint32, c_void_p),
#         ),
#         STDMETHOD(c_int32, "GetClosestSize", (POINTER(c_uint32), POINTER(c_uint32))),
#         STDMETHOD(c_int32, "GetClosestPixelFormat", (POINTER(WICPixelFormatGUID),)),
#         STDMETHOD(c_int32, "DoesSupportTransform", (c_int32, POINTER(c_int32))),
#         STDMETHOD(c_int32, "", ()),
#     ]


# class IWICPlanarBitmapSourceTransform(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{3AFF9CCE-BE95-4303-B927-E7D16FF4A613}")
#     _methods_ = [
#         STDMETHOD(
#             c_int32,
#             "DoesSupportTransform",
#             (
#                 POINTER(c_uint32),
#                 POINTER(c_uint32),
#                 c_int32,
#                 c_int32,
#                 POINTER(WICPixelFormatGUID),
#                 POINTER(WICBitmapPlaneDescription),
#                 c_uint32,
#                 POINTER(c_int32),
#             ),
#         ),
#         STDMETHOD(
#             c_int32,
#             "CopyPixels",
#             (POINTER(WICRect), c_uint32, c_uint32, c_int32, c_int32, POINTER(WICBitmapPlane), c_uint32),
#         ),
#     ]


# class IWICBitmapFrameDecode(IWICBitmapSource):
#     __slots__ = ()
#     _iid_ = GUID("{3B16811B-6A43-4ec9-A813-3D930C13B940}")
#     _methods_ = [
#         STDMETHOD(c_int32, "GetMetadataQueryReader", (POINTER(POINTER(IWICMetadataQueryReader)),)),
#         STDMETHOD(c_int32, "GetColorContexts", (c_uint32, POINTER(POINTER(IWICColorContext)), POINTER(c_uint32))),
#         STDMETHOD(c_int32, "GetThumbnail", (POINTER(POINTER(IWICBitmapSource)),)),
#     ]


# class IWICProgressiveLevelControl(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{DAAC296F-7AA5-4dbf-8D15-225C5976F891}")
#     _methods_ = [
#         STDMETHOD(c_int32, "GetLevelCount", (POINTER(c_uint32),)),
#         STDMETHOD(c_int32, "GetCurrentLevel", (POINTER(c_uint32),)),
#         STDMETHOD(c_int32, "SetCurrentLevel", (c_uint32,)),
#     ]


# class IWICDisplayAdaptationControl(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{de9d91d2-70b4-4f41-836c-25fcd39626d3}")
#     _methods_ = [
#         STDMETHOD(c_int32, "DoesSupportChangingMaxLuminance", (POINTER(WICPixelFormatGUID), POINTER(c_int32))),
#         STDMETHOD(c_int32, "SetDisplayMaxLuminance", (c_float,)),
#         STDMETHOD(c_int32, "GetDisplayMaxLuminance", (POINTER(c_float),)),
#     ]


# class IWICProgressCallback(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{4776F9CD-9517-45FA-BF24-E89C5EC5C60C}")
#     _methods_ = [
#         STDMETHOD(c_int32, "Notify", (c_uint32, c_int32, c_double)),
#     ]


# PFNProgressNotification = WINFUNCTYPE(c_int32, c_void_p, c_uint32, c_int32, c_double)


# class IWICBitmapCodecProgressNotification(IUnknown):
#     __slots__ = ()
#     _iid_ = GUID("{64C1024E-C3CF-4462-8078-88C2B11C46D9}")
#     _methods_ = [
#         STDMETHOD(c_int32, "RegisterProgressNotification", (PFNProgressNotification, c_void_p, c_uint32)),
#     ]


#######################################


#
# COMインターフェイスラッパー
#


class WICComponentInfo:
    """WICコンポーネント情報。IWICComponentInfoインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICComponentInfo)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICComponentInfo)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def componenttype_nothrow(self) -> ComResult[WICComponentType]:
        x = c_int32()
        return cr(self.__o.GetComponentType(byref(x)), WICComponentType(x.value))

    @property
    def componenttype(self) -> WICComponentType:
        return self.componenttype_nothrow.value

    @property
    def clsid_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetCLSID(byref(x)), x)

    @property
    def clsid(self) -> GUID:
        return self.clsid_nothrow.value

    @property
    def signingstatus_nothrow(self) -> ComResult[WICComponentSigning]:
        x = c_int32()
        return cr(self.__o.GetSigningStatus(byref(x)), WICComponentSigning(x.value))

    @property
    def signingstatus(self) -> WICComponentSigning:
        return self.signingstatus_nothrow.value

    @property
    def author_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetAuthor(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetAuthor(len.value, buf, byref(len)), buf.value)

    @property
    def author(self) -> str:
        return self.author_nothrow.value

    @property
    def vendorguid_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetVendorGUID(byref(x)), x)

    @property
    def vendorguid(self) -> GUID:
        return self.vendorguid_nothrow.value

    @property
    def version_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetVersion(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetVersion(len.value, buf, byref(len)), buf.value)

    @property
    def version(self) -> str:
        return self.version_nothrow.value

    @property
    def specversion_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetSpecVersion(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetSpecVersion(len.value, buf, byref(len)), buf.value)

    @property
    def specversion(self) -> str:
        return self.specversion_nothrow.value

    @property
    def friendlyname_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetFriendlyName(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetFriendlyName(len.value, buf, byref(len)), buf.value)

    @property
    def friendlyname(self) -> str:
        return self.friendlyname_nothrow.value


class WICBitmapCodecInfo(WICComponentInfo):
    """ビットマップエンコーダー情報。IWICBitmapCodecInfoインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapCodecInfo)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IWICBitmapCodecInfo)

    @property
    @override
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def containerformat_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetContainerFormat(byref(x)), x)

    @property
    def containerformat(self) -> GUID:
        return self.containerformat_nothrow.value

    @property
    def pixelformats_nothrow(self) -> ComResult[tuple[GUID, ...]]:
        size = c_uint32()
        self.__o.GetPixelFormats(0, None, size)

        count = size.value // sizeof(GUID)
        guids = (GUID * count)()
        return cr(self.__o.GetPixelFormats(count, guids, size), tuple(guids))

    @property
    def pixelformats(self) -> tuple[GUID, ...]:
        return self.pixelformats_nothrow.value

    @property
    def colormanagementversion_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetColorManagementVersion(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetColorManagementVersion(len.value, buf, byref(len)), buf.value)

    @property
    def colormanagementversion(self) -> str:
        return self.colormanagementversion_nothrow.value

    @property
    def devicemanufacturer_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetDeviceManufacturer(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetDeviceManufacturer(len.value, buf, byref(len)), buf.value)

    @property
    def devicemanufacturer(self) -> str:
        return self.devicemanufacturer_nothrow.value

    @property
    def devicemodels_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetDeviceModels(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetDeviceModels(len.value, buf, byref(len)), buf.value)

    @property
    def devicemodels(self) -> str:
        return self.devicemodels_nothrow.value

    @property
    def mimetypes_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetMimeTypes(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetMimeTypes(len.value, buf, byref(len)), buf.value)

    @property
    def mimetypes(self) -> str:
        return self.mimetypes_nothrow.value

    @property
    def fileextensions_nothrow(self) -> ComResult[str]:
        len = c_uint32()
        self.__o.GetFileExtensions(0, None, byref(len))
        buf = (c_wchar * len.value)()
        return cr(self.__o.GetFileExtensions(len.value, buf, byref(len)), buf.value)

    @property
    def fileextensions(self) -> str:
        return self.fileextensions_nothrow.value

    @property
    def does_supports_animation_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.DoesSupportAnimation(byref(x)), x.value != 0)

    @property
    def does_supports_animation(self) -> bool:
        return self.does_supports_animation_nothrow.value

    @property
    def does_supports_chromakey_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.DoesSupportChromakey(byref(x)), x.value != 0)

    @property
    def does_supports_chromakey(self) -> bool:
        return self.does_supports_chromakey_nothrow.value

    @property
    def does_supports_lossless_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.DoesSupportLossless(byref(x)), x.value != 0)

    @property
    def does_supports_lossless(self) -> bool:
        return self.does_supports_lossless_nothrow.value

    @property
    def does_supports_multiframe_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.DoesSupportMultiframe(byref(x)), x.value != 0)

    @property
    def does_supports_multiframe(self) -> bool:
        return self.does_supports_multiframe_nothrow.value

    def matches_mimetype_nothrow(self, mimetype: str) -> ComResult[bool]:
        """MIME種類とコーデックのMIME種類が一致する場合は真を返します。Windows標準のコーデックは常にE_NOTIMPLで失敗します。"""
        x = c_int32()
        return cr(self.__o.MatchesMimeType(mimetype, byref(x)), x.value != 0)

    def matches_mimetype(self, mimetype: str) -> bool:
        """MIME種類とコーデックのMIME種類が一致する場合は真を返します。Windows標準のコーデックは常にE_NOTIMPLで失敗します。"""
        return self.matches_mimetype_nothrow(mimetype).value


class WICBitmapEncoderInfo(WICBitmapCodecInfo):
    """ビットマップエンコーダー情報。IWICBitmapEncoderInfoインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapEncoderInfo)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IWICBitmapEncoderInfo)

    @property
    @override
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def create_instance_nothrow(self) -> "ComResult[WICBitmapEncoder]":
        x = POINTER(IWICBitmapEncoder)()
        return cr(self.__o.CreateInstance(byref(x)), WICBitmapEncoder(x))

    def create_instance(self) -> "WICBitmapEncoder":
        return self.create_instance_nothrow().value


class WICBitmapDecoderInfo(WICBitmapCodecInfo):
    """ビットマップデコーダー情報。IWICBitmapDecoderInfoインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapDecoderInfo)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IWICBitmapDecoderInfo)

    @property
    @override
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    # STDMETHOD(c_int32, "GetPatterns", (c_uint32, POINTER(WICBitmapPattern), POINTER(c_uint32), POINTER(c_uint32))),
    # STDMETHOD(c_int32, "MatchesPattern", (POINTER(IStream), POINTER(c_int32))),

    # def create_instance_nothrow(self) -> "ComResult[WICBitmapDecoder]":
    #     x = POINTER(IWICBitmapDecoder)()
    #     return cr(self.__o.CreateInstance(byref(x)), WICBitmapDecoder(x))

    # def create_instance(self) -> "WICBitmapDecoder":
    #     return self.create_instance_nothrow().value


class WICBitmapEncoder:
    """ビットマップエンコーダー。IWICBitmapEncoderインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapEncoder)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICBitmapEncoder)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def initialize_nothrow(self, stream: ComStream, options: WICBitmapEncoderCacheOption | int) -> ComResult[None]:
        return cr(self.__o.Initialize(stream.wrapped_obj, int(options)), None)

    def initialize(self, stream: ComStream, options: WICBitmapEncoderCacheOption | int) -> None:
        return self.initialize_nothrow(stream, options).value

    @property
    def containerformat_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetContainerFormat(byref(x)), x)

    @property
    def containerformat(self) -> GUID:
        return self.containerformat_nothrow.value

    @property
    def encoderinfo_nothrow(self) -> ComResult[WICBitmapEncoderInfo]:
        x = POINTER(IWICBitmapEncoderInfo)()
        return cr(self.__o.GetEncoderInfo(byref(x)), WICBitmapEncoderInfo(x))

    @property
    def encoderinfo(self) -> WICBitmapEncoderInfo:
        return self.encoderinfo_nothrow.value

    # TODO
    # STDMETHOD(c_int32, "SetColorContexts", (c_uint32, POINTER(IWICColorContext))),
    # STDMETHOD(c_int32, "SetPalette", (POINTER(IWICPalette),)),
    # STDMETHOD(c_int32, "SetThumbnail", (POINTER(IWICBitmapSource),)),
    # STDMETHOD(c_int32, "SetPreview", (POINTER(IWICBitmapSource),)),
    # STDMETHOD(c_int32, "CreateNewFrame", (POINTER(POINTER(IWICBitmapFrameEncode)), POINTER(POINTER(IPropertyBag2)))),

    def commit_nothrow(self) -> ComResult[None]:
        return cr(self.__o.Commit(), None)

    def commit(self) -> None:
        return self.commit_nothrow().value


#     STDMETHOD(c_int32, "GetMetadataQueryWriter", (POINTER(POINTER(IWICMetadataQueryWriter)),)),
# ]


class WICBitmapDecoder:
    """ビットマップデコーダー。IWICBitmapDecoderインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICBitmapDecoder)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICBitmapDecoder)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def query_capability_nothrow(self, stream: ComStream) -> ComResult[WICBitmapDecoderCaps]:
        """ストリームのデコーダー機能を取得します。"""
        x = c_uint32()
        return cr(self.__o.QueryCapability(stream.wrapped_obj, byref(x)), WICBitmapDecoderCaps(x.value))

    def query_capability(self, stream: ComStream) -> WICBitmapDecoderCaps:
        """ストリームのデコーダー機能を取得します。"""
        return self.query_capability_nothrow(stream).value

    def initialize_nothrow(self, stream: ComStream, options: WICDecodeOption | int) -> ComResult[None]:
        return cr(self.__o.Initialize(stream.wrapped_obj, int(options)), None)

    def initialize(self, stream: ComStream, options: WICDecodeOption | int) -> None:
        return self.initialize_nothrow(stream, options).value

    @property
    def containerformat_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetContainerFormat(byref(x)), x)

    @property
    def containerformat(self) -> GUID:
        return self.containerformat_nothrow.value

    @property
    def decoderinfo_nothrow(self) -> ComResult[WICBitmapDecoderInfo]:
        x = POINTER(IWICBitmapDecoderInfo)()
        return cr(self.__o.GetDecoderInfo(byref(x)), WICBitmapDecoderInfo(x))

    @property
    def decoderinfo(self) -> WICBitmapDecoderInfo:
        return self.decoderinfo_nothrow.value

    # TODO
    #     STDMETHOD(c_int32, "CopyPalette", (POINTER(IWICPalette),)),
    #     STDMETHOD(c_int32, "GetMetadataQueryReader", (POINTER(POINTER(IWICMetadataQueryReader)),)),
    #     STDMETHOD(c_int32, "GetPreview", (POINTER(POINTER(IWICBitmapSource)),)),
    #     STDMETHOD(c_int32, "GetColorContexts", (c_uint32, POINTER(POINTER(IWICColorContext)), POINTER(c_uint32))),
    #     STDMETHOD(
    #         c_int32,
    #         "GetThumbnail",
    #         (POINTER(POINTER(IWICBitmapSource)),),
    #     ),
    #     STDMETHOD(c_int32, "GetFrameCount", (POINTER(c_uint32),)),

    @property
    def framecount_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetFrameCount(byref(x)), x.value)

    @property
    def framecount(self) -> int:
        return self.framecount_nothrow.value

    def get_frame_nothrow(self, index: int) -> ComResult[WICBitmapFrameDecode]:
        x = POINTER(IWICBitmapFrameDecode)()
        return cr(self.__o.GetFrame(index, byref(x)), WICBitmapFrameDecode(x))

    def get_frame(self, index: int) -> WICBitmapFrameDecode:
        return self.get_frame_nothrow(index).value


class WICImageEncoder:
    """ID2D1ImageインターフェイスからIWICBitmapEncoderへのエンコード。
    IWICImageEncoderインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICImageEncoder)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICImageEncoder)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def write_frame_nothrow(
        self, image: ID2D1Image, options: WICImageParameters | None = None
    ) -> ComResult[WICBitmapFrameEncode]:
        x = POINTER(IWICBitmapFrameEncode)()
        return cr(self.__o.WriteFrame(image, byref(x), options), WICBitmapFrameEncode(x))

    def write_frame(self, image: ID2D1Image, options: WICImageParameters | None = None) -> WICBitmapFrameEncode:
        return self.write_frame_nothrow(image, options).value

    def write_framethumbnail_nothrow(
        self, image: ID2D1Image, options: WICImageParameters | None = None
    ) -> ComResult[WICBitmapFrameEncode]:
        x = POINTER(IWICBitmapFrameEncode)()
        return cr(self.__o.WriteFrameThumbnail(image, byref(x), options), WICBitmapFrameEncode(x))

    def write_framethumbnail(
        self, image: ID2D1Image, options: WICImageParameters | None = None
    ) -> WICBitmapFrameEncode:
        return self.write_framethumbnail_nothrow(image, options).value

    def write_thumbnail_nothrow(
        self, image: ID2D1Image, options: WICImageParameters | None = None
    ) -> ComResult[WICBitmapEncoder]:
        x = POINTER(IWICBitmapEncoder)()
        return cr(self.__o.WriteThumbnail(image, byref(x), options), WICBitmapEncoder(x))

    def write_thumbnail(self, image: ID2D1Image, options: WICImageParameters | None = None) -> WICBitmapEncoder:
        return self.write_thumbnail_nothrow(image, options).value


class WICMetadataQueryReader:
    """WICのメタデータクエリリーダー。IWICMetadataQueryReaderインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICMetadataQueryReader)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICMetadataQueryReader)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def containerformat_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetContainerFormat(byref(x)), x)

    @property
    def containerformat(self) -> GUID:
        return self.containerformat_nothrow.value

    @property
    def location_nothrow(self) -> ComResult[str]:
        cb = c_uint32()
        self.__o.GetLocation(0, None, byref(cb))

        buf = (c_wchar * cb.value)()
        return cr(self.__o.GetLocation(cb.value, buf, byref(cb)), buf.value)

    def get_metadata_by_name_nothrow(self, name: str) -> ComResult[PropVariant]:
        x = PropVariant()
        return cr(self.__o.GetMetadataByName(name, byref(x)), x)

    def get_metadata_by_name(self, name: str) -> PropVariant:
        return self.get_metadata_by_name_nothrow(name).value

    @property
    def enumerator_nothrow(self) -> ComResult[ComStringEnumerator]:
        x = POINTER(IEnumString)()
        return cr(self.__o.GetEnumerator(byref(x)), ComStringEnumerator(x))

    @property
    def enumerator(self) -> ComStringEnumerator:
        return self.enumerator_nothrow.value


class WICMetadataQueryWriter(WICMetadataQueryReader):
    """WICのメタデータクエリライター。IWICMetadataQueryWriterインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICMetadataQueryWriter)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICMetadataQueryWriter)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def set_metadata_by_name_nothrow(self, name: str, value: PropVariant) -> ComResult[None]:
        return cr(self.__o.SetMetadataByName(name, value), None)

    def set_metadata_by_name(self, name: str, value: PropVariant) -> None:
        return self.set_metadata_by_name_nothrow(name, value).value

    def remove_metadata_by_name_nothrow(self, name: str) -> ComResult[None]:
        return cr(self.__o.RemoveMetadataByName(name), None)

    def remove_metadata_by_name(self, name: str) -> None:
        return self.remove_metadata_by_name_nothrow(name).value


class WICFastMetadataEncoder:
    """WICの高速メタデータエンコーダー。IWICFastMetadataEncoderインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICFastMetadataEncoder)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICFastMetadataEncoder)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def commit_nothrow(self) -> ComResult[None]:
        return cr(self.__o.Commit(), None)

    def commit(self) -> None:
        return self.commit_nothrow().value

    @property
    def metadataquerywriter_nothrow(self) -> ComResult[WICMetadataQueryWriter]:
        x = POINTER(IWICMetadataQueryWriter)()
        return cr(self.__o.GetMetadataQueryWriter(byref(x)), WICMetadataQueryWriter(x))

    @property
    def metadataquerywriter(self) -> WICMetadataQueryWriter:
        return self.metadataquerywriter_nothrow.value


class WICImagingFactory:
    """WICイメージングのファクトリクラス。IWICImagingFactoryインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICImagingFactory)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWICImagingFactory)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "WICImagingFactory":
        return WICImagingFactory(CoCreateInstance(CLSID_WIC_IMAGING_FACTORY, IWICImagingFactory))

    def create_decoder_from_filename_nothrow(
        self,
        filename: str,
        access: FileAccess | int,
        options: WICDecodeOption | int,
        vendor_guid: GUID | None = None,
    ) -> ComResult[WICBitmapDecoder]:
        x = POINTER(IWICBitmapDecoder)()
        return cr(
            self.__o.CreateDecoderFromFilename(filename, vendor_guid, int(access), int(options), byref(x)),
            WICBitmapDecoder(x),
        )

    def create_decoder_from_filename(
        self,
        filename: str,
        access: FileAccess | int,
        options: WICDecodeOption | int,
        vendor_guid: GUID | None = None,
    ) -> WICBitmapDecoder:
        return self.create_decoder_from_filename_nothrow(filename, access, options, vendor_guid).value

    def create_decoder_from_stream_nothrow(
        self,
        stream: ComStream,
        options: WICDecodeOption | int,
        vendor_guid: GUID | None = None,
    ) -> ComResult[WICBitmapDecoder]:
        x = POINTER(IWICBitmapDecoder)()
        return cr(
            self.__o.CreateDecoderFromStream(stream.wrapped_obj, vendor_guid, int(options), byref(x)),
            WICBitmapDecoder(x),
        )

    def create_decoder_from_stream(
        self,
        stream: ComStream,
        options: WICDecodeOption | int,
        vendor_guid: GUID | None = None,
    ) -> WICBitmapDecoder:
        return self.create_decoder_from_stream_nothrow(stream, options, vendor_guid).value

    def create_decoder_from_filehandle_nothrow(
        self,
        handle: int,
        options: WICDecodeOption | int,
        vendor_guid: GUID | None = None,
    ) -> ComResult[WICBitmapDecoder]:
        x = POINTER(IWICBitmapDecoder)()
        return cr(
            self.__o.CreateDecoderFromFileHandle(handle, vendor_guid, int(options), byref(x)), WICBitmapDecoder(x)
        )

    def create_decoder_from_filehandle(
        self,
        handle: int,
        options: WICDecodeOption | int,
        vendor_guid: GUID | None = None,
    ) -> WICBitmapDecoder:
        return self.create_decoder_from_filehandle_nothrow(handle, options, vendor_guid).value

    def create_componentinfo_nothrow(self, component_guid: GUID) -> ComResult[WICComponentInfo]:
        x = POINTER(IWICComponentInfo)()
        return cr(self.__o.CreateComponentInfo(component_guid, byref(x)), WICComponentInfo(x))

    def create_componentinfo(self, component_guid: GUID) -> WICComponentInfo:
        return self.create_componentinfo_nothrow(component_guid).value

    def create_decoder_nothrow(self, container_format: GUID, vendor: GUID | None = None) -> ComResult[WICBitmapDecoder]:
        x = POINTER(IWICBitmapDecoder)()
        return cr(self.__o.CreateDecoder(container_format, vendor, byref(x)), WICBitmapDecoder(x))

    def create_decoder(self, container_format: GUID, vendor: GUID | None = None) -> WICBitmapDecoder:
        return self.create_decoder_nothrow(container_format, vendor).value

    def create_encoder_nothrow(self, container_format: GUID, vendor: GUID | None = None) -> ComResult[WICBitmapEncoder]:
        x = POINTER(IWICBitmapEncoder)()
        return cr(self.__o.CreateEncoder(container_format, vendor, byref(x)), WICBitmapEncoder(x))

    def create_encoder(self, container_format: GUID, vendor: GUID | None = None) -> WICBitmapEncoder:
        return self.create_encoder_nothrow(container_format, vendor).value

    def create_palette_nothrow(self) -> ComResult[WICPalette]:
        x = POINTER(IWICPalette)()
        return cr(self.__o.CreatePalette(byref(x)), WICPalette(x))

    def create_palette(self) -> WICPalette:
        return self.create_palette_nothrow().value

    # TODO
    # STDMETHOD(c_int32, "CreateFormatConverter", (POINTER(POINTER(IWICFormatConverter)),)),
    # STDMETHOD(c_int32, "CreateBitmapScaler", (POINTER(POINTER(IWICBitmapScaler)),)),
    # STDMETHOD(c_int32, "CreateBitmapClipper", (POINTER(POINTER(IWICBitmapClipper)),)),
    # STDMETHOD(
    #     c_int32,
    #     "CreateBitmapFlipRotator",
    #     (POINTER(IWICBitmapFileRotator)),
    # ),
    # STDMETHOD(c_int32, "CreateStream", (POINTER(POINTER(IWICStream)),)),
    # STDMETHOD(c_int32, "CreateColorContext", (POINTER(POINTER(IWICColorContext)))),
    # STDMETHOD(c_int32, "CreateColorTransformer", (POINTER(POINTER(IWICColorTransform)))),
    # STDMETHOD(
    #     c_int32,
    #     "CreateBitmap",
    #     (c_uint32, c_uint32, POINTER(WICPixelFormatGUID), c_int32, POINTER(POINTER(IWICBitmap))),
    # ),
    # STDMETHOD(c_int32, "CreateBitmapFromSource", (POINTER(IWICBitmapSource), c_int32, POINTER(POINTER(IWICBitmap)))),
    # STDMETHOD(
    #     c_int32,
    #     "CreateBitmapFromSourceRect",
    #     (POINTER(IWICBitmapSource), c_uint32, c_uint32, c_uint32, c_uint32, POINTER(POINTER(IWICBitmap))),
    # ),
    # STDMETHOD(
    #     c_int32,
    #     "CreateBitmapFromMemory",
    #     (
    #         c_uint32,
    #         c_uint32,
    #         POINTER(WICPixelFormatGUID),
    #         c_uint32,
    #         c_uint32,
    #         POINTER(c_byte),
    #         POINTER(POINTER(IWICBitmap)),
    #     ),
    # ),

    def create_bitmap_from_bitmaphandle_nothrow(
        self, bmp_handle: int, options: WICBitmapAlphaChannelOption | int, palette_handle: int = 0
    ) -> ComResult[WICBitmap]:
        x = POINTER(IWICBitmap)()
        return cr(self.__o.CreateBitmapFromHBITMAP(bmp_handle, palette_handle, int(options), byref(x)), WICBitmap(x))

    def create_bitmap_from_bitmaphandle(
        self, bmp_handle: int, options: WICBitmapAlphaChannelOption | int, palette_handle: int = 0
    ) -> WICBitmap:
        return self.create_bitmap_from_bitmaphandle_nothrow(bmp_handle, options, palette_handle).value

    def create_bitmap_from_iconhandle_nothrow(self, icon_handle: int) -> ComResult[WICBitmap]:
        x = POINTER(IWICBitmap)()
        return cr(self.__o.CreateBitmapFromHICON(icon_handle, byref(x)), WICBitmap(x))

    def create_bitmap_from_iconhandle(self, icon_handle: int) -> WICBitmap:
        return self.create_bitmap_from_iconhandle_nothrow(icon_handle).value

    def create_componentenumerator_nothrow(
        self, comptype: WICComponentType, options: WICComponentEnumOption = WICComponentEnumOption.DEFAULT
    ) -> ComResult[IUnknownEnumerator]:
        x = POINTER(IEnumUnknown)()
        return cr(self.__o.CreateComponentEnumerator(int(comptype), int(options), byref(x)), IUnknownEnumerator(x))

    def __create_componentenumerator(
        self, comptype: WICComponentType, options: WICComponentEnumOption = WICComponentEnumOption.DEFAULT
    ) -> IUnknownEnumerator:
        return self.create_componentenumerator_nothrow(comptype, options).value

    def iter_infos(
        self, comptype: WICComponentType, options: WICComponentEnumOption = WICComponentEnumOption.DEFAULT
    ) -> Iterator[WICComponentInfo]:
        for p in self.__create_componentenumerator(comptype, options):
            yield WICComponentInfo(p)

    def iter_infos_all(
        self, options: WICComponentEnumOption = WICComponentEnumOption.DEFAULT
    ) -> Iterator[WICComponentInfo]:
        for p in self.__create_componentenumerator(WICComponentType.ALL_COMPONENTS, options):
            yield WICComponentInfo(p)

    def iter_infos_encoder(
        self, options: WICComponentEnumOption = WICComponentEnumOption.DEFAULT
    ) -> Iterator[WICBitmapEncoderInfo]:
        for p in self.__create_componentenumerator(WICComponentType.ENCODER, options):
            yield WICBitmapEncoderInfo(p)

    def create_fast_metadata_encoder_from_decoder_nothrow(
        self, decoder: WICBitmapDecoder
    ) -> ComResult[WICFastMetadataEncoder]:
        x = POINTER(IWICFastMetadataEncoder)()
        return cr(
            self.__o.CreateFastMetadataEncoderFromDecoder(decoder.wrapped_obj, byref(x)), WICFastMetadataEncoder(x)
        )

    def create_fast_metadata_encoder_from_decoder(self, decoder: WICBitmapDecoder) -> WICFastMetadataEncoder:
        return self.create_fast_metadata_encoder_from_decoder_nothrow(decoder).value

    # TODO
    # STDMETHOD(
    #     c_int32,
    #     "CreateFastMetadataEncoderFromFrameDecode",
    #     (POINTER(IWICBitmapFrameDecode), POINTER(POINTER(IWICFastMetadataEncoder))),
    # ),

    def create_querywriter_nothrow(
        self, metadata_type: GUID, vendor: GUID | None = None
    ) -> ComResult[WICFastMetadataEncoder]:
        x = POINTER(IWICFastMetadataEncoder)()
        return cr(self.__o.CreateQueryWriter(metadata_type, vendor, byref(x)), WICFastMetadataEncoder(x))

    # STDMETHOD(
    #     c_int32,
    #     "CreateQueryWriterFromReader",
    #     (POINTER(IWICMetadataQueryReader), POINTER(GUID), POINTER(POINTER(IWICMetadataQueryWriter))),
    # ),


class WICImagingFactory2(WICImagingFactory):
    """WICイメージングのファクトリクラス。IWICImagingFactory2インターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWICImagingFactory2)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IWICImagingFactory2)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def create_imageencoder_nothrow(self, device: ID2D1Device) -> ComResult[WICImageEncoder]:
        x = POINTER(IWICImageEncoder)()
        return cr(self.__o.CreateImageEncoder(device, byref(x)), WICImageEncoder(x))

    def create_imageencoder(self, device: ID2D1Device) -> WICImageEncoder:
        return self.create_imageencoder_nothrow(device).value
