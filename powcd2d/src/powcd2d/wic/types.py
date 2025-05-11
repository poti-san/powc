"""WIC 構造体とCOMインターフェイス。"""

# 単一ファイルにまとめるには多すぎるので分離します。

from ctypes import (
    POINTER,
    WINFUNCTYPE,
    Structure,
    c_byte,
    c_double,
    c_float,
    c_int32,
    c_uint32,
    c_uint64,
    c_void_p,
)
from typing import TYPE_CHECKING

from comtypes import GUID, IUnknown

WICPixelFormatGUID = GUID
WICColor = c_uint32


class WICRect(Structure):
    __slots__ = ()
    _fields_ = (
        ("x", c_int32),
        ("y", c_int32),
        ("width", c_int32),
        ("height", c_int32),
    )
    if TYPE_CHECKING:
        x: int
        y: int
        width: int
        height: int


class WICBitmapPattern(Structure):
    __slots__ = ()
    _fields_ = (
        ("position", c_uint64),
        ("length", c_uint32),
        ("pattern", POINTER(c_byte)),
        ("mask", POINTER(c_byte)),
        ("end_of_stream", c_int32),
    )


class D2D1PixelFormat(Structure):
    """D2D1_PIXEL_FORMAT"""

    __slots__ = ()
    _fields_ = (
        ("format", c_int32),  # DXGI_FORMAT
        ("alpha_mode", c_int32),  # D2D1_ALPHA_MODE
    )


class WICImageParameters(Structure):
    __slots__ = ()
    _fields_ = (
        ("pixel_format", D2D1PixelFormat),
        ("dpi_x", c_float),
        ("dpi_y", c_float),
        ("top", c_float),
        ("left", c_float),
        ("pixel_width", c_uint32),
        ("pixel_height", c_uint32),
    )


class WICBitmapPlaneDesc(Structure):
    __slots__ = ()
    _fields_ = (("format", WICPixelFormatGUID), ("width", c_uint32), ("height", c_uint32))


class WICBitmapPlane(Structure):
    __slots__ = ()
    _fields_ = (
        ("format", WICPixelFormatGUID),
        ("buffer_ptr", POINTER(c_byte)),
        ("stride", c_uint32),
        ("buffer_size", c_uint32),
    )


class WICJpegFrameHeader(Structure):
    __slots__ = ()
    _fields_ = (
        ("width", c_uint32),
        ("hright", c_uint32),
        ("transfer_matrix", c_int32),  # WICJpegTransferMatrix
        ("scan_type", c_int32),  # WICJpegScanType
        ("c_components", c_uint32),
        ("component_ids", c_uint32),
        ("sample_factors", c_uint32),
        ("quantization_table_indices", c_uint32),
    )


class WICJpegScanHeader(Structure):
    __slots__ = ()
    _fields_ = (
        ("c_components", c_uint32),
        ("restart_interval", c_uint32),
        ("component_selectors", c_uint32),
        ("huffman_table_indices", c_uint32),
        ("start_spectral_selection", c_byte),
        ("end_spectral_selection", c_byte),
        ("successive_approximation_high", c_byte),
        ("successive_approximation_low", c_byte),
    )


# TODO
IPropertyBag2 = IUnknown
IWICFastMetadataDecoder = IUnknown
ID2D1Image = IUnknown
ID2D1Device = IUnknown


class IWICPalette(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000040-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICBitmapSource(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000120-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICFormatConverter(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{00000301-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICPlanarFormatConverter(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{BEBEE9CB-83B0-4DCC-8132-B0AAA55EAC96}")


class IWICBitmapScaler(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{00000302-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICBitmapClipper(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{E4FBCF03-223D-4e81-9333-D635556DD1B5}")


class IWICBitmapFlipRotator(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{5009834F-2D6A-41ce-9E1B-17C5AFF7A782}")


class IWICBitmapLock(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000123-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICBitmap(IWICBitmapSource):
    __slots__ = ()
    _iid_ = GUID("{00000121-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICColorContext(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{3C613A02-34B2-44ea-9A7C-45AEA9C6FD6D}")


class IWICColorTransform(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{B66F034F-D0E2-40ab-B436-6DE39E321A94}")


class IWICFastMetadataEncoder(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{B84E2C09-78C9-4AC4-8BD3-524AE1663A2F}")


class IWICStream(IUnknown):  # IStream
    __slots__ = ()
    _iid_ = GUID("{135FF860-22B7-4ddf-B0F6-218F4F299A43}")


class IWICEnumMetadataItem(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{DC2BB46D-3F07-481E-8625-220C4AEDBB33}")


class IWICMetadataQueryReader(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{30989668-E1C9-4597-B395-458EEDB808DF}")


class IWICMetadataQueryWriter(IUnknown):  # IWICMetadataQueryReader
    __slots__ = ()
    _iid_ = GUID("{A721791A-0DEF-4d06-BD91-2118BF1DB10B}")


class IWICBitmapEncoder(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000103-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICBitmapFrameEncode(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{00000105-a8f2-4877-ba0a-fd2b6645fb94}")


class IWICPlanarBitmapFrameEncode(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{F928B7B8-2221-40C1-B72E-7E82F1974D1A}")


class IWICImageEncoder(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{04C75BF8-3CE1-473B-ACC5-3CC4F5E94999}")


class IWICBitmapDecoder(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{9EDDE9E7-8DEE-47ea-99DF-E6FAF2ED44BF}")


class IWICBitmapSourceTransform(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{3B16811B-6A43-4ec9-B713-3D5A0C13B940}")


class IWICPlanarBitmapSourceTransform(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{3AFF9CCE-BE95-4303-B927-E7D16FF4A613}")


class IWICBitmapFrameDecode(IUnknown):  # IWICBitmapSource
    __slots__ = ()
    _iid_ = GUID("{3B16811B-6A43-4ec9-A813-3D930C13B940}")


class IWICProgressiveLevelControl(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{DAAC296F-7AA5-4dbf-8D15-225C5976F891}")


class IWICDisplayAdaptationControl(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{de9d91d2-70b4-4f41-836c-25fcd39626d3}")


class IWICProgressCallback(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{4776F9CD-9517-45FA-BF24-E89C5EC5C60C}")


PFNProgressNotification = WINFUNCTYPE(c_int32, c_void_p, c_uint32, c_int32, c_double)


class IWICBitmapCodecProgressNotification(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{64C1024E-C3CF-4462-8078-88C2B11C46D9}")


class IWICImagingFactory(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{ec5ec8a9-c395-4314-9c77-54d7a935ff70}")


class IWICImagingFactory2(IUnknown):  # IWICImagingFactory
    __slots__ = ()
    _iid_ = GUID("{7B816B45-1996-4476-B132-DE9E247C8AF0}")


class IWICComponentInfo(IUnknown):
    """IWICComponentInfoインターフェイス"""

    __slots__ = ()
    _iid_ = GUID("{23BC3F0A-698B-4357-886B-F24D50671334}")


class IWICFormatConverterInfo(IUnknown):  # IWICComponentInfo
    """IWICFormatConverterInfoインターフェイス"""

    __slots__ = ()
    _iid_ = GUID("{9F34FB65-13F4-4f15-BC57-3726B5E53D9F}")


class IWICBitmapCodecInfo(IUnknown):  # IWICComponentInfo
    __slots__ = ()
    _iid_ = GUID("{E87A44C4-B76E-4c47-8B09-298EB12A2714}")


class IWICBitmapEncoderInfo(IUnknown):  # IWICBitmapCodecInfo
    __slots__ = ()
    _iid_ = GUID("{94C9B4EE-A09F-4f92-8A1E-4A9BCE7E76FB}")


class IWICBitmapDecoderInfo(IUnknown):  # IWICBitmapCodecInfo
    __slots__ = ()
    _iid_ = GUID("{D8CD007F-D08F-4191-9BFC-236EA7F0E4B5}")


class IWICPixelFormatInfo(IUnknown):  # IWICComponentInfo
    __slots__ = ()
    _iid_ = GUID("{E8EDA601-3D48-431a-AB44-69059BE88BBE}")


class IWICPixelFormatInfo2(IUnknown):  # IWICPixelFormatInfo
    __slots__ = ()
    _iid_ = GUID("{A9DB33A2-AF5F-43C7-B679-74F5984B5AA4}")
