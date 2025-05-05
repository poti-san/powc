from ctypes import c_wchar_p

from comtypes import STDMETHOD
from powc.comobj import IEnumString, IEnumUnknown
from powc.stream import IStream
from powcpropsys.propvariant import PropVariant

from .types import *

IWICPalette._methods_ = [
    STDMETHOD(c_int32, "InitializePredefined", (c_int32, c_int32)),
    STDMETHOD(c_int32, "InitializeCustom", (POINTER(WICColor), c_uint32)),
    STDMETHOD(c_int32, "InitializeFromBitmap", (POINTER(IWICBitmapSource), c_uint32, c_int32)),
    STDMETHOD(c_int32, "InitializeFromPalette", (POINTER(IWICPalette),)),
    STDMETHOD(c_int32, "GetType", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "GetColorCount", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetColors", (c_uint32, POINTER(WICColor), POINTER(c_uint32))),
    STDMETHOD(c_int32, "IsBlackWhite", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "IsGrayscale", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "HasAlpha", (POINTER(c_int32),)),
]

IWICBitmapSource._methods_ = [
    STDMETHOD(c_int32, "GetSize", (POINTER(c_uint32), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetPixelFormat", (POINTER(WICPixelFormatGUID),)),
    STDMETHOD(c_int32, "GetResolution", (POINTER(c_double), POINTER(c_double))),
    STDMETHOD(c_int32, "CopyPalette", (POINTER(IWICPalette),)),
    STDMETHOD(c_int32, "CopyPixels", (POINTER(WICRect), c_uint32, c_uint32, POINTER(c_byte))),
]
IWICFormatConverter._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(
        c_int32,
        "Initialize",
        (POINTER(IWICBitmapSource), POINTER(WICPixelFormatGUID), c_int32, POINTER(IWICPalette), c_double, c_int32),
    ),
    STDMETHOD(c_int32, "CanConvert", (POINTER(WICPixelFormatGUID), POINTER(WICPixelFormatGUID), POINTER(c_int32))),
]
IWICPlanarFormatConverter._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(
        c_int32,
        "Initialize",
        (
            POINTER(POINTER(IWICBitmapSource)),
            c_uint32,
            POINTER(WICPixelFormatGUID),
            c_int32,
            POINTER(IWICPalette),
            c_double,
            c_int32,
        ),
    ),
    STDMETHOD(c_int32, "CanConvert", (POINTER(WICPixelFormatGUID), c_uint32, POINTER(GUID), POINTER(c_int32))),
]
IWICBitmapScaler._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(c_int32, "Initialize", (POINTER(IWICBitmapSource), c_uint32, c_uint32, c_int32)),
]
IWICBitmapClipper._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(c_int32, "Initialize", (POINTER(IWICBitmapSource), POINTER(WICRect))),
]
IWICBitmapFlipRotator._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(c_int32, "Initialize", (POINTER(IWICBitmapSource), c_int32)),
]
IWICBitmapLock._methods_ = [
    STDMETHOD(c_int32, "GetSize", (POINTER(c_uint32), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetStride", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetDataPointer", (POINTER(c_uint32), POINTER(POINTER(c_byte)))),
    STDMETHOD(c_int32, "GetPixelFormat", (POINTER(WICPixelFormatGUID),)),
]
IWICBitmap._methods_ = [
    STDMETHOD(c_int32, "Lock", (POINTER(WICRect), c_uint32, POINTER(IWICBitmapLock))),
    STDMETHOD(c_int32, "SetPalette", (POINTER(IWICPalette),)),
    STDMETHOD(c_int32, "SetResolution", (c_double, c_double)),
]
IWICColorContext._methods_ = [
    STDMETHOD(c_int32, "InitializeFromFilename", (c_wchar_p,)),
    STDMETHOD(c_int32, "InitializeFromMemory", (POINTER(c_byte), c_uint32)),
    STDMETHOD(c_int32, "InitializeFromExifColorSpace", (c_uint32,)),
    STDMETHOD(c_int32, "GetType", (c_int32,)),  # WICColorContextType
    STDMETHOD(c_int32, "GetProfileBytes", (c_uint32, POINTER(c_byte), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetExifColorSpace", (POINTER(c_uint32),)),
]
IWICColorTransform._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(
        c_int32,
        "Initialize",
        (
            POINTER(IWICBitmapSource),
            POINTER(IWICColorContext),
            POINTER(IWICColorContext),
            POINTER(WICPixelFormatGUID),
        ),
    ),
]
IWICFastMetadataEncoder._methods_ = [
    STDMETHOD(c_int32, "Commit", ()),
    STDMETHOD(c_int32, "GetMetadataQueryWriter", (POINTER(IWICMetadataQueryWriter),)),
]

IWICStream._methods_ = [
    *IStream._methods_,
    STDMETHOD(c_int32, "InitializeFromIStream", (POINTER(IStream),)),
    STDMETHOD(c_int32, "InitializeFromFilename", (c_wchar_p, c_uint32)),
    STDMETHOD(c_int32, "InitializeFromMemory", (c_void_p, c_uint32)),
    STDMETHOD(c_int32, "InitializeFromIStreamRegion", (POINTER(IStream), c_uint64, c_uint64)),
]


IWICEnumMetadataItem._methods_ = [
    STDMETHOD(c_int32, "Next", (POINTER(PropVariant), POINTER(PropVariant), POINTER(PropVariant), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IWICEnumMetadataItem)),)),
]
IWICMetadataQueryReader._methods_ = [
    STDMETHOD(c_int32, "GetContainerFormat", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetLocation", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetMetadataByName", (c_wchar_p, POINTER(PropVariant))),
    STDMETHOD(c_int32, "GetEnumerator", (POINTER(IEnumString),)),
]
IWICMetadataQueryWriter._methods_ = [
    *IWICMetadataQueryReader._methods_,
    STDMETHOD(c_int32, "SetMetadataByName", (c_wchar_p, POINTER(PropVariant))),
    STDMETHOD(c_int32, "RemoveMetadataByName", (c_wchar_p,)),
]


IWICBitmapEncoder._methods_ = [
    STDMETHOD(c_int32, "Initialize", (POINTER(IStream), c_int32)),
    STDMETHOD(c_int32, "GetContainerFormat", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetEncoderInfo", (POINTER(IWICBitmapEncoderInfo),)),
    STDMETHOD(c_int32, "SetColorContexts", (c_uint32, POINTER(IWICColorContext))),
    STDMETHOD(c_int32, "SetPalette", (POINTER(IWICPalette),)),
    STDMETHOD(c_int32, "SetThumbnail", (POINTER(IWICBitmapSource),)),
    STDMETHOD(c_int32, "SetPreview", (POINTER(IWICBitmapSource),)),
    STDMETHOD(c_int32, "CreateNewFrame", (POINTER(POINTER(IWICBitmapFrameEncode)), POINTER(POINTER(IPropertyBag2)))),
    STDMETHOD(c_int32, "Commit", ()),
    STDMETHOD(c_int32, "GetMetadataQueryWriter", (POINTER(POINTER(IWICMetadataQueryWriter)),)),
]
IWICBitmapFrameEncode._methods_ = [
    STDMETHOD(c_int32, "Initialize", (POINTER(IPropertyBag2),)),
    STDMETHOD(c_int32, "SetSize", (c_uint32, c_uint32)),
    STDMETHOD(c_int32, "SetResolution", (c_double, c_double)),
    STDMETHOD(c_int32, "SetPixelFormat", (POINTER(WICPixelFormatGUID),)),
    STDMETHOD(c_int32, "SetColorContexts", (c_uint32, POINTER(IWICColorContext))),
    STDMETHOD(c_int32, "SetPalette", (POINTER(IWICPalette),)),
    STDMETHOD(c_int32, "SetThumbnail", (POINTER(IWICBitmapSource),)),
    STDMETHOD(c_int32, "WritePixels", (c_uint32, c_uint32, c_uint32, POINTER(c_byte))),
    STDMETHOD(c_int32, "WriteSource", (POINTER(IWICBitmapSource), POINTER(WICRect))),
    STDMETHOD(c_int32, "Commit", ()),
    STDMETHOD(c_int32, "GetMetadataQueryWriter", (POINTER(IWICMetadataQueryWriter),)),
]
IWICPlanarBitmapFrameEncode._methods_ = [
    STDMETHOD(c_int32, "WritePixels", (c_uint32, POINTER(WICBitmapPlane), c_uint32)),
    STDMETHOD(c_int32, "WriteSource", (POINTER(POINTER(IWICBitmapSource)), c_uint32, POINTER(WICRect))),
]
IWICImageEncoder._methods_ = [
    STDMETHOD(
        c_int32, "WriteFrame", (POINTER(ID2D1Image), POINTER(IWICBitmapFrameEncode), POINTER(WICImageParameters))
    ),
    STDMETHOD(
        c_int32,
        "WriteFrameThumbnail",
        (POINTER(ID2D1Image), POINTER(IWICBitmapFrameEncode), POINTER(WICImageParameters)),
    ),
    STDMETHOD(
        c_int32,
        "WriteThumbnail",
        (POINTER(ID2D1Image), POINTER(IWICBitmapEncoder), POINTER(WICImageParameters)),
    ),
]

IWICBitmapDecoder._methods_ = [
    STDMETHOD(c_int32, "QueryCapability", (POINTER(IStream), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Initialize", (POINTER(IStream), c_int32)),
    STDMETHOD(c_int32, "GetContainerFormat", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetDecoderInfo", (POINTER(POINTER(IWICBitmapDecoderInfo)),)),
    STDMETHOD(c_int32, "CopyPalette", (POINTER(IWICPalette),)),
    STDMETHOD(c_int32, "GetMetadataQueryReader", (POINTER(POINTER(IWICMetadataQueryReader)),)),
    STDMETHOD(c_int32, "GetPreview", (POINTER(POINTER(IWICBitmapSource)),)),
    STDMETHOD(c_int32, "GetColorContexts", (c_uint32, POINTER(POINTER(IWICColorContext)), POINTER(c_uint32))),
    STDMETHOD(
        c_int32,
        "GetThumbnail",
        (POINTER(POINTER(IWICBitmapSource)),),
    ),
    STDMETHOD(c_int32, "GetFrameCount", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetFrame", (c_uint32, POINTER(POINTER(IWICBitmapFrameDecode)))),
]

IWICBitmapSourceTransform._methods_ = [
    STDMETHOD(
        c_int32,
        "CopyPixels",
        (POINTER(WICRect), c_uint32, c_uint32, POINTER(WICPixelFormatGUID), c_int32, c_uint32, c_uint32, c_void_p),
    ),
    STDMETHOD(c_int32, "GetClosestSize", (POINTER(c_uint32), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetClosestPixelFormat", (POINTER(WICPixelFormatGUID),)),
    STDMETHOD(c_int32, "DoesSupportTransform", (c_int32, POINTER(c_int32))),
    STDMETHOD(c_int32, "", ()),
]
IWICPlanarBitmapSourceTransform._methods_ = [
    STDMETHOD(
        c_int32,
        "DoesSupportTransform",
        (
            POINTER(c_uint32),
            POINTER(c_uint32),
            c_int32,
            c_int32,
            POINTER(WICPixelFormatGUID),
            POINTER(WICBitmapPlaneDesc),
            c_uint32,
            POINTER(c_int32),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CopyPixels",
        (POINTER(WICRect), c_uint32, c_uint32, c_int32, c_int32, POINTER(WICBitmapPlane), c_uint32),
    ),
]
IWICBitmapFrameDecode._methods_ = [
    *IWICBitmapSource._methods_,
    STDMETHOD(c_int32, "GetMetadataQueryReader", (POINTER(POINTER(IWICMetadataQueryReader)),)),
    STDMETHOD(c_int32, "GetColorContexts", (c_uint32, POINTER(POINTER(IWICColorContext)), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetThumbnail", (POINTER(POINTER(IWICBitmapSource)),)),
]
IWICProgressiveLevelControl._methods_ = [
    STDMETHOD(c_int32, "GetLevelCount", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetCurrentLevel", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "SetCurrentLevel", (c_uint32,)),
]
IWICDisplayAdaptationControl._methods_ = [
    STDMETHOD(c_int32, "DoesSupportChangingMaxLuminance", (POINTER(WICPixelFormatGUID), POINTER(c_int32))),
    STDMETHOD(c_int32, "SetDisplayMaxLuminance", (c_float,)),
    STDMETHOD(c_int32, "GetDisplayMaxLuminance", (POINTER(c_float),)),
]
IWICProgressCallback._methods_ = [
    STDMETHOD(c_int32, "Notify", (c_uint32, c_int32, c_double)),
]
IWICBitmapCodecProgressNotification._methods_ = [
    STDMETHOD(c_int32, "RegisterProgressNotification", (PFNProgressNotification, c_void_p, c_uint32)),
]

IWICImagingFactory._methods_ = [
    STDMETHOD(
        c_int32,
        "CreateDecoderFromFilename",
        (c_wchar_p, POINTER(GUID), c_uint32, c_int32, POINTER(POINTER(IWICBitmapDecoder))),
    ),
    STDMETHOD(
        c_int32,
        "CreateDecoderFromStream",
        (POINTER(IStream), POINTER(GUID), c_int32, POINTER(POINTER(IWICBitmapDecoder))),
    ),
    STDMETHOD(
        c_int32,
        "CreateDecoderFromFileHandle",
        (c_void_p, POINTER(GUID), c_int32, POINTER(POINTER(IWICBitmapDecoder))),
    ),
    STDMETHOD(c_int32, "CreateComponentInfo", (POINTER(GUID), POINTER(POINTER(IWICComponentInfo)))),
    STDMETHOD(c_int32, "CreateDecoder", (POINTER(GUID), POINTER(GUID), POINTER(POINTER(IWICBitmapDecoder)))),
    STDMETHOD(c_int32, "CreateEncoder", (POINTER(GUID), POINTER(GUID), POINTER(POINTER(IWICBitmapEncoder)))),
    STDMETHOD(c_int32, "CreatePalette", (POINTER(POINTER(IWICPalette)),)),
    STDMETHOD(c_int32, "CreateFormatConverter", (POINTER(POINTER(IWICFormatConverter)),)),
    STDMETHOD(c_int32, "CreateBitmapScaler", (POINTER(POINTER(IWICBitmapScaler)),)),
    STDMETHOD(c_int32, "CreateBitmapClipper", (POINTER(POINTER(IWICBitmapClipper)),)),
    STDMETHOD(
        c_int32,
        "CreateBitmapFlipRotator",
        (POINTER(POINTER(IWICBitmapFlipRotator)),),
    ),
    STDMETHOD(c_int32, "CreateStream", (POINTER(POINTER(IWICStream)),)),
    STDMETHOD(c_int32, "CreateColorContext", (POINTER(POINTER(IWICColorContext)),)),
    STDMETHOD(c_int32, "CreateColorTransformer", (POINTER(POINTER(IWICColorTransform)),)),
    STDMETHOD(
        c_int32,
        "CreateBitmap",
        (c_uint32, c_uint32, POINTER(WICPixelFormatGUID), c_int32, POINTER(POINTER(IWICBitmap))),
    ),
    STDMETHOD(c_int32, "CreateBitmapFromSource", (POINTER(IWICBitmapSource), c_int32, POINTER(POINTER(IWICBitmap)))),
    STDMETHOD(
        c_int32,
        "CreateBitmapFromSourceRect",
        (POINTER(IWICBitmapSource), c_uint32, c_uint32, c_uint32, c_uint32, POINTER(POINTER(IWICBitmap))),
    ),
    STDMETHOD(
        c_int32,
        "CreateBitmapFromMemory",
        (
            c_uint32,
            c_uint32,
            POINTER(WICPixelFormatGUID),
            c_uint32,
            c_uint32,
            POINTER(c_byte),
            POINTER(POINTER(IWICBitmap)),
        ),
    ),
    STDMETHOD(c_int32, "CreateBitmapFromHBITMAP", (c_void_p, c_void_p, c_int32, POINTER(POINTER(IWICBitmap)))),
    STDMETHOD(c_int32, "CreateBitmapFromHICON", (c_void_p, POINTER(POINTER(IWICBitmap)))),
    STDMETHOD(c_int32, "CreateComponentEnumerator", (c_uint32, c_uint32, POINTER(POINTER(IEnumUnknown)))),
    STDMETHOD(
        c_int32,
        "CreateFastMetadataEncoderFromDecoder",
        (POINTER(IWICBitmapDecoder), POINTER(POINTER(IWICFastMetadataEncoder))),
    ),
    STDMETHOD(
        c_int32,
        "CreateFastMetadataEncoderFromFrameDecode",
        (POINTER(IWICBitmapFrameDecode), POINTER(POINTER(IWICFastMetadataEncoder))),
    ),
    STDMETHOD(c_int32, "CreateQueryWriter", (POINTER(GUID), POINTER(GUID), POINTER(POINTER(IWICFastMetadataEncoder)))),
    STDMETHOD(
        c_int32,
        "CreateQueryWriterFromReader",
        (POINTER(IWICMetadataQueryReader), POINTER(GUID), POINTER(POINTER(IWICMetadataQueryWriter))),
    ),
]

IWICImagingFactory2._methods_ = [
    *IWICImagingFactory._methods_,
    STDMETHOD(c_int32, "CreateImageEncoder", (POINTER(ID2D1Device), POINTER(POINTER(IWICImageEncoder)))),
]
IWICComponentInfo._methods_ = [
    STDMETHOD(c_int32, "GetComponentType", (POINTER(c_int32),)),  # WICComponentType
    STDMETHOD(c_int32, "GetCLSID", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetSigningStatus", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "GetAuthor", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetVendorGUID", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetVersion", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetSpecVersion", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetFriendlyName", (c_uint32, c_wchar_p, POINTER(c_uint32))),
]
IWICFormatConverterInfo._methods_ = [
    *IWICComponentInfo._methods_,
    STDMETHOD(c_int32, "GetPixelFormats", (c_uint32, POINTER(WICPixelFormatGUID), POINTER(c_uint32))),
    STDMETHOD(c_int32, "CreateInstance", (POINTER(POINTER(IWICFormatConverter)),)),
]
IWICBitmapCodecInfo._methods_ = [
    *IWICComponentInfo._methods_,
    STDMETHOD(c_int32, "GetContainerFormat", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetPixelFormats", (c_uint32, POINTER(GUID), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetColorManagementVersion", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetDeviceManufacturer", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetDeviceModels", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetMimeTypes", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetFileExtensions", (c_uint32, c_wchar_p, POINTER(c_uint32))),
    STDMETHOD(c_int32, "DoesSupportAnimation", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "DoesSupportChromakey", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "DoesSupportLossless", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "DoesSupportMultiframe", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "MatchesMimeType", (c_wchar_p, POINTER(c_int32))),
]
IWICBitmapEncoderInfo._methods_ = [
    *IWICBitmapCodecInfo._methods_,
    STDMETHOD(c_int32, "CreateInstance", (POINTER(POINTER(IWICBitmapEncoder)),)),
]
IWICBitmapDecoderInfo._methods_ = [
    *IWICBitmapCodecInfo._methods_,
    STDMETHOD(c_int32, "GetPatterns", (c_uint32, POINTER(WICBitmapPattern), POINTER(c_uint32), POINTER(c_uint32))),
    STDMETHOD(c_int32, "MatchesPattern", (POINTER(IStream), POINTER(c_int32))),
    STDMETHOD(c_int32, "CreateInstance", (POINTER(POINTER(IWICBitmapDecoder)),)),
]

IWICPixelFormatInfo._methods_ = [
    *IWICComponentInfo._methods_,
    STDMETHOD(c_int32, "GetFormatGUID", (POINTER(GUID),)),
    STDMETHOD(c_int32, "GetColorContext", (POINTER(POINTER(IWICColorContext)),)),
    STDMETHOD(c_int32, "GetBitsPerPixel", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetChannelCount", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetChannelMask", (c_uint32, c_uint32, POINTER(c_byte), POINTER(c_uint32))),
]
IWICPixelFormatInfo2._methods_ = [
    *IWICPixelFormatInfo._methods_,
    STDMETHOD(c_int32, "SupportsTransparency", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "GetNumericRepresentation", (POINTER(c_int32),)),
]
