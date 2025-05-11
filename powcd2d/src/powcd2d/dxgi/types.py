"""Direct2Dに必要な一部のDIGI型を定義します。"""

# IUnknown派生クラスの_methods_はinterfacebody.pyに記述します。

from ctypes import POINTER, Structure, c_size_t, c_uint32, c_uint64, c_void_p, c_wchar
from ctypes.wintypes import RECT
from enum import IntEnum, IntFlag
from typing import TYPE_CHECKING

from comtypes import GUID, IUnknown


class DXGIFormat(IntEnum):
    """DXGI_FORMAT"""

    UNKNOWN = 0
    R32G32B32A32_TYPELESS = 1
    R32G32B32A32_FLOAT = 2
    R32G32B32A32_UINT = 3
    R32G32B32A32_SINT = 4
    R32G32B32_TYPELESS = 5
    R32G32B32_FLOAT = 6
    R32G32B32_UINT = 7
    R32G32B32_SINT = 8
    R16G16B16A16_TYPELESS = 9
    R16G16B16A16_FLOAT = 10
    R16G16B16A16_UNORM = 11
    R16G16B16A16_UINT = 12
    R16G16B16A16_SNORM = 13
    R16G16B16A16_SINT = 14
    R32G32_TYPELESS = 15
    R32G32_FLOAT = 16
    R32G32_UINT = 17
    R32G32_SINT = 18
    R32G8X24_TYPELESS = 19
    D32_FLOAT_S8X24_UINT = 20
    R32_FLOAT_X8X24_TYPELESS = 21
    X32_TYPELESS_G8X24_UINT = 22
    R10G10B10A2_TYPELESS = 23
    R10G10B10A2_UNORM = 24
    R10G10B10A2_UINT = 25
    R11G11B10_FLOAT = 26
    R8G8B8A8_TYPELESS = 27
    R8G8B8A8_UNORM = 28
    R8G8B8A8_UNORM_SRGB = 29
    R8G8B8A8_UINT = 30
    R8G8B8A8_SNORM = 31
    R8G8B8A8_SINT = 32
    R16G16_TYPELESS = 33
    R16G16_FLOAT = 34
    R16G16_UNORM = 35
    R16G16_UINT = 36
    R16G16_SNORM = 37
    R16G16_SINT = 38
    R32_TYPELESS = 39
    D32_FLOAT = 40
    R32_FLOAT = 41
    R32_UINT = 42
    R32_SINT = 43
    R24G8_TYPELESS = 44
    D24_UNORM_S8_UINT = 45
    R24_UNORM_X8_TYPELESS = 46
    X24_TYPELESS_G8_UINT = 47
    R8G8_TYPELESS = 48
    R8G8_UNORM = 49
    R8G8_UINT = 50
    R8G8_SNORM = 51
    R8G8_SINT = 52
    R16_TYPELESS = 53
    R16_FLOAT = 54
    D16_UNORM = 55
    R16_UNORM = 56
    R16_UINT = 57
    R16_SNORM = 58
    R16_SINT = 59
    R8_TYPELESS = 60
    R8_UNORM = 61
    R8_UINT = 62
    R8_SNORM = 63
    R8_SINT = 64
    A8_UNORM = 65
    R1_UNORM = 66
    R9G9B9E5_SHAREDEXP = 67
    R8G8_B8G8_UNORM = 68
    G8R8_G8B8_UNORM = 69
    BC1_TYPELESS = 70
    BC1_UNORM = 71
    BC1_UNORM_SRGB = 72
    BC2_TYPELESS = 73
    BC2_UNORM = 74
    BC2_UNORM_SRGB = 75
    BC3_TYPELESS = 76
    BC3_UNORM = 77
    BC3_UNORM_SRGB = 78
    BC4_TYPELESS = 79
    BC4_UNORM = 80
    BC4_SNORM = 81
    BC5_TYPELESS = 82
    BC5_UNORM = 83
    BC5_SNORM = 84
    B5G6R5_UNORM = 85
    B5G5R5A1_UNORM = 86
    B8G8R8A8_UNORM = 87
    B8G8R8X8_UNORM = 88
    R10G10B10_XR_BIAS_A2_UNORM = 89
    B8G8R8A8_TYPELESS = 90
    B8G8R8A8_UNORM_SRGB = 91
    B8G8R8X8_TYPELESS = 92
    B8G8R8X8_UNORM_SRGB = 93
    BC6H_TYPELESS = 94
    BC6H_UF16 = 95
    BC6H_SF16 = 96
    BC7_TYPELESS = 97
    BC7_UNORM = 98
    BC7_UNORM_SRGB = 99
    AYUV = 100
    Y410 = 101
    Y416 = 102
    NV12 = 103
    P010 = 104
    P016 = 105
    _420_OPAQUE = 106
    YUY2 = 107
    Y210 = 108
    Y216 = 109
    NV11 = 110
    AI44 = 111
    IA44 = 112
    P8 = 113
    A8P8 = 114
    B4G4R4A4_UNORM = 115

    P208 = 130
    V208 = 131
    V408 = 132

    SAMPLER_FEEDBACK_MIN_MIP_OPAQUE = 189
    SAMPLER_FEEDBACK_MIP_REGION_USED_OPAQUE = 190

    A4B4G4R4_UNORM = 191


class DXGIRational(Structure):
    """DXGI_RATIONAL"""

    __slots__ = ()
    _fields_ = (
        ("numerator", c_uint32),
        ("denominator", c_uint32),
    )


class DXGISampleDescQuality(IntEnum):
    """DXGI_SAMPLE_DESC_QUALITY"""

    DXGI_STANDARD_MULTISAMPLE_QUALITY_PATTERN = 0xFFFFFFFF
    DXGI_CENTER_MULTISAMPLE_QUALITY_PATTERN = 0xFFFFFFFE


class DXGISampleDesc(Structure):
    """DXGI_SAMPLE_DESC"""

    __slots__ = ()
    _fields_ = (("count", c_uint32), ("quality", c_uint32))


class DXGIColorSpaceType(IntEnum):
    """DXGI_COLOR_SPACE_TYPE"""

    RGB_FULL_G22_NONE_P709 = 0
    RGB_FULL_G10_NONE_P709 = 1
    RGB_STUDIO_G22_NONE_P709 = 2
    RGB_STUDIO_G22_NONE_P2020 = 3
    RESERVED = 4
    YCBCR_FULL_G22_NONE_P709_X601 = 5
    YCBCR_STUDIO_G22_LEFT_P601 = 6
    YCBCR_FULL_G22_LEFT_P601 = 7
    YCBCR_STUDIO_G22_LEFT_P709 = 8
    YCBCR_FULL_G22_LEFT_P709 = 9
    YCBCR_STUDIO_G22_LEFT_P2020 = 10
    YCBCR_FULL_G22_LEFT_P2020 = 11
    RGB_FULL_G2084_NONE_P2020 = 12
    YCBCR_STUDIO_G2084_LEFT_P2020 = 13
    RGB_STUDIO_G2084_NONE_P2020 = 14
    YCBCR_STUDIO_G22_TOPLEFT_P2020 = 15
    YCBCR_STUDIO_G2084_TOPLEFT_P2020 = 16
    RGB_FULL_G22_NONE_P2020 = 17
    YCBCR_STUDIO_GHLG_TOPLEFT_P2020 = 18
    YCBCR_FULL_GHLG_TOPLEFT_P2020 = 19
    RGB_STUDIO_G24_NONE_P709 = 20
    RGB_STUDIO_G24_NONE_P2020 = 21
    YCBCR_STUDIO_G24_LEFT_P709 = 22
    YCBCR_STUDIO_G24_LEFT_P2020 = 23
    YCBCR_STUDIO_G24_TOPLEFT_P2020 = 24
    CUSTOM = 0xFFFFFFFF


from ctypes import Structure, c_byte, c_float, c_int32, c_uint32
from enum import IntEnum

# #define MAKE_DXGI_HRESULT(code) MAKE_HRESULT(1, _FACDXGI, code)
# #define MAKE_DXGI_STATUS(code)  MAKE_HRESULT(0, _FACDXGI, code)


class DXGICpuAccess(IntEnum):
    """DXGI_CPU_ACCESS"""

    NONE = 0
    DYNAMIC = 1
    READ_WRITE = 2
    SCRATCH = 3
    FIELD = 15


class DXGIRGB(Structure):
    """DXGI_RGB"""

    __slots__ = ()
    _fields_ = (("red", c_float), ("green", c_float), ("blue", c_float))


class D3DColorValue(Structure):
    """D3DCOLORVALUE"""

    __slots__ = ()
    _fields_ = (("r", c_float), ("g", c_float), ("b", c_float), ("a", c_float))


DXGIRGBA = D3DColorValue
"""DXGI_RGBA"""


class DXGIGammaControl(Structure):
    """DXGI_GAMMA_CONTROL"""

    __slots__ = ()
    _fields_ = (
        ("scale", DXGIRGB),
        ("offset", DXGIRGB),
        ("gammacurve", DXGIRGB * 1025),
    )


class DXGIGammaControlCaps(Structure):
    """DXGI_GAMMA_CONTROL_CAPABILITIES"""

    __slots__ = ()
    _fields_ = (
        ("scale_and_offset_supported", c_uint32),
        ("max_converted_value", c_float),
        ("min_converted_value", c_float),
        ("num_gamma_control_points", c_uint32),
        ("control_point_positions", c_float * 1025),
    )


class DXGIModeScanlineOrder(IntEnum):
    """DXGI_MODE_SCANLINE_ORDER"""

    UNSPECIFIED = 0
    PROGRESSIVE = 1
    UPPER_FIELD_FIRST = 2
    LOWER_FIELD_FIRST = 3


class DXGIModeScaling(IntEnum):
    """DXGI_MODE_SCALING"""

    UNSPECIFIED = 0
    CENTERED = 1
    STRETCHED = 2


class DXGIModeRotation(IntEnum):
    """DXGI_MODE_ROTATION"""

    UNSPECIFIED = 0
    IDENTITY = 1
    ROTATE90 = 2
    ROTATE180 = 3
    ROTATE270 = 4


class DXGIModeDesc(Structure):
    """DXGI_MODE_DESC"""

    __slots__ = ()
    _fields_ = (
        ("width", c_uint32),
        ("height", c_uint32),
        ("refresh_rate", c_int32),  # DXGI_RATIONAL
        ("format", c_int32),  # DXGI_FORMAT
        ("scanline_ordering", c_int32),  # DXGI_MODE_SCANLINE_ORDER
        ("scaling", c_int32),  # DXGI_MODE_SCALING
    )


class DXGIJpegDCHuffmanTable(Structure):
    """DXGI_JPEG_DC_HUFFMAN_TABLE"""

    __slots__ = ()
    _fields_ = (
        ("code_counts", c_byte * 12),
        ("code_values", c_byte * 12),
    )


class DXGIJpegACHuffmanTable(Structure):
    """DXGI_JPEG_AC_HUFFMAN_TABLE"""

    __slots__ = ()
    _fields_ = (
        ("code_counts", c_byte * 16),
        ("code_values", c_byte * 162),
    )


class DXGIJpegQuantizationTable(Structure):
    """DXGI_JPEG_QUANTIZATION_TABLE"""

    __slots__ = ()
    _fields_ = (("elements", c_byte * 64),)


D2DColorF = D3DColorValue
"""D2D_COLOR_F"""


class DXGIUsage(IntFlag):
    """DXGI_USAGE"""

    SHADER_INPUT = 0x00000010
    RENDER_TARGET_OUTPUT = 0x00000020
    BACK_BUFFER = 0x00000040
    SHARED = 0x00000080
    READ_ONLY = 0x00000100
    DISCARD_ON_PRESENT = 0x00000200
    UNORDERED_ACCESS = 0x00000400


class DXGIFrameStats(Structure):
    """DXGI_FRAME_STATISTICS"""

    __slots__ = ()
    _fields_ = (
        ("present_count", c_uint32),
        ("present_refresh_count", c_uint32),
        ("sync_refresh_count", c_uint32),
        ("sync_qpc_time", c_uint64),
        ("sync_gpu_time", c_uint64),
    )


class DXGIMappedRect(Structure):
    """DXGI_MAPPED_RECT"""

    __slots__ = ()
    _fields_ = (
        ("pitch", c_int32),
        ("bits", POINTER(c_byte)),
    )


class LUID(Structure):
    """LUID構造体。"""

    __slots__ = ()
    _fields_ = (
        ("low_part", c_uint32),
        ("high_part", c_int32),
    )


class DXGIAdapterDesc(Structure):
    """DXGI_ADAPTER_DESC構造体。"""

    __slots__ = ()
    _fields_ = (
        ("description", c_wchar * 128),
        ("vendor_id", c_uint32),
        ("subsystem_id", c_uint32),
        ("revision", c_uint32),
        ("dedicated_video_memory", c_size_t),
        ("dedicated_system_memory", c_size_t),
        ("shared_system_memory", c_size_t),
        ("adapter_luid", LUID),
    )
    if TYPE_CHECKING:
        description: str
        vendor_id: int
        subsystem_id: int
        revision: int
        dedicated_vide_memory: int
        dedicated_system_memory: int
        shared_system_memory: int
        adapater_luid: LUID


class DXGIOutputDesc(Structure):
    """DXGI_OUTPUT_DESC構造体。"""

    __slots__ = ()
    _fields_ = (
        ("device_name", c_wchar * 32),
        ("desktop_coordinates", RECT),
        ("attached_to_desktop", c_int32),
        ("_rotation", c_int32),
        ("monitor", c_void_p),
    )
    if TYPE_CHECKING:
        device_name: str
        desktop_coordinates: RECT
        attached_to_desktop: int
        _rotation: int
        monitor: int

    @property
    def rotation(self) -> DXGIModeRotation:
        return DXGIModeRotation(self._rotation)

    @rotation.setter
    def rotation(self, value: DXGIModeRotation) -> None:
        self._rotation = int(value)


class DXGISharedResource(Structure):
    """DXGI_SHARED_RESOURCE構造体。"""

    __slots__ = ()
    _fields_ = (("handle", c_void_p),)


class DXGIResourcePriority(IntEnum):
    """DXGI_RESOURCE_PRIORITY_*定数。"""

    MINIMUM = 0x28000000
    LOW = 0x50000000
    NORMAL = 0x78000000
    HIGH = 0xA0000000
    MAXIMUM = 0xC8000000


class DXGIResidency(IntEnum):
    """DXGI_RESIDENCY列挙型。"""

    FULLY_RESIDENT = 1
    RESIDENT_IN_SHARED_MEMORY = 2
    EVICTED_TO_DISK = 3


class DXGISurfaceDesc(Structure):
    """DXGI_SURFACE_DESC"""

    __slots__ = ()
    _fields_ = (
        ("width", c_uint32),
        ("height", c_uint32),
        ("_format", c_int32),
        ("_sample_desc", c_int32),
    )

    @property
    def format(self) -> DXGIFormat:
        return DXGIFormat(self._format)

    @format.setter
    def format(self, value: DXGIFormat) -> None:
        self._format = int(value)

    @property
    def sample_desc(self) -> DXGISampleDesc:
        return DXGISampleDesc(self._format)

    @sample_desc.setter
    def sample_desc(self, value: DXGISampleDesc) -> None:
        self._format = int(value)


class DXGISwapEffect(IntEnum):
    """DXGI_SWAP_EFFECT列挙型。"""

    DISCARD = 0
    SEQUENTIAL = 1
    FLIP_SEQUENTIAL = 3
    FLIP_DISCARD = 4


class DXGISwapChainFlag(IntEnum):
    """DXGI_SWAP_CHAIN_FLAG列挙型。"""

    NONPREROTATED = 1
    ALLOW_MODE_SWITCH = 2
    GDI_COMPATIBLE = 4
    RESTRICTED_CONTENT = 8
    RESTRICT_SHARED_RESOURCE_DRIVER = 16
    DISPLAY_ONLY = 32
    FRAME_LATENCY_WAITABLE_OBJECT = 64
    FOREGROUND_LAYER = 128
    FULLSCREEN_VIDEO = 256
    YUV_VIDEO = 512
    HW_PROTECTED = 1024
    ALLOW_TEARING = 2048
    RESTRICTED_TO_ALL_HOLOGRAPHIC_DISPLAYS = 4096


class DXGISwapChainDesc(Structure):
    """DXGI_SWAP_CHAIN_DESC構造体。"""

    __slots__ = ()
    _fields_ = (
        ("buffer_desc", DXGIModeDesc),
        ("sample_desc", DXGISampleDesc),
        ("_buffer_usage", c_int32),
        ("buffer_count", c_int32),
        ("output_window", c_void_p),
        ("windowed", c_int32),
        ("_swap_effect", c_int32),
        ("flags", c_int32),
    )

    @property
    def buffer_usage(self) -> DXGIUsage:
        return DXGIUsage(self._buffer_usage)

    @buffer_usage.setter
    def buffer_usage(self, value: DXGIUsage) -> None:
        self._buffer_usage = int(value)

    @property
    def swap_effect(self) -> DXGISwapEffect:
        return DXGISwapEffect(self._buffer_usage)

    @swap_effect.setter
    def swap_effect(self, value: DXGISwapEffect) -> None:
        self._buffer_usage = int(value)


class DXGIMap(IntFlag):
    """DXGI_MAP_*定数。"""

    READ = 1
    WRITE = 2
    DISCARD = 4


class DXGIEnumMode(IntFlag):
    """DXGI_ENUM_MODES_*定数。"""

    INTERLACED = 1
    DXGI_ENUM_MODES_SCALING = 2


DXGI_MAX_SWAP_CHAIN_BUFFERS = 16


class DXGIPresent(IntFlag):
    """DXGI_PRESENT_*定数。"""

    TEST = 0x00000001
    DO_NOT_SEQUENCE = 0x00000002
    RESTART = 0x00000004
    DO_NOT_WAIT = 0x00000008
    STEREO_PREFER_RIGHT = 0x00000010
    STEREO_TEMPORARY_MONO = 0x00000020
    RESTRICT_TO_OUTPUT = 0x00000040
    USE_DURATION = 0x00000100
    ALLOW_TEARING = 0x00000200


class DXGIAdapterFlag(IntFlag):
    """DXGI_ADAPTER_FLAG列挙型。"""

    NONE = 0
    REMOTE = 1
    SOFTWARE = 2


class DXGIAdapterDesc1(Structure):
    """DXGI_ADAPTER_DESC1構造体。"""

    __slots__ = ()
    _fields_ = (
        ("description", c_wchar * 128),
        ("vendor_id", c_uint32),
        ("subsystem_id", c_uint32),
        ("revision", c_uint32),
        ("dedicated_video_memory", c_size_t),
        ("dedicated_system_memory", c_size_t),
        ("shared_system_memory", c_size_t),
        ("adapter_luid", LUID),
        ("flags", c_uint32),
    )


class DXGIDisplayColorSpace(Structure):
    """DXGI_DISPLAY_COLOR_SPACE構造体"""

    __slots__ = ()
    _fields_ = (("primary_coordinates", (c_float * 2) * 8), ("white_points", (c_float * 16) * 2))


class IDXGIObject(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{aec22fb8-76f3-4639-9be0-28eb43a67a2e}")


class IDXGIDeviceSubObject(IUnknown):  # IDXGIObject
    __slots__ = ()
    _iid_ = GUID("{3d3e0379-f9de-4d58-bb6c-18d62992f1a6}")


class IDXGIResource(IUnknown):  # IDXGIDeviceSubObject
    __slots__ = ()
    _iid_ = GUID("{035f3ab4-482e-4e50-b41f-8a7f8bd8960b}")


class IDXGIKeyedMutex(IUnknown):  # IDXGIDeviceSubObject
    __slots__ = ()
    _iid_ = GUID("{9d8e1289-d7b3-465f-8126-250e349af85d}")


class IDXGISurface(IUnknown):  # IDXGIDeviceSubObject
    __slots__ = ()
    _iid_ = GUID("{cafcb56c-6ac3-4889-bf47-9e23bbd260ec}")


class IDXGISurface1(IUnknown):  # IDXGISurface
    __slots__ = ()
    _iid_ = GUID("{4AE63092-6327-4c1b-80AE-BFE12EA32B86}")


class IDXGIAdapter(IUnknown):  # IDXGIObject
    __slots__ = ()
    _iid_ = GUID("{2411e7e1-12ac-4ccf-bd14-9798e8534dc0}")


class IDXGIOutput(IUnknown):  # IDXGIObject
    __slots__ = ()
    _iid_ = GUID("{ae02eedb-c735-4690-8d52-5a8dc20213aa}")


class IDXGISwapChain(IUnknown):  # IDXGIDeviceSubObject
    __slots__ = ()
    _iid_ = GUID("{310d36a0-d2e7-4c0a-aa04-6a9d23b8886a}")


class IDXGIFactory(IUnknown):  # IDXGIObject
    __slots__ = ()
    _iid_ = GUID("{7b7166ec-21c7-44ae-b21a-c9ae321ae369}")


class IDXGIDevice(IUnknown):  # IDXGIObject
    __slots__ = ()
    _iid_ = GUID("{54ec77fa-1377-44e6-8c32-88fd5f44c84c}")


class IDXGIFactory1(IUnknown):  # IDXGIFactory
    __slots__ = ()
    _iid_ = GUID("{770aae78-f26f-4dba-a829-253c83d1b387}")


class IDXGIAdapter1(IUnknown):  # IDXGIAdapter
    __slots__ = ()
    _iid_ = GUID("{29038f61-3839-4626-91fd-086879011a05}")


class IDXGIDevice1(IUnknown):  # IDXGIDevice
    __slots__ = ()
    _iid_ = GUID("{77db970f-6276-48ba-ba28-070143b4392c}")
