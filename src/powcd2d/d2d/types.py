"""Direct2DのCOMインターフェイス型と構造体の定義。

COMインターフェイスの型や構造体のみ必要な場合に直接インポートしてください。"""

# 注意
# このクラスではcID2D1Image -> ID2D1Resource -> IUnknownのような派生関係を継承で表しません。
# 代わりに継承元クラスの_methods_を派生クラスの_methods_内でアンパックします。
# インターフェイスの参照が複雑でエラーが生じるためです。

from ctypes import POINTER, c_float, c_int32, c_uint64, c_void_p
from enum import IntEnum, IntFlag

from comtypes import GUID, IUnknown

from ..dcommon.types import (
    D2D1PixelFormat,
    D2DMatrix3X2F,
    D2DPoint2F,
    D2DPoint2U,
    D2DRectF,
    D2DRectU,
    D2DSizeF,
    D2DSizeU,
)
from ..dxgi.types import *

# ---------------------------------------------------------------------------------------------------
# COMインターフェイスのメソッド以外
# ---------------------------------------------------------------------------------------------------


class ID2D1Resource(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{2cd90691-12e2-11dc-9fed-001143a055f9}")


class ID2D1Image(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{65019f75-8da2-497c-b32c-dfa34e48ede6}")


class ID2D1Bitmap(IUnknown):  # ID2D1Image
    __slots__ = ()
    _iid_ = GUID("{a2296057-ea42-4099-983b-539fb6505426}")


class ID2D1GradientStopCollection(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd906a7-12e2-11dc-9fed-001143a055f9}")


class ID2D1Brush(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd906a8-12e2-11dc-9fed-001143a055f9}")


class ID2D1BitmapBrush(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd906aa-12e2-11dc-9fed-001143a055f9}")


class ID2D1SolidColorBrush(IUnknown):  # ID2D1Brush
    __slots__ = ()
    _iid_ = GUID("{2cd906a9-12e2-11dc-9fed-001143a055f9}")


class ID2D1LinearGradientBrush(IUnknown):  # ID2D1Brush
    __slots__ = ()
    _iid_ = GUID("{2cd906ab-12e2-11dc-9fed-001143a055f9}")


class ID2D1RadialGradientBrush(IUnknown):  # ID2D1Brush
    __slots__ = ()
    _iid_ = GUID("{2cd906ac-12e2-11dc-9fed-001143a055f9}")


class ID2D1StrokeStyle(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd9069d-12e2-11dc-9fed-001143a055f9}")


class ID2D1Geometry(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd906a1-12e2-11dc-9fed-001143a055f9}")


class ID2D1RectangleGeometry(IUnknown):  # ID2D1Geometry
    __slots__ = ()
    _iid_ = GUID("{2cd906a2-12e2-11dc-9fed-001143a055f9}")


class ID2D1RoundedRectangleGeometry(IUnknown):  # ID2D1Geometry
    __slots__ = ()
    _iid_ = GUID("{2cd906a3-12e2-11dc-9fed-001143a055f9}")


class ID2D1EllipseGeometry(IUnknown):  # ID2D1Geometry
    __slots__ = ()
    _iid_ = GUID("{2cd906a4-12e2-11dc-9fed-001143a055f9}")


class ID2D1GeometryGroup(IUnknown):  # ID2D1Geometry
    __slots__ = ()
    _iid_ = GUID("{2cd906a6-12e2-11dc-9fed-001143a055f9}")


class ID2D1TransformedGeometry(IUnknown):  # ID2D1Geometry
    __slots__ = ()
    _iid_ = GUID("{2cd906bb-12e2-11dc-9fed-001143a055f9}")


class ID2D1SimplifiedGeometrySink(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{2cd9069e-12e2-11dc-9fed-001143a055f9}")


class ID2D1GeometrySink(IUnknown):  # ID2D1SimplifiedGeometrySink
    __slots__ = ()
    _iid_ = GUID("{2cd9069f-12e2-11dc-9fed-001143a055f9}")


class ID2D1TessellationSink(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{2cd906c1-12e2-11dc-9fed-001143a055f9}")


class ID2D1PathGeometry(IUnknown):  # ID2D1Geometry
    __slots__ = ()
    _iid_ = GUID("{2cd906a5-12e2-11dc-9fed-001143a055f9}")


class ID2D1Mesh(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd906c2-12e2-11dc-9fed-001143a055f9}")


class ID2D1Layer(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd9069b-12e2-11dc-9fed-001143a055f9}")


class ID2D1DrawingStateBlock(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{28506e39-ebf6-46a1-bb47-fd85565ab957}")


class ID2D1RenderTarget(IUnknown):  # ID2D1Resource
    __slots__ = ()
    _iid_ = GUID("{2cd90694-12e2-11dc-9fed-001143a055f9}")


class ID2D1BitmapRenderTarget(IUnknown):  # ID2D1RenderTarget
    __slots__ = ()
    _iid_ = GUID("{2cd90695-12e2-11dc-9fed-001143a055f9}")


class ID2D1HwndRenderTarget(IUnknown):  # ID2D1RenderTarget
    __slots__ = ()
    _iid_ = GUID("{2cd90698-12e2-11dc-9fed-001143a055f9}")


class ID2D1GdiInteropRenderTarget(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{e0db51c3-6f77-4bae-b3d5-e47509b35838}")


class ID2D1DCRenderTarget(IUnknown):  # ID2D1RenderTarget
    __slots__ = ()
    _iid_ = GUID("{1c51bc64-de61-46fd-9899-63a5d8f03950}")


class ID2D1Factory(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{06152247-6f50-465a-9245-118bfd3b6007}")


IDWriteGeometrySink = ID2D1SimplifiedGeometrySink

# ---------------------------------------------------------------------------------------------------
# COMインターフェイス以外の型
# ---------------------------------------------------------------------------------------------------

# region
D2D1_INVALID_TAG = 0xFFFFFFFFFFFFFFFF
D2D1_DEFAULT_FLATTENING_TOLERANCE = 0.25


class D2D1InterpolationModeDefinition(IntEnum):
    """D2D1_INTERPOLATION_MODE_DEFINITION。"""

    NEAREST_NEIGHBOR = 0
    LINEAR = 1
    CUBIC = 2
    MULTI_SAMPLE_LINEAR = 3
    ANISOTROPIC = 4
    HIGH_QUALITY_CUBIC = 5
    FANT = 6
    MIPMAP_LINEAR = 7


class D2D1Gamma(IntEnum):
    """D2D1_GAMMA。"""

    _2_2 = 0
    _1_0 = 1


class D2D1OpacityMaskContent(IntEnum):
    """D2D1_OPACITY_MASK_CONTENT。"""

    GRAPHICS = 0
    TEXT_NATURAL = 1
    TEXT_GDI_COMPATIBLE = 2


class D2D1ExtendMode(IntEnum):
    """D2D1_EXTEND_MODE。"""

    CLAMP = 0
    WRAP = 1
    MIRROR = 2


class D2D1AntialiasMode(IntEnum):
    """D2D1_ANTIALIAS_MODE。"""

    PER_PRIMITIVE = 0
    ALIASED = 1


class D2D1TextAntialiasMode(IntEnum):
    """D2D1_TEXT_ANTIALIAS_MODE。"""

    DEFAULT = 0
    CLEARTYPE = 1
    GRAYSCALE = 2
    ALIASED = 3


class D2D1BitmapInterpolationMode(IntEnum):
    """D2D1_BITMAP_INTERPOLATION_MODE。"""

    NEAREST_NEIGHBOR = D2D1InterpolationModeDefinition.NEAREST_NEIGHBOR
    LINEAR = D2D1InterpolationModeDefinition.LINEAR


class D2D1DrawTextOption(IntEnum):
    """D2D1_DRAW_TEXT_OPTIONS"""

    NO_SNAP = 0x00000001
    CLIP = 0x00000002
    ENABLE_COLOR_FONT = 0x00000004
    DISABLE_COLOR_BITMAP_SNAPPING = 0x00000008
    NONE = 0x00000000


D2D1Point2U = D2DPoint2U
D2D1Point2F = D2DPoint2F
D2D1RectF = D2DRectF
D2D1RectU = D2DRectU
D2D1SizeF = D2DSizeF
D2D1SizeU = D2DSizeU
D2D1ColorF = D2DColorF
D2D1Matrix3X2F = D2DMatrix3X2F
D2D1Tag = c_uint64


class D2D1BitmapProps(Structure):
    """D2D1_BITMAP_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("pixel_format", D2D1PixelFormat),
        ("dpi_x", c_float),
        ("dpi_y", c_float),
    )


class D2D1GradientStop(Structure):
    """D2D1_GRADIENT_STOP"""

    __slots__ = ()
    _fields_ = (
        ("position", c_float),
        ("color", D2D1ColorF),
    )


class D2D1BrushProps(Structure):
    """D2D1_BRUSH_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("opacity", c_float),
        ("transform", D2D1Matrix3X2F),
    )


class D2D1BitmapBrushProps(Structure):
    """D2D1_BITMAP_BRUSH_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("extend_mode_x", c_int32),  # D2D1_EXTEND_MODE
        ("extend_mode_y", c_int32),  # D2D1_EXTEND_MODE
        ("interpolation_mode", c_int32),  # D2D1_BITMAP_INTERPOLATION_MODE
    )


class D2D1LinearGradientBrushProps(Structure):
    """D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("start_point", D2D1Point2F),
        ("end_point", D2D1Point2F),
    )


class D2D1RadialGradientBrushProps(Structure):
    """D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("center", D2D1Point2F),
        ("gradient_origin_offset", D2D1Point2F),
        ("radius_x", c_float),
        ("radius_y", c_float),
    )


class D2D1ArcSize(IntEnum):
    """D2D1_ARC_SIZE"""

    SMALL = 0
    LARGE = 1


class D2D1CapStyle(IntEnum):
    """D2D1_CAP_STYLE"""

    FLAT = 0
    SQUARE = 1
    ROUND = 2
    TRIANGLE = 3


class D2D1DashStyle(IntEnum):
    """D2D1_DASH_STYLE"""

    SOLID = 0
    DASH = 1
    DOT = 2
    DASH_DOT = 3
    DASH_DOT_DOT = 4
    CUSTOM = 5


class D2D1LineJoin(IntEnum):
    """D2D1_LINE_JOIN"""

    MITER = 0
    BEVEL = 1
    ROUND = 2
    MITER_OR_BEVEL = 3


class D2D1CombineMode(IntEnum):
    """D2D1_COMBINE_MODE"""

    UNION = 0
    INTERSECT = 1
    XOR = 2
    EXCLUDE = 3


class D2D1GeometryRelation(IntEnum):
    """D2D1_GEOMETRY_RELATION"""

    UNKNOWN = 0
    DISJOINT = 1
    IS_CONTAINED = 2
    CONTAINS = 3
    OVERLAP = 4


class D2D1GeometrySimplificationOption(IntEnum):
    """D2D1_GEOMETRY_SIMPLIFICATION_OPTION"""

    CUBICS_AND_LINES = 0
    LINES = 1


class D2D1FigureBegin(IntEnum):
    """D2D1_FIGURE_BEGIN"""

    FILLED = 0
    HOLLOW = 1


class D2D1FigureEnd(IntEnum):
    """D2D1_FIGURE_END"""

    OPEN = 0
    CLOSED = 1


class D2D1BezierSegment(Structure):
    """D2D1_BEZIER_SEGMENT"""

    __slots__ = ()
    _fields_ = (
        ("point1", D2D1Point2F),
        ("point2", D2D1Point2F),
        ("point3", D2D1Point2F),
    )


class D2D1Triangle(Structure):
    """D2D1_TRIANGLE"""

    __slots__ = ()
    _fields_ = (
        ("point1", D2D1Point2F),
        ("point2", D2D1Point2F),
        ("point3", D2D1Point2F),
    )


class D2D1PathSegment(IntFlag):
    """D2D1_PATH_SEGMENT"""

    NONE = 0x00000000
    FORCE_UNSTROKED = 0x00000001
    FORCE_ROUND_LINE_JOIN = 0x00000002


class D2D1SweepDirection(IntEnum):
    """D2D1_SWEEP_DIRECTION"""

    COUNTER_CLOCKWISE = 0
    CLOCKWISE = 1


class D2D1FillMode(IntEnum):
    """D2D1_FILL_MODE"""

    ALTERNATE = 0
    WINDING = 1


class D2D1ArgSegment(Structure):
    """D2D1_ARC_SEGMENT"""

    __slots__ = ()
    _fields_ = (
        ("point", D2D1Point2F),
        ("size", D2D1SizeF),
        ("rotation_angle", c_float),
        ("sweep_direction", c_int32),  # D2D1_SWEEP_DIRECTION
        ("arc_size", c_int32),  # D2D1_ARC_SIZE
    )


class D2D1QuardraticBezierSegment(Structure):
    """D2D1_QUADRATIC_BEZIER_SEGMENT"""

    __slots__ = ()
    _fields_ = (
        ("point1", D2D1Point2F),
        ("point2", D2D1Point2F),
    )


class D2D1Ellipse(Structure):
    """D2D1_ELLIPSE"""

    __slots__ = ()
    _fields_ = (
        ("point", D2D1Point2F),
        ("radius_x", c_float),
        ("radius_y", c_float),
    )


class D2D1RoundedRect(Structure):
    """D2D1_ROUNDED_RECT"""

    __slots__ = ()
    _fields_ = (
        ("rect", D2D1RectF),
        ("radius_x", c_float),
        ("radius_y", c_float),
    )


class D2D1StrokeStyleProps(Structure):
    """D2D1_STROKE_STYLE_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("start_cap", c_int32),  # D2D1_CAP_STYLE
        ("end_cap", c_int32),  # D2D1_CAP_STYLE
        ("dash_cap", c_int32),  # D2D1_CAP_STYLE
        ("line_join", c_int32),  # D2D1_LINE_JOIN
        ("miter_limit", c_float),
        ("dash_style", c_int32),  # D2D1_DASH_STYLE
        ("dash_offset", c_float),
    )


class D2D1LayerOption(IntFlag):
    """D2D1_LAYER_OPTIONS"""

    NONE = 0x00000000
    INITIALIZE_FOR_CLEARTYPE = 0x00000001


class D2D1LayerParams(Structure):
    """D2D1_LAYER_PARAMETERS"""

    __slots__ = ()
    _fields_ = (
        ("content_bounds", D2D1RectF),
        ("geometric_mask", POINTER(ID2D1Geometry)),
        ("mask_antialias_mode", c_int32),
        ("mask_transform", D2D1Matrix3X2F),
        ("opacity", c_float),
        ("opacity_brush", POINTER(ID2D1Brush)),
        ("layer_options", c_int32),  # D2D1_LAYER_OPTIONS
    )


class D2D1WindowState(IntFlag):
    """D2D1_WINDOW_STATE"""

    NONE = 0x0000000
    OCCLUDED = 0x0000001


class D2D1RenderTargetType(IntEnum):
    """D2D1_RENDER_TARGET_TYPE"""

    DEFAULT = 0
    SOFTWARE = 1
    HARDWARE = 2


class D2D1FeatureLevel(IntEnum):
    """D2D1_FEATURE_LEVEL"""

    DEFAULT = 0
    _9 = 0x9100  # D3D_FEATURE_LEVEL_9_1
    _10 = 0xA000  # D3D_FEATURE_LEVEL_10_0


class D2D1RenderTargetUsage(IntFlag):
    """D2D1_RENDER_TARGET_USAGE"""

    NONE = 0x00000000
    FORCE_BITMAP_REMOTING = 0x00000001
    GDI_COMPATIBLE = 0x00000002


class D2D1PresentOption(IntFlag):
    """D2D1_PRESENT_OPTIONS"""

    NONE = 0x00000000
    RETAIN_CONTENTS = 0x00000001
    IMMEDIATELY = 0x00000002


class D2D1RenderTargetProps(Structure):
    """D2D1_RENDER_TARGET_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("type", c_int32),  # D2D1_RENDER_TARGET_TYPE
        ("pixel_format", c_int32),  # D2D1_PIXEL_FORMAT
        ("dpi_x", c_float),
        ("dpi_y", c_float),
        ("usage", c_int32),  # D2D1_RENDER_TARGET_USAGE
        ("min_level", c_int32),  # D2D1_FEATURE_LEVEL
    )


class D2D1HwndRenderTargetProps(Structure):
    """D2D1_HWND_RENDER_TARGET_PROPERTIES"""

    __slots__ = ()
    _fields_ = (
        ("hwnd", c_void_p),
        ("pixel_size", D2D1SizeU),
        ("present_options", c_int32),  # D2D1_PRESENT_OPTIONS
    )


class D2D1CompatibleRenderTargetOption(IntFlag):
    """D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS"""

    NONE = 0x00000000
    COMPATIBLE = 0x00000001


class D2D1DrawingStateDesc(Structure):
    """D2D1_DRAWING_STATE_DESCRIPTION"""

    __slots__ = ()
    _fields_ = (
        ("antialias_mode", c_int32),  # D2D1_ANTIALIAS_MODE
        ("text_antialias_mode", c_int32),  # D2D1_TEXT_ANTIALIAS_MODE
        ("tag1", D2D1Tag),
        ("tag2", D2D1Tag),
        ("transform", D2D1Matrix3X2F),
    )


class D2D1DCInitializeMode(IntEnum):
    """D2D1_DC_INITIALIZE_MODE"""

    COPY = 0
    CLEAR = 1


class D2D1DebugLevel(IntEnum):
    """D2D1_DEBUG_LEVEL"""

    NONE = 0
    ERROR = 1
    WARNING = 2
    INFORMATION = 3


class D2D1FactoryType(IntEnum):
    """D2D1_FACTORY_TYPE"""

    SINGLE_THREADED = 0
    MULTI_THREADED = 1


class D2D1FactoryOption(Structure):
    """D2D1_FACTORY_OPTIONS"""

    __slots__ = ()
    _fields_ = (("debug_level", c_int32),)  # D2D1_DEBUG_LEVEL


# endregion
