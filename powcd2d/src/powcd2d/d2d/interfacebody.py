from ctypes import c_wchar_p
from ctypes.wintypes import RECT

from comtypes import STDMETHOD

from ..dwrite.types import (
    DWriteGlyphRun,
    IDWriteRenderingParams,
    IDWriteTextFormat,
    IDWriteTextLayout,
)
from ..wic.types import *
from .types import *

ID2D1Resource._methods_ = [
    STDMETHOD(c_int32, "GetFactory", (POINTER(ID2D1Factory),)),
]

ID2D1Image._methods_ = [*ID2D1Resource._methods_]

ID2D1Bitmap._methods_ = [
    *ID2D1Image._methods_,
    STDMETHOD(D2D1SizeF, "GetSize", ()),
    STDMETHOD(D2D1SizeU, "GetPixelSize", ()),
    STDMETHOD(D2D1PixelFormat, "GetPixelFormat", ()),
    STDMETHOD(None, "GetDpi", (POINTER(c_float), POINTER(c_float))),
    STDMETHOD(c_int32, "CopyFromBitmap", (POINTER(D2D1Point2U), POINTER(ID2D1Bitmap), POINTER(D2D1RectU))),
    STDMETHOD(c_int32, "CopyFromRenderTarget", (POINTER(D2D1Point2U), POINTER(ID2D1RenderTarget), POINTER(D2D1RectU))),
    STDMETHOD(c_int32, "CopyFromMemory", (POINTER(D2D1RectU), c_void_p, c_uint32)),
]

ID2D1GradientStopCollection._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(c_uint32, "GetGradientStopCount", ()),
    STDMETHOD(None, "GetGradientStops", (POINTER(D2D1GradientStop), c_uint32)),
    STDMETHOD(c_int32, "GetColorInterpolationGamma", ()),  # D2D1_GAMMA
    STDMETHOD(c_int32, "GetExtendMode", ()),  # D2D1_EXTEND_MODE
]

ID2D1Brush._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(None, "", ()),
    STDMETHOD(None, "SetOpacity", (c_float,)),
    STDMETHOD(None, "SetTransform", (POINTER(D2D1Matrix3X2F),)),
    STDMETHOD(c_float, "GetOpacity", ()),
    STDMETHOD(None, "GetTransform", (POINTER(D2D1Matrix3X2F),)),
]

ID2D1BitmapBrush._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(None, "SetExtendModeX", (c_int32,)),
    STDMETHOD(None, "SetExtendModeY", (c_int32,)),
    STDMETHOD(None, "SetInterpolationMode", (c_int32,)),
    STDMETHOD(None, "SetBitmap", (POINTER(ID2D1Bitmap),)),
    STDMETHOD(c_int32, "GetExtendModeX", ()),
    STDMETHOD(c_int32, "GetExtendModeY", ()),
    STDMETHOD(c_int32, "GetInterpolationMode", ()),
    STDMETHOD(None, "GetBitmap", (POINTER(ID2D1Bitmap),)),
]

ID2D1SolidColorBrush._methods_ = [
    *ID2D1BitmapBrush._methods_,
    STDMETHOD(None, "SetColor", (POINTER(D2D1ColorF),)),
    STDMETHOD(D2D1ColorF, "Getcolor", ()),
]

ID2D1LinearGradientBrush._methods_ = [
    *ID2D1BitmapBrush._methods_,
    STDMETHOD(None, "SetStartPoint", (D2D1Point2F,)),
    STDMETHOD(None, "SetEndPoint", (D2D1Point2F,)),
    STDMETHOD(D2D1Point2F, "GetStartPoint", ()),
    STDMETHOD(D2D1Point2F, "GetEndPoint", ()),
    STDMETHOD(None, "GetGradientStopCollection", (POINTER(ID2D1GradientStopCollection),)),
]

ID2D1RadialGradientBrush._methods_ = [
    *ID2D1BitmapBrush._methods_,
    STDMETHOD(None, "SetCenter", (D2D1Point2F,)),
    STDMETHOD(None, "SetGradientOriginOffset", (D2D1Point2F,)),
    STDMETHOD(None, "SetRadiusX", (c_float,)),
    STDMETHOD(None, "SetRadiusY", (c_float,)),
    STDMETHOD(D2D1Point2F, "GetCenter", ()),
    STDMETHOD(D2D1Point2F, "GetGradientOriginOffset", ()),
    STDMETHOD(c_float, "GetRadiusX", ()),
    STDMETHOD(c_float, "GetRadiusY", ()),
    STDMETHOD(None, "GetGradientStopCollection", (POINTER(ID2D1GradientStopCollection),)),
]

ID2D1StrokeStyle._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(c_int32, "GetStartCap", ()),
    STDMETHOD(c_int32, "GetEndCap", ()),
    STDMETHOD(c_int32, "GetDashCap", ()),
    STDMETHOD(c_float, "GetMiterLimit", ()),
    STDMETHOD(c_int32, "GetLineJoin", ()),
    STDMETHOD(c_float, "GetDashOffset", ()),
    STDMETHOD(c_int32, "GetDashStyle", ()),
    STDMETHOD(c_uint32, "GetDashesCount", ()),
    STDMETHOD(None, "GetDashes", (POINTER(c_float), c_uint32)),
]

ID2D1Geometry._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(c_int32, "GetBounds", (POINTER(D2D1Matrix3X2F), POINTER(D2D1RectF))),
    STDMETHOD(
        c_int32,
        "GetWidenedBounds",
        (c_float, POINTER(ID2D1StrokeStyle), POINTER(D2D1Matrix3X2F), c_float, POINTER(D2D1RectF)),
    ),
    STDMETHOD(
        c_int32,
        "StrokeContainsPoint",
        (D2D1Point2F, c_float, POINTER(ID2D1StrokeStyle), POINTER(D2D1Matrix3X2F), c_float, POINTER(c_int32)),
    ),
    STDMETHOD(c_int32, "FillContainsPoint", (D2D1Point2F, POINTER(D2D1Matrix3X2F), c_float, POINTER(c_int32))),
    STDMETHOD(
        c_int32,
        "CompareWithGeometry",
        (POINTER(ID2D1Geometry), POINTER(D2D1Matrix3X2F), c_float, POINTER(c_int32)),
    ),
    STDMETHOD(c_int32, "Simplify", (POINTER(D2D1Matrix3X2F), c_float, POINTER(ID2D1SimplifiedGeometrySink))),
    STDMETHOD(c_int32, "Tessellate", (POINTER(D2D1Matrix3X2F), c_float, POINTER(ID2D1TessellationSink))),
    STDMETHOD(
        c_int32,
        "CombineWithGeometry",
        (
            POINTER(ID2D1Geometry),
            c_int32,
            POINTER(D2D1Matrix3X2F),
            c_float,
            POINTER(ID2D1SimplifiedGeometrySink),
        ),
    ),
    STDMETHOD(c_int32, "Outline", (POINTER(D2D1Matrix3X2F), c_float, POINTER(ID2D1SimplifiedGeometrySink))),
    STDMETHOD(c_int32, "ComputeArea", (POINTER(D2D1Matrix3X2F), POINTER(c_float))),
    STDMETHOD(c_int32, "ComputeLength", (POINTER(D2D1Matrix3X2F), c_float, POINTER(c_float))),
    STDMETHOD(
        c_int32,
        "ComputePointAtLength",
        (POINTER(D2D1Matrix3X2F), c_float, POINTER(D2D1Point2F), POINTER(D2D1Point2F)),
    ),
    STDMETHOD(
        c_int32,
        "Widen",
        (
            c_float,
            POINTER(ID2D1StrokeStyle),
            POINTER(D2D1Matrix3X2F),
            c_float,
            POINTER(ID2D1SimplifiedGeometrySink),
        ),
    ),
]

ID2D1RectangleGeometry._methods_ = [
    *ID2D1Geometry._methods_,
    STDMETHOD(None, "GetRect", (POINTER(D2D1RectF),)),
]

ID2D1RoundedRectangleGeometry._methods_ = [
    *ID2D1Geometry._methods_,
    STDMETHOD(None, "GetRoundedRect", (POINTER(D2D1RoundedRect),)),
]

ID2D1EllipseGeometry._methods_ = [
    *ID2D1Geometry._methods_,
    STDMETHOD(None, "GetRoundedRect", (POINTER(D2D1Ellipse),)),
]

ID2D1GeometryGroup._methods_ = [
    *ID2D1Geometry._methods_,
    STDMETHOD(c_int32, "GetFillMode", ()),
    STDMETHOD(c_uint32, "GetSourceGeometryCount", ()),
    STDMETHOD(None, "GetsourceGeometries", (POINTER(POINTER(ID2D1Geometry)), c_uint32)),
]

ID2D1TransformedGeometry._methods_ = [
    *ID2D1Geometry._methods_,
    STDMETHOD(None, "GetSourceGeometry", (POINTER(POINTER(ID2D1Geometry)),)),
    STDMETHOD(None, "GetTransform", (POINTER(D2D1Matrix3X2F),)),
]

ID2D1SimplifiedGeometrySink._methods_ = [
    STDMETHOD(None, "SetFillMode", (c_int32,)),
    STDMETHOD(None, "SetSegmentFlags", (c_int32,)),
    STDMETHOD(None, "BeginFigure", (D2D1Point2F, c_int32)),
    STDMETHOD(None, "AddLines", (POINTER(D2D1Point2F), c_uint32)),
    STDMETHOD(None, "AddBeziers", (POINTER(D2D1BezierSegment), c_uint32)),
    STDMETHOD(None, "EndFigure", (c_int32,)),
    STDMETHOD(None, "Close", ()),
]

ID2D1GeometrySink._methods_ = [
    *ID2D1SimplifiedGeometrySink._methods_,
    STDMETHOD(None, "AddLine", (D2D1Point2F,)),
    STDMETHOD(None, "AddBezier", (POINTER(D2D1BezierSegment),)),
    STDMETHOD(None, "AddQuadraticBezier", (POINTER(D2D1QuardraticBezierSegment),)),
    STDMETHOD(None, "AddQuadraticBeziers", (POINTER(D2D1QuardraticBezierSegment), c_uint32)),
    STDMETHOD(None, "AddArc", (POINTER(D2D1ArgSegment),)),
]

ID2D1TessellationSink._methods_ = [
    STDMETHOD(None, "AddTriangles", (POINTER(D2D1Triangle), c_uint32)),
    STDMETHOD(c_int32, "Close", ()),
]

ID2D1PathGeometry._methods_ = [
    *ID2D1Geometry._methods_,
    STDMETHOD(c_int32, "Open", (POINTER(ID2D1GeometrySink),)),
    STDMETHOD(c_int32, "Stream", (POINTER(ID2D1GeometrySink),)),
    STDMETHOD(c_int32, "GetSegmentCount", (POINTER(c_uint32),)),
    STDMETHOD(c_int32, "GetFigureCount", (POINTER(c_uint32),)),
]

ID2D1Mesh._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(c_int32, "Open", (POINTER(ID2D1TessellationSink),)),
]

ID2D1Layer._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(D2D1SizeF, "GetSize", ()),
]

ID2D1DrawingStateBlock._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(None, "GetDescription", (POINTER(D2D1DrawingStateDesc),)),
    STDMETHOD(None, "SetDescription", (POINTER(D2D1DrawingStateDesc),)),
    STDMETHOD(None, "SetTextRenderingParams", (POINTER(IDWriteRenderingParams),)),
    STDMETHOD(None, "GetTextRenderingParams", (POINTER(POINTER(IDWriteRenderingParams)),)),
]

ID2D1RenderTarget._methods_ = [
    *ID2D1Resource._methods_,
    STDMETHOD(
        c_int32,
        "CreateBitmap",
        (D2D1SizeU, c_void_p, c_uint32, POINTER(D2D1BitmapProps), POINTER(POINTER(ID2D1Bitmap))),
    ),
    STDMETHOD(
        c_int32,
        "CreateBitmapFromWicBitmap",
        (POINTER(IWICBitmapSource), POINTER(D2D1BitmapProps), POINTER(POINTER(ID2D1Bitmap))),
    ),
    STDMETHOD(
        c_int32,
        "CreateSharedBitmap",
        (POINTER(GUID), c_void_p, POINTER(D2D1BitmapProps), POINTER(POINTER(ID2D1Bitmap))),
    ),
    STDMETHOD(
        c_int32,
        "CreateBitmapBrush",
        (
            POINTER(ID2D1Bitmap),
            POINTER(D2D1BitmapBrushProps),
            POINTER(D2D1BrushProps),
            POINTER(POINTER(ID2D1BitmapBrush)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CreateSolidColorBrush",
        (POINTER(D2D1ColorF), POINTER(D2D1BrushProps), POINTER(POINTER(ID2D1SolidColorBrush))),
    ),
    STDMETHOD(
        c_int32,
        "CreateGradientStopCollection",
        (POINTER(D2D1GradientStop), c_uint32, c_int32, c_int32, POINTER(POINTER(ID2D1GradientStopCollection))),
    ),
    STDMETHOD(
        c_int32,
        "CreateLinearGradientBrush",
        (
            POINTER(D2D1LinearGradientBrushProps),
            POINTER(D2D1BrushProps),
            POINTER(ID2D1GradientStopCollection),
            POINTER(POINTER(ID2D1LinearGradientBrush)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CreateRadialGradientBrush",
        (
            POINTER(D2D1RadialGradientBrushProps),
            POINTER(D2D1BrushProps),
            POINTER(ID2D1GradientStopCollection),
            POINTER(POINTER(ID2D1RadialGradientBrush)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CreateCompatibleRenderTarget",
        (
            POINTER(D2D1SizeF),
            POINTER(D2D1SizeU),
            POINTER(c_int32),
            POINTER(c_int32),
            POINTER(POINTER(ID2D1BitmapRenderTarget)),
        ),
    ),
    STDMETHOD(c_int32, "CreateLayer", (POINTER(D2D1SizeF), POINTER(POINTER(ID2D1Layer)))),
    STDMETHOD(c_int32, "CreateMesh", (POINTER(POINTER(ID2D1Mesh)),)),
    STDMETHOD(None, "DrawLine", (D2D1Point2F, D2D1Point2F, POINTER(ID2D1Brush), c_float, POINTER(ID2D1StrokeStyle))),
    STDMETHOD(None, "DrawRectangle", (POINTER(D2D1RectF), POINTER(ID2D1Brush), c_float, POINTER(ID2D1StrokeStyle))),
    STDMETHOD(None, "FillRectangle", (POINTER(D2D1RectF), POINTER(ID2D1Brush))),
    STDMETHOD(
        None,
        "DrawRoundedRectangle",
        (POINTER(D2D1RoundedRect), POINTER(ID2D1Brush), c_float, POINTER(ID2D1StrokeStyle)),
    ),
    STDMETHOD(None, "FillRoundedRectangle", (POINTER(D2D1RoundedRect), POINTER(ID2D1Brush))),
    STDMETHOD(None, "DrawEllipse", (POINTER(D2D1Ellipse), POINTER(ID2D1Brush), c_float, POINTER(ID2D1StrokeStyle))),
    STDMETHOD(None, "FillEllipse", (POINTER(D2D1Ellipse), POINTER(ID2D1Brush))),
    STDMETHOD(None, "DrawGeometry", (POINTER(ID2D1Geometry), POINTER(ID2D1Brush), c_float, POINTER(ID2D1StrokeStyle))),
    STDMETHOD(None, "FillGeometry", (POINTER(ID2D1Geometry), POINTER(ID2D1Brush), POINTER(ID2D1Brush))),
    STDMETHOD(None, "FillMesh", (POINTER(ID2D1Mesh), POINTER(ID2D1Brush))),
    STDMETHOD(
        None,
        "FillOpacityMask",
        (POINTER(ID2D1Bitmap), POINTER(ID2D1Brush), c_int32, POINTER(D2D1RectF), POINTER(D2D1RectF)),
    ),
    STDMETHOD(None, "DrawBitmap", (POINTER(ID2D1Bitmap), POINTER(D2D1RectF), c_float, c_int32, POINTER(D2D1RectF))),
    STDMETHOD(
        None,
        "DrawText",
        (
            c_wchar_p,
            c_uint32,
            POINTER(IDWriteTextFormat),
            POINTER(D2D1RectF),
            POINTER(ID2D1Brush),
            c_int32,
            c_int32,
        ),
    ),
    STDMETHOD(None, "DrawTextLayout", (D2D1Point2F, POINTER(IDWriteTextLayout), POINTER(ID2D1Brush), c_int32)),
    STDMETHOD(None, "DrawGlyphRun", (D2DPoint2F, POINTER(DWriteGlyphRun), POINTER(ID2D1Brush), c_int32)),
    STDMETHOD(None, "SetTransform", (POINTER(D2D1Matrix3X2F),)),
    STDMETHOD(None, "GetTransform", (POINTER(D2D1Matrix3X2F),)),
    STDMETHOD(None, "SetAntialiasMode", (c_int32,)),
    STDMETHOD(c_int32, "GetAntialiasMode", ()),
    STDMETHOD(None, "SetTextAntialiasMode", (c_int32,)),
    STDMETHOD(c_int32, "GetTextAntialiasMode", ()),
    STDMETHOD(None, "SetTextRenderingParams", (POINTER(IDWriteRenderingParams),)),
    STDMETHOD(None, "GetTextRenderingParams", (POINTER(POINTER(IDWriteRenderingParams)),)),
    STDMETHOD(None, "SetTags", (D2D1Tag, D2D1Tag)),
    STDMETHOD(None, "GetTags", (POINTER(D2D1Tag), POINTER(D2D1Tag))),
    STDMETHOD(None, "PushLayer", (POINTER(D2D1LayerParams), POINTER(ID2D1Layer))),
    STDMETHOD(None, "PopPayer", ()),
    STDMETHOD(c_int32, "Flush", (POINTER(D2D1Tag), POINTER(D2D1Tag))),
    STDMETHOD(None, "SaveDrawingState", (POINTER(ID2D1DrawingStateBlock),)),
    STDMETHOD(None, "RestoreDrawingState", (POINTER(ID2D1DrawingStateBlock),)),
    STDMETHOD(None, "PushAxisAlignedClip", (POINTER(D2D1RectF), c_int32)),
    STDMETHOD(None, "PopAxisAlignedClip", ()),
    STDMETHOD(None, "Clear", (POINTER(D2D1ColorF),)),
    STDMETHOD(None, "BeginDraw", ()),
    STDMETHOD(c_int32, "EndDraw", (POINTER(D2D1Tag), POINTER(D2D1Tag))),
    STDMETHOD(D2D1PixelFormat, "GetPixelFormat", ()),
    STDMETHOD(None, "SetDpi", (c_float, c_float)),
    STDMETHOD(None, "GetDpi", (POINTER(c_float), POINTER(c_float))),
    STDMETHOD(D2D1SizeF, "GetSize", ()),
    STDMETHOD(D2D1SizeU, "GetPixelSize", ()),
    STDMETHOD(c_uint32, "GetMaximumBitmapSize", ()),
    STDMETHOD(c_int32, "IsSupported", (POINTER(D2D1RenderTargetProps),)),
]

ID2D1BitmapRenderTarget_methods_ = [
    *ID2D1RenderTarget._methods_,
    STDMETHOD(c_int32, "GetBitmap", (POINTER(POINTER(ID2D1Bitmap)),)),
]
ID2D1HwndRenderTarget._methods_ = [
    *ID2D1RenderTarget._methods_,
    STDMETHOD(c_int32, "CheckWindowState", ()),
    STDMETHOD(c_int32, "Resize", (POINTER(D2D1SizeU),)),
    STDMETHOD(c_void_p, "GetHWnd", ()),
]
ID2D1GdiInteropRenderTarget._methods_ = [
    STDMETHOD(c_int32, "GetDC", (c_int32, POINTER(c_void_p))),
    STDMETHOD(c_int32, "ReleaseDC", (POINTER(RECT),)),
]
ID2D1DCRenderTarget._methods_ = [
    *ID2D1RenderTarget._methods_,
    STDMETHOD(c_int32, "BindDC", (c_void_p, POINTER(RECT))),
]

ID2D1Factory._methods_ = [
    STDMETHOD(c_int32, "ReloadSystemMetrics", ()),
    # [[deprecated("Deprecated. Use DisplayInformation::LogicalDpi for Windows Store Apps or GetDpiForWindow for desktop apps.")]]
    STDMETHOD(c_int32, "GetDesktopDpi", (POINTER(c_float), POINTER(c_float))),
    STDMETHOD(c_int32, "CreateRectangleGeometry", (POINTER(D2D1RectF), POINTER(POINTER(ID2D1RectangleGeometry)))),
    STDMETHOD(
        c_int32,
        "CreateRoundedRectangleGeometry",
        (POINTER(D2D1RoundedRect), POINTER(POINTER(ID2D1RoundedRectangleGeometry))),
    ),
    STDMETHOD(c_int32, "CreateEllipseGeometry", (POINTER(D2D1Ellipse), POINTER(POINTER(ID2D1EllipseGeometry)))),
    STDMETHOD(
        c_int32,
        "CreateGeometryGroup",
        (c_int32, POINTER(POINTER(ID2D1Geometry)), c_uint32, POINTER(POINTER(ID2D1GeometryGroup))),
    ),
    STDMETHOD(
        c_int32,
        "CreateTransformedGeometry",
        (POINTER(ID2D1Geometry), POINTER(D2D1Matrix3X2F), POINTER(POINTER(ID2D1TransformedGeometry))),
    ),
    STDMETHOD(c_int32, "CreatePathGeometry", (POINTER(POINTER(ID2D1PathGeometry)),)),
    STDMETHOD(
        c_int32,
        "CreateStrokeStyle",
        (POINTER(D2D1StrokeStyleProps), POINTER(c_float), c_uint32, POINTER(POINTER(ID2D1StrokeStyle))),
    ),
    STDMETHOD(
        c_int32,
        "CreateDrawingStateBlock",
        (
            POINTER(D2D1DrawingStateDesc),
            POINTER(IDWriteRenderingParams),
            POINTER(POINTER(ID2D1DrawingStateBlock)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CreateWicBitmapRenderTarget",
        (POINTER(IWICBitmap), POINTER(D2D1RenderTargetProps), POINTER(POINTER(ID2D1RenderTarget))),
    ),
    STDMETHOD(
        c_int32,
        "CreateHwndRenderTarget",
        (
            POINTER(D2D1RenderTargetProps),
            POINTER(D2D1HwndRenderTargetProps),
            POINTER(POINTER(ID2D1HwndRenderTarget)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CreateDxgiSurfaceRenderTarget",
        (POINTER(IDXGISurface), POINTER(D2D1RenderTargetProps), POINTER(POINTER(ID2D1RenderTarget))),
    ),
    STDMETHOD(
        c_int32,
        "CreateDCRenderTarget",
        (POINTER(D2D1RenderTargetProps), POINTER(POINTER(ID2D1DCRenderTarget))),
    ),
]
