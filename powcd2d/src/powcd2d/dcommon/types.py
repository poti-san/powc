from ctypes import Structure, Union, c_float, c_int32, c_uint32
from ctypes.wintypes import POINT, RECT
from enum import IntEnum, IntFlag


class DWriteMeasuringMode(IntEnum):
    """DWRITE_MEASURING_MODE"""

    NATURAL = 0
    GDI_CLASSIC = 1
    GDI_NATURAL = 2


class DWriteGlyphImageFormat(IntFlag):
    """DWRITE_GLYPH_IMAGE_FORMATS"""

    NONE = 0x00000000
    TRUETYPE = 0x00000001
    CFF = 0x00000002
    COLR = 0x00000004
    SVG = 0x00000008
    PNG = 0x00000010
    JPEG = 0x00000020
    TIFF = 0x00000040
    PREMULTIPLIED_B8G8R8A8 = 0x00000080
    COLR_PAINT_TREE = 0x00000100


class D2D1AlphaMode(IntEnum):
    """D2D1_ALPHA_MODE"""

    UNKNOWN = 0
    PREMULTIPLIED = 1
    STRAIGHT = 2
    IGNORE = 3


class D2D1PixelFormat(Structure):
    """D2D1_PIXEL_FORMAT"""

    __slots__ = ()
    _fields_ = (
        ("format", c_int32),  # DXGI_FORMAT
        ("alphamode", c_int32),  # D2D1_ALPHA_MODE
    )


class D2DPoint2U(Structure):
    """D2D_POINT_2U"""

    __slots__ = ()
    _fields_ = (("x", c_uint32), ("y", c_uint32))


D2DPoint2L = POINT
"""D2D_POINT_2L"""


class D2DPoint2F(Structure):
    """D2D_POINT_2F"""

    __slots__ = ()
    _fields_ = (("x", c_float), ("y", c_float))


class D2DVector2F(Structure):
    """D2D_VECTOR_2F"""

    __slots__ = ()
    _fields_ = (("x", c_float), ("y", c_float))


class D2DVector3F(Structure):
    """D2D_VECTOR_3F"""

    __slots__ = ()
    _fields_ = (("x", c_float), ("y", c_float), ("z", c_float))


class D2DVector4F(Structure):
    """D2D_VECTOR_4F"""

    __slots__ = ()
    _fields_ = (("x", c_float), ("y", c_float), ("z", c_float), ("w", c_float))


class D2DRectF(Structure):
    """D2D_RECT_F"""

    __slots__ = ()
    _fields_ = (("left", c_float), ("top", c_float), ("right", c_float), ("bottom", c_float))


class D2DRectU(Structure):
    """D2D_RECT_U"""

    __slots__ = ()
    _fields_ = (("left", c_uint32), ("top", c_uint32), ("right", c_uint32), ("bottom", c_uint32))


D2DRectL = RECT
"""D2D_RECT_L"""


class D2DSizeF(Structure):
    """D2D_SIZE_F"""

    __slots__ = ()
    _fields_ = (("width", c_float), ("height", c_float))


class D2DSizeU(Structure):
    """D2D_SIZE_U"""

    __slots__ = ()
    _fields_ = (("width", c_uint32), ("height", c_uint32))


class D2DMatrix3X2F(Union):
    """D2D_MATRIX_3X2_F"""

    class DUMMYSTRUCTURE1(Structure):
        __slots__ = ()
        _fields_ = (
            ("m11", c_float),
            ("m12", c_float),
            ("m21", c_float),
            ("dx", c_float),
            ("dy", c_float),
        )

    class DUMMYSTRUCTURE2(Structure):
        __slots__ = ()
        _fields_ = (
            ("_11", c_float),
            ("_12", c_float),
            ("_21", c_float),
            ("_22", c_float),
            ("_31", c_float),
            ("_32", c_float),
        )

    __slots__ = ()
    _anonymous_ = ("u1", "u2")
    _fields = (("u1", DUMMYSTRUCTURE1), ("u2", DUMMYSTRUCTURE2), ("m", (c_float * 2) * 3))


class D2DMatrix4X3F(Union):
    """D2D_MATRIX_4X3_F"""

    class DUMMYSTRUCTURE(Structure):
        __slots__ = ()
        _fields_ = (
            ("_11", c_float),
            ("_12", c_float),
            ("_13", c_float),
            ("_21", c_float),
            ("_22", c_float),
            ("_23", c_float),
            ("_31", c_float),
            ("_32", c_float),
            ("_33", c_float),
            ("_41", c_float),
            ("_42", c_float),
            ("_43", c_float),
        )

    __slots__ = ()
    _anonymous_ = ("u",)
    _fields = (("u", DUMMYSTRUCTURE), ("m", (c_float * 3) * 4))


class D2DMatrix4X4F(Union):
    """D2D_MATRIX_4X4_F"""

    class DUMMYSTRUCTURE(Structure):
        __slots__ = ()
        _fields_ = (
            ("_11", c_float),
            ("_12", c_float),
            ("_13", c_float),
            ("_14", c_float),
            ("_21", c_float),
            ("_22", c_float),
            ("_23", c_float),
            ("_24", c_float),
            ("_31", c_float),
            ("_32", c_float),
            ("_33", c_float),
            ("_34", c_float),
            ("_41", c_float),
            ("_42", c_float),
            ("_43", c_float),
            ("_44", c_float),
        )

    __slots__ = ()
    _anonymous_ = ("u",)
    _fields = (("u", DUMMYSTRUCTURE), ("m", (c_float * 4) * 4))


class D2DMatrix5X4F(Union):
    """D2D_MATRIX_5X4_F"""

    class DUMMYSTRUCTURE(Structure):
        __slots__ = ()
        _fields_ = (
            ("_11", c_float),
            ("_12", c_float),
            ("_13", c_float),
            ("_14", c_float),
            ("_21", c_float),
            ("_22", c_float),
            ("_23", c_float),
            ("_24", c_float),
            ("_31", c_float),
            ("_32", c_float),
            ("_33", c_float),
            ("_34", c_float),
            ("_41", c_float),
            ("_42", c_float),
            ("_43", c_float),
            ("_44", c_float),
            ("_51", c_float),
            ("_52", c_float),
            ("_53", c_float),
            ("_54", c_float),
        )

    __slots__ = ()
    _anonymous_ = ("u",)
    _fields = (("u", DUMMYSTRUCTURE), ("m", (c_float * 4) * 5))


D2D1Point2F = D2DPoint2F
"""D2D1_POINT_2F"""

D2D1Point2U = D2DPoint2U
"""D2D1_POINT_2U"""

D2D1Point2L = D2DPoint2L
"""D2D1_POINT_2L"""

D2D1RectF = D2DRectF
"""D2D1_RECT_F"""

D2D1RectU = D2DRectU
"""D2D1_RECT_U"""

D2D1RectL = D2DRectL
"""D2D1_RECT_L"""

D2D1SizeF = D2DSizeF
"""D2D1_SIZE_F"""

D2D1SizeU = D2DSizeU
"""D2D1_SIZE_U"""

D2D1Matrix3X2F = D2DMatrix3X2F
"""D2D1_MATRIX_3X2_F"""
