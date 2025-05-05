from ctypes import (
    POINTER,
    Structure,
    c_float,
    c_int16,
    c_int32,
    c_uint8,
    c_uint16,
    c_uint32,
    c_void_p,
    c_wchar_p,
)
from enum import IntEnum, IntFlag

from comtypes import GUID, IUnknown


class DWriteFontFileType(IntEnum):
    """DWRITE_FONT_FILE_TYPE"""

    UNKNOWN = 0
    CFF = 1
    TRUETYPE = 2
    OPENTYPE_COLLECTION = 3
    TYPE1_PFM = 4
    TYPE1_PFB = 5
    VECTOR = 6
    BITMAP = 7
    TRUETYPE_COLLECTION = OPENTYPE_COLLECTION


class DWriteFontFaceType(IntEnum):
    """DWRITE_FONT_FACE_TYPE"""

    CFF = 0
    TRUETYPE = 1
    OPENTYPE_COLLECTION = 2
    TYPE1 = 3
    VECTOR = 4
    BITMAP = 5
    UNKNOWN = 6
    RAW_CFF = 7
    TRUETYPE_COLLECTION = OPENTYPE_COLLECTION


class DWriteFontSimulation(IntFlag):
    """DWRITE_FONT_SIMULATIONS"""

    NONE = 0x0000
    BOLD = 0x0001
    OBLIQUE = 0x0002


class DWriteFontWeight(IntEnum):
    """DWRITE_FONT_WEIGHT"""

    THIN = 100
    EXTRA_LIGHT = 200
    ULTRA_LIGHT = 200
    LIGHT = 300
    SEMI_LIGHT = 350
    NORMAL = 400
    REGULAR = 400
    MEDIUM = 500
    DEMI_BOLD = 600
    SEMI_BOLD = 600
    BOLD = 700
    EXTRA_BOLD = 800
    ULTRA_BOLD = 800
    BLACK = 900
    HEAVY = 900
    EXTRA_BLACK = 950
    ULTRA_BLACK = 950


class DWriteFontStretch(IntEnum):
    """DWRITE_FONT_STRETCH"""

    UNDEFINED = 0
    ULTRA_CONDENSED = 1
    EXTRA_CONDENSED = 2
    CONDENSED = 3
    SEMI_CONDENSED = 4
    NORMAL = 5
    MEDIUM = 5
    SEMI_EXPANDED = 6
    EXPANDED = 7
    EXTRA_EXPANDED = 8
    ULTRA_EXPANDED = 9


class DWriteFontStyle(IntEnum):
    """DWRITE_FONT_STYLE"""

    NORMAL = 0
    OBLIQUE = 1
    ITALIC = 2


class DWriteInfoStringID(IntEnum):
    """DWRITE_INFORMATIONAL_STRING_ID"""

    NONE = 0
    COPYRIGHT_NOTICE = 1
    VERSION_STRINGS = 2
    TRADEMARK = 3
    MANUFACTURER = 4
    DESIGNER = 5
    DESIGNER_URL = 6
    DESCRIPTION = 7
    FONT_VENDOR_URL = 8
    LICENSE_DESCRIPTION = 9
    LICENSE_INFO_URL = 10
    WIN32_FAMILY_NAMES = 11
    WIN32_SUBFAMILY_NAMES = 12
    TYPOGRAPHIC_FAMILY_NAMES = 13
    TYPOGRAPHIC_SUBFAMILY_NAMES = 14
    SAMPLE_TEXT = 15
    FULL_NAME = 16
    POSTSCRIPT_NAME = 17
    POSTSCRIPT_CID_NAME = 18
    WEIGHT_STRETCH_STYLE_FAMILY_NAME = 19
    DESIGN_SCRIPT_LANGUAGE_TAG = 20
    SUPPORTED_SCRIPT_LANGUAGE_TAG = 21
    #     // Obsolete aliases kept to avoid breaking existing code.
    #     PREFERRED_FAMILY_NAMES = TYPOGRAPHIC_FAMILY_NAMES
    #     PREFERRED_SUBFAMILY_NAMES = TYPOGRAPHIC_SUBFAMILY_NAMES
    #     WWS_FAMILY_NAME = WEIGHT_STRETCH_STYLE_FAMILY_NAME


class DWriteFontMetrics(Structure):
    """DWRITE_FONT_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("design_units_per_em", c_uint16),
        ("ascent", c_uint16),
        ("descent", c_uint16),
        ("line_gap", c_int16),
        ("cap_height", c_uint16),
        ("x_height", c_uint16),
        ("underline_position", c_int16),
        ("underline_thickness", c_uint16),
        ("strikethrough_position", c_int16),
        ("stridethrough_thickness", c_uint16),
    )


class DWriteGlyphMetrics(Structure):
    """DWRITE_GLYPH_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("left_side_bearing", c_int32),
        ("advance_width", c_uint32),
        ("right_side_bearing", c_int32),
        ("top_side_bearing", c_int32),
        ("advance_height", c_uint32),
        ("bottom_side_bearing", c_int32),
        ("vertical_origin_y", c_int32),
    )


class DWriteGlyphOffset(Structure):
    """DWRITE_GLYPH_OFFSET"""

    __slots__ = ()
    _fields_ = (
        ("advance_offset", c_float),
        ("ascender_offset", c_float),
    )


class DWriteFactoryType(IntEnum):
    """DWRITE_FACTORY_TYPE"""

    SHARED = 0
    ISOLATED = 1


class DWritePixelGeometry(IntEnum):
    """DWRITE_PIXEL_GEOMETRY"""

    FLAT = 0
    RGB = 1
    BGR = 2


class DWriteRenderingMode(IntEnum):
    """DWRITE_RENDERING_MODE"""

    DEFAULT = 0
    ALIASED = 1
    GDI_CLASSIC = 2
    GDI_NATURAL = 3
    NATURAL = 4
    NATURAL_SYMMETRIC = 5
    OUTLINE = 6
    # // obsolete
    # DWRITE_RENDERING_MODE_CLEARTYPE_GDI_CLASSIC         = DWRITE_RENDERING_MODE_GDI_CLASSIC,
    # DWRITE_RENDERING_MODE_CLEARTYPE_GDI_NATURAL         = DWRITE_RENDERING_MODE_GDI_NATURAL,
    # DWRITE_RENDERING_MODE_CLEARTYPE_NATURAL             = DWRITE_RENDERING_MODE_NATURAL,
    # DWRITE_RENDERING_MODE_CLEARTYPE_NATURAL_SYMMETRIC   = DWRITE_RENDERING_MODE_NATURAL_SYMMETRIC


class DWriteMatrix(Structure):
    """DWRITE_MATRIX"""

    __slots__ = ()
    _fields_ = (
        ("m11", c_float),
        ("m12", c_float),
        ("m21", c_float),
        ("m22", c_float),
        ("dx", c_float),
        ("dy", c_float),
    )


class DWriteReadingDirection(IntEnum):
    """DWRITE_READING_DIRECTION"""

    LEFT_TO_RIGHT = 0
    RIGHT_TO_LEFT = 1
    TOP_TO_BOTTOM = 2
    BOTTOM_TO_TOP = 3


class DWriteFlowDirection(IntEnum):
    """DWRITE_FLOW_DIRECTION"""

    TOP_TO_BOTTOM = 0
    BOTTOM_TO_TOP = 1
    LEFT_TO_RIGHT = 2
    RIGHT_TO_LEFT = 3


class DWriteTextAlign(IntEnum):
    """DWRITE_TEXT_ALIGNMENT"""

    LEADING = 0
    TRAILING = 1
    CENTER = 2
    JUSTIFIED = 3


class DWriteParagraphAlign(IntEnum):
    """DWRITE_PARAGRAPH_ALIGNMENT"""

    NEAR = 0
    FAR = 1
    CENTER = 2


class DWriteWordWrap(IntEnum):
    """DWRITE_WORD_WRAPPING"""

    WRAP = 0
    NO_WRAP = 1
    EMERGENCY_BREAK = 2
    WHOLE_WORD = 3
    CHARACTER = 4


class DWriteLineSpacingMethod(IntEnum):
    """DWRITE_LINE_SPACING_METHOD"""

    DEFAULT = 0
    UNIFORM = 1
    PROPORTIONAL = 2


class DWriteTrimmingGranularity(IntEnum):
    """DWRITE_TRIMMING_GRANULARITY"""

    NONE = 0
    CHARACTER = 1
    WORD = 2


def dwrite_make_opentype_tag(fourcc_str: str) -> int:
    return int.from_bytes(fourcc_str.encode("ascii"), byteorder="little")


class DWriteFontFeatureTag(IntEnum):
    """DWRITE_FONT_FEATURE_TAG"""

    ALTERNATIVE_FRACTIONS = dwrite_make_opentype_tag("afrc")
    PETITE_CAPITALS_FROM_CAPITALS = dwrite_make_opentype_tag("c2pc")
    SMALL_CAPITALS_FROM_CAPITALS = dwrite_make_opentype_tag("c2sc")
    CONTEXTUAL_ALTERNATES = dwrite_make_opentype_tag("calt")
    CASE_SENSITIVE_FORMS = dwrite_make_opentype_tag("case")
    GLYPH_COMPOSITION_DECOMPOSITION = dwrite_make_opentype_tag("ccmp")
    CONTEXTUAL_LIGATURES = dwrite_make_opentype_tag("clig")
    CAPITAL_SPACING = dwrite_make_opentype_tag("cpsp")
    CONTEXTUAL_SWASH = dwrite_make_opentype_tag("cswh")
    CURSIVE_POSITIONING = dwrite_make_opentype_tag("curs")
    DEFAULT = dwrite_make_opentype_tag("dflt")
    DISCRETIONARY_LIGATURES = dwrite_make_opentype_tag("dlig")
    EXPERT_FORMS = dwrite_make_opentype_tag("expt")
    FRACTIONS = dwrite_make_opentype_tag("frac")
    FULL_WIDTH = dwrite_make_opentype_tag("fwid")
    HALF_FORMS = dwrite_make_opentype_tag("half")
    HALANT_FORMS = dwrite_make_opentype_tag("haln")
    ALTERNATE_HALF_WIDTH = dwrite_make_opentype_tag("halt")
    HISTORICAL_FORMS = dwrite_make_opentype_tag("hist")
    HORIZONTAL_KANA_ALTERNATES = dwrite_make_opentype_tag("hkna")
    HISTORICAL_LIGATURES = dwrite_make_opentype_tag("hlig")
    HALF_WIDTH = dwrite_make_opentype_tag("hwid")
    HOJO_KANJI_FORMS = dwrite_make_opentype_tag("hojo")
    JIS04_FORMS = dwrite_make_opentype_tag("jp04")
    JIS78_FORMS = dwrite_make_opentype_tag("jp78")
    JIS83_FORMS = dwrite_make_opentype_tag("jp83")
    JIS90_FORMS = dwrite_make_opentype_tag("jp90")
    KERNING = dwrite_make_opentype_tag("kern")
    STANDARD_LIGATURES = dwrite_make_opentype_tag("liga")
    LINING_FIGURES = dwrite_make_opentype_tag("lnum")
    LOCALIZED_FORMS = dwrite_make_opentype_tag("locl")
    MARK_POSITIONING = dwrite_make_opentype_tag("mark")
    MATHEMATICAL_GREEK = dwrite_make_opentype_tag("mgrk")
    MARK_TO_MARK_POSITIONING = dwrite_make_opentype_tag("mkmk")
    ALTERNATE_ANNOTATION_FORMS = dwrite_make_opentype_tag("nalt")
    NLC_KANJI_FORMS = dwrite_make_opentype_tag("nlck")
    OLD_STYLE_FIGURES = dwrite_make_opentype_tag("onum")
    ORDINALS = dwrite_make_opentype_tag("ordn")
    PROPORTIONAL_ALTERNATE_WIDTH = dwrite_make_opentype_tag("palt")
    PETITE_CAPITALS = dwrite_make_opentype_tag("pcap")
    PROPORTIONAL_FIGURES = dwrite_make_opentype_tag("pnum")
    PROPORTIONAL_WIDTHS = dwrite_make_opentype_tag("pwid")
    QUARTER_WIDTHS = dwrite_make_opentype_tag("qwid")
    REQUIRED_LIGATURES = dwrite_make_opentype_tag("rlig")
    RUBY_NOTATION_FORMS = dwrite_make_opentype_tag("ruby")
    STYLISTIC_ALTERNATES = dwrite_make_opentype_tag("salt")
    SCIENTIFIC_INFERIORS = dwrite_make_opentype_tag("sinf")
    SMALL_CAPITALS = dwrite_make_opentype_tag("smcp")
    SIMPLIFIED_FORMS = dwrite_make_opentype_tag("smpl")
    STYLISTIC_SET_1 = dwrite_make_opentype_tag("ss01")
    STYLISTIC_SET_2 = dwrite_make_opentype_tag("ss02")
    STYLISTIC_SET_3 = dwrite_make_opentype_tag("ss03")
    STYLISTIC_SET_4 = dwrite_make_opentype_tag("ss04")
    STYLISTIC_SET_5 = dwrite_make_opentype_tag("ss05")
    STYLISTIC_SET_6 = dwrite_make_opentype_tag("ss06")
    STYLISTIC_SET_7 = dwrite_make_opentype_tag("ss07")
    STYLISTIC_SET_8 = dwrite_make_opentype_tag("ss08")
    STYLISTIC_SET_9 = dwrite_make_opentype_tag("ss09")
    STYLISTIC_SET_10 = dwrite_make_opentype_tag("ss10")
    STYLISTIC_SET_11 = dwrite_make_opentype_tag("ss11")
    STYLISTIC_SET_12 = dwrite_make_opentype_tag("ss12")
    STYLISTIC_SET_13 = dwrite_make_opentype_tag("ss13")
    STYLISTIC_SET_14 = dwrite_make_opentype_tag("ss14")
    STYLISTIC_SET_15 = dwrite_make_opentype_tag("ss15")
    STYLISTIC_SET_16 = dwrite_make_opentype_tag("ss16")
    STYLISTIC_SET_17 = dwrite_make_opentype_tag("ss17")
    STYLISTIC_SET_18 = dwrite_make_opentype_tag("ss18")
    STYLISTIC_SET_19 = dwrite_make_opentype_tag("ss19")
    STYLISTIC_SET_20 = dwrite_make_opentype_tag("ss20")
    SUBSCRIPT = dwrite_make_opentype_tag("subs")
    SUPERSCRIPT = dwrite_make_opentype_tag("sups")
    SWASH = dwrite_make_opentype_tag("swsh")
    TITLING = dwrite_make_opentype_tag("titl")
    TRADITIONAL_NAME_FORMS = dwrite_make_opentype_tag("tnam")
    TABULAR_FIGURES = dwrite_make_opentype_tag("tnum")
    TRADITIONAL_FORMS = dwrite_make_opentype_tag("trad")
    THIRD_WIDTHS = dwrite_make_opentype_tag("twid")
    UNICASE = dwrite_make_opentype_tag("unic")
    VERTICAL_WRITING = dwrite_make_opentype_tag("vert")
    VERTICAL_ALTERNATES_AND_ROTATION = dwrite_make_opentype_tag("vrt2")
    SLASHED_ZERO = dwrite_make_opentype_tag("zero")


class DWriteTextRange(Structure):
    """DWRITE_TEXT_RANGE"""

    __slots__ = ()
    _fields_ = (
        ("start_position", c_uint32),
        ("length", c_uint32),
    )


class DWriteFontFeature(Structure):
    """DWRITE_FONT_FEATURE"""

    __slots__ = ()
    _fields_ = (
        ("_name_tag", c_int32),  # DWRITE_FONT_FEATURE_TAG
        ("parameter", c_uint32),
    )

    @property
    def name_tag(self) -> DWriteFontFeatureTag:
        return DWriteFontFeatureTag(self._name_tag)

    @name_tag.setter
    def name_tag(self, value: DWriteFontFeatureTag) -> None:
        self._name_tag = value


class DWriteTypegraphicFeatures(Structure):
    """DWRITE_TYPOGRAPHIC_FEATURES"""

    __slots__ = ()
    _fields_ = (
        ("features", POINTER(DWriteFontFeature)),
        ("feature_count", c_uint32),
    )


class DWriteTrimming(Structure):
    """DWRITE_TRIMMING"""

    __slots__ = ()
    _fields_ = (
        ("_granularity", c_int32),
        ("delimiter", c_uint32),
        ("delimiter_count", c_uint32),
    )

    @property
    def granularity(self) -> DWriteTrimmingGranularity:
        return DWriteTrimmingGranularity(self._granularity)

    @granularity.setter
    def granularity(self, value: DWriteTrimmingGranularity) -> None:
        self._granularity = value


class DWriteScriptShape(IntEnum):
    """DWRITE_SCRIPT_SHAPES"""

    DEFAULT = 0
    NO_VISUAL = 1


class DWriteScriptAnalysis(Structure):
    """DWRITE_SCRIPT_ANALYSIS"""

    __slots__ = ()
    _fields_ = (
        ("script", c_uint16),
        ("_shapes", c_int32),
    )

    @property
    def shapes(self) -> DWriteScriptShape:
        return DWriteScriptShape(self._shapes)

    @shapes.setter
    def shapes(self, value: DWriteScriptShape) -> None:
        self._shapes = value


class DWriteBreakCondition(IntEnum):
    """DWRITE_BREAK_CONDITION"""

    NEUTRAL = 0
    CAN_BREAK = 1
    MAY_NOT_BREAK = 2
    MUST_BREAK = 3


class DWriteLineBreakpoint(Structure):
    """DWRITE_LINE_BREAKPOINT"""

    __slots__ = ()
    _fields_ = (("_x", c_uint8),)
    #     UINT8 breakConditionBefore  : 2;
    #     UINT8 breakConditionAfter   : 2;
    #     UINT8 isWhitespace          : 1;
    #     UINT8 isSoftHyphen          : 1;
    #     UINT8 padding               : 2;

    @property
    def break_condition_before(self) -> DWriteBreakCondition:
        return DWriteBreakCondition((self._x & 0b00000011) >> 0)

    @break_condition_before.setter
    def break_condition_before(self, value: DWriteBreakCondition) -> None:
        if not (0 <= int(value) <= 0b00000011):
            raise ValueError
        self._x = (self._x & ~0b00000011) | (int(value) << 0)

    @property
    def break_condition_after(self) -> DWriteBreakCondition:
        return DWriteBreakCondition((self._x & 0b00001100) >> 2)

    @break_condition_before.setter
    def break_condition_before(self, value: DWriteBreakCondition) -> None:
        if not (0 <= int(value) <= 0b00000011):
            raise ValueError
        self._x = (self._x & ~0b00001100) | (int(value) << 2)

    @property
    def is_whitespace(self) -> bool:
        return (self._x & 0b00010000) != 0

    @is_whitespace.setter
    def is_whitespace(self, value: bool) -> None:
        self._x = (self._x & ~0b00010000) | (0b00010000 if value << 4 else 0)

    @property
    def is_softhypen(self) -> bool:
        return (self._x & 0b00100000) != 0

    @is_softhypen.setter
    def is_softhypen(self, value: bool) -> None:
        self._x = (self._x & ~0b00100000) | (0b00100000 if value else 0)

    @property
    def padding(self) -> int:
        return (self._x & 0b11000000) >> 6

    @padding.setter
    def padding(self, value: int) -> None:
        if not (0 <= value <= 0b00000011):
            raise ValueError
        self._x = (self._x & ~0b11000000) | (int(value) << 6)


class DWriteNumberSubstitutionMethod(IntEnum):
    """DWRITE_NUMBER_SUBSTITUTION_METHOD"""

    FROM_CULTURE = 0
    CONTEXTUAL = 1
    NONE = 2
    NATIONAL = 3
    TRADITIONAL = 4


class DWriteShapingTextProps(Structure):
    """DWRITE_SHAPING_TEXT_PROPERTIES"""

    __slots__ = ()
    _fields_ = (("_x", c_uint16),)

    @property
    def is_shapedalone(self) -> bool:
        return (self._x & 0b00000000_00000001) != 0

    @is_shapedalone.setter
    def is_shapedalone(self, value: bool) -> None:
        self._x = (self._x & ~0b00000000_00000001) | (0b00000000_00000001 if value else 0)

    @property
    def reserved1(self) -> bool:
        return (self._x & 0b00000000_00000010) != 0

    @reserved1.setter
    def reserved1(self, value: bool) -> None:
        self._x = (self._x & ~0b00000000_00000010) | (0b00000000_00000010 if value else 0)

    @property
    def can_break_shapingafter(self) -> bool:
        return (self._x & 0b00000000_00000100) != 0

    @can_break_shapingafter.setter
    def can_break_shapingafter(self, value: bool) -> None:
        self._x = (self._x & ~0b00000000_00000100) | (0b00000000_00000100 if value else 0)

    @property
    def reserved(self) -> int:
        return (self._x & 0b11111111_11111000) >> 3

    @reserved.setter
    def reserved(self, value: int) -> None:
        if not (0 <= value <= 0b00011111_11111111):
            raise ValueError
        self._x = (self._x & ~0b11111111_11111000) | (int(value) << 3)


class DWriteShapingGlyphProps(Structure):
    """DWRITE_SHAPING_GLYPH_PROPERTIES"""

    __slots__ = ()
    _fields_ = (("_x", c_uint16),)
    #     UINT16 justification        : 4;
    #     UINT16 isClusterStart       : 1;
    #     UINT16 isDiacritic          : 1;
    #     UINT16 isZeroWidthSpace     : 1;
    #     UINT16 reserved             : 9;

    @property
    def justification(self) -> int:
        return self._x & 0b00000000_00001111

    @justification.setter
    def justification(self, value: int) -> None:
        if not (0 <= value <= 0b00000000_00001111):
            raise ValueError
        self._x = (self._x & ~0b00000000_00001111) | (int(value) << 0)

    @property
    def is_clusterstart(self) -> bool:
        return (self._x & 0b00000000_00010000) != 0

    @is_clusterstart.setter
    def is_clusterstart(self, value: bool) -> None:
        self._x = (self._x & ~0b00000000_00010000) | (0b00000000_00010000 if value else 0)

    @property
    def is_diacritic(self) -> bool:
        return (self._x & 0b00000000_00100000) != 0

    @is_diacritic.setter
    def is_diacritic(self, value: bool) -> None:
        self._x = (self._x & ~0b00000000_00100000) | (0b00000000_00100000 if value else 0)

    @property
    def is_zerowidthspace(self) -> bool:
        return (self._x & 0b00000000_01000000) != 0

    @is_zerowidthspace.setter
    def is_zerowidthspace(self, value: bool) -> None:
        self._x = (self._x & ~0b00000000_01000000) | (0b00000000_01000000 if value else 0)

    @property
    def reserved(self) -> int:
        return (self._x & 0b11111111_10000000) >> 7

    @reserved.setter
    def reserved(self, value: int) -> None:
        if not (0 <= value <= 0b00000001_11111111):
            raise ValueError
        self._x = (self._x & ~0b11111111_10000000) | (int(value) << 7)


class DWriteGlyphRun(Structure):
    """DWRITE_GLYPH_RUN"""

    __slots__ = ()


class DWriteGlyphRunDesc(Structure):
    """DWRITE_GLYPH_RUN_DESCRIPTION"""

    __slots__ = ()
    _fields_ = (
        ("locale_name", c_wchar_p),
        ("string", c_wchar_p),
        ("string_lengthlocale_name", c_uint32),
        ("cluster_map", POINTER(c_uint16)),
        ("text_position", c_uint32),
    )


class DWriteUnderline(Structure):
    """DWRITE_UNDERLINE"""

    __slots__ = ()
    _fields_ = (
        ("width", c_float),
        ("thickness", c_float),
        ("offset", c_float),
        ("run_height", c_float),
        ("reading_direction", c_int32),
        ("flow_direction", c_int32),
        ("locale_name", c_void_p),
        ("measuring_mode", c_int32),
    )


class DWriteStrikethrough(Structure):
    """DWRITE_STRIKETHROUGH"""

    __slots__ = ()
    _fields_ = (
        ("width", c_float),
        ("thickness", c_float),
        ("offset", c_float),
        ("run_height", c_float),
        ("reading_direction", c_int32),
        ("flow_direction", c_int32),
        ("locale_name", c_void_p),
        ("measuring_mode", c_int32),
    )


class DWriteLineMetrics(Structure):
    """DWRITE_LINE_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("length", c_uint32),
        ("trailing_whitespace_length", c_uint32),
        ("newline_length", c_uint32),
        ("height", c_float),
        ("baseline", c_float),
        ("is_trimmed", c_int32),
    )


class DWriteClusterMetrics(Structure):
    """DWRITE_CLUSTER_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("width", c_float),
        ("length", c_uint16),
        ("_x", c_uint16),
        # UINT16 canWrapLineAfter : 1;
        # UINT16 isWhitespace : 1;
        # UINT16 isNewline : 1;
        # UINT16 isSoftHyphen : 1;
        # UINT16 isRightToLeft : 1;
        # UINT16 padding : 11;
    )

    @property
    def can_wrap_line_after(self) -> bool:
        return (self.x & 0b00000000_00000001) != 0

    @can_wrap_line_after.setter
    def can_wrap_line_after(self, value: bool) -> None:
        self.x = (self.x & ~0b00000000_00000001) | (0b00000000_00000001 if value else 0)

    @property
    def is_whitespace(self) -> bool:
        return (self.x & 0b00000000_00000010) != 0

    @is_whitespace.setter
    def is_whitespace(self, value: bool) -> None:
        self.x = (self.x & ~0b00000000_00000010) | (0b00000000_00000010 if value else 0)

    @property
    def is_newline(self) -> bool:
        return (self.x & 0b00000000_00000100) != 0

    @is_newline.setter
    def is_newline(self, value: bool) -> None:
        self.x = (self.x & ~0b00000000_00000100) | (0b00000000_00000100 if value else 0)

    @property
    def is_soft_hyphen(self) -> bool:
        return (self.x & 0b00000000_00001000) != 0

    @is_soft_hyphen.setter
    def is_soft_hyphen(self, value: bool) -> None:
        self.x = (self.x & ~0b00000000_00001000) | (0b00000000_00001000 if value else 0)

    @property
    def is_rtl(self) -> bool:
        return (self.x & 0b00000000_00010000) != 0

    @is_rtl.setter
    def is_rtl(self, value: bool) -> None:
        self.x = (self.x & ~0b00000000_00010000) | (0b00000000_00010000 if value else 0)

    @property
    def padding(self) -> int:
        return (self.x & 0b11111111_11100000) != 0

    @padding.setter
    def padding(self, value: int) -> None:
        if not (0 <= value <= 0b00000111_11111111):
            raise ValueError
        self.x = (self.x & ~0b11111111_11100000) | (value << 13)


class DWriteTextMetrics(Structure):
    """DWRITE_TEXT_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("left", c_float),
        ("top", c_float),
        ("width", c_float),
        ("width_including_trailing_whitespace", c_float),
        ("height", c_float),
        ("layout_width", c_float),
        ("layout_height", c_float),
        ("max_bidi_reordering_deptch", c_uint32),
        ("line_count", c_uint32),
    )


class DWriteInlineObjectMetrics(Structure):
    """DWRITE_INLINE_OBJECT_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("width", c_float),
        ("baseline", c_float),
        ("supports_sideways", c_int32),
    )


class DWriteOverhangMetrics(Structure):
    """DWRITE_OVERHANG_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("left", c_float),
        ("top", c_float),
        ("right", c_float),
        ("bottom", c_float),
    )


class DWriteHitTestMetrics(Structure):
    """DWRITE_HIT_TEST_METRICS"""

    __slots__ = ()
    _fields_ = (
        ("text_position", c_uint32),
        ("length", c_uint32),
        ("left", c_float),
        ("top", c_float),
        ("width", c_float),
        ("height", c_float),
        ("bidi_level", c_uint32),
        ("is_text", c_int32),
        ("is_trimmed", c_int32),
    )


class DWriteTextureType(IntEnum):
    """DWRITE_TEXTURE_TYPE"""

    ALIASED_1x1 = 0
    CLEARTYPE_3x1 = 1


DWRITE_ALPHA_MAX = 255


FACILITY_DWRITE = 0x898
DWRITE_ERR_BASE = 0x5000

# #define MAKE_DWRITE_HR(severity, code) MAKE_HRESULT(severity, FACILITY_DWRITE, (DWRITE_ERR_BASE + code))
# #define MAKE_DWRITE_HR_ERR(code) MAKE_DWRITE_HR(SEVERITY_ERROR, code)


class IDWriteFontFileLoader(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{727cad4e-d6af-4c9e-8a08-d695b11caa49}")


class IDWriteLocalFontFileLoader(IUnknown):  # IDWriteFontFileLoader
    __slots__ = ()
    _iid_ = GUID("{b2d9f3ec-c9fe-4a11-a2ec-d86208f7c0a2}")


class IDWriteFontFileStream(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{6d4865fe-0ab8-4d91-8f62-5dd6be34a3e0}")


class IDWriteRenderingParams(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{2f0da53a-2add-47cd-82ee-d9ec34688e75}")


class IDWriteFontFace(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{5f49804d-7024-4d43-bfa9-d25984f53849}")


class IDWriteFontCollectionLoader(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{cca920e4-52f0-492b-bfa8-29c72ee0a468}")


class IDWriteFontFileEnumerator(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{72755049-5ff7-435d-8348-4be97cfa6c7c}")


class IDWriteLocalizedStrings(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{08256209-099a-4b34-b86d-c22b110e7771}")


class IDWriteFontCollection(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{a84cee02-3eea-4eee-a827-87c1a02a0fcc}")


class IDWriteFontList(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{1a0d8438-1d97-4ec1-aef9-a2fb86ed6acb}")


class IDWriteFontFamily(IUnknown):  # IDWriteFontList
    __slots__ = ()
    _iid_ = GUID("{da20d8ef-812a-4c43-9802-62ec4abd7add}")


class IDWriteFont(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{acd16696-8c14-4f5d-877e-fe3fc1d32737}")


class IDWriteFontFile(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{739d886a-cef5-47dc-8769-1a8b41bebbb0}")


class IDWriteTextFormat(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{9c906818-31d7-4fd3-a151-7c5e225db55a}")


class IDWriteTypography(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{55f1112b-1dc2-4b3c-9541-f46894ed85b6}")


class IDWriteNumberSubstitution(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{14885CC9-BAB0-4f90-B6ED-5C366A2CD03D}")


class IDWriteTextAnalysisSource(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{688e1a58-5094-47c8-adc8-fbcea60ae92b}")


class IDWriteTextAnalysisSink(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{5810cd44-0ca0-4701-b3fa-bec5182ae4f6}")


class IDWriteTextAnalyzer(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{b7e6163e-7f46-43b4-84b3-e4e6249c365d}")


class IDWriteInlineObject(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{8339FDE3-106F-47ab-8373-1C6295EB10B3}")


class IDWritePixelSnapping(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{eaf3a2da-ecf4-4d24-b644-b34f6842024b}")


class IDWriteTextRenderer(IUnknown):  # IDWritePixelSnapping
    __slots__ = ()
    _iid_ = GUID("{ef8a8135-5cc6-45fe-8825-c5a0724eb819}")


class IDWriteTextLayout(IUnknown):  # IDWriteTextFormat
    __slots__ = ()
    _iid_ = GUID("{53737037-6d14-410b-9bfe-0b182bb70961}")


class IDWriteBitmapRenderTarget(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{5e5a32a3-8dff-4773-9ff6-0696eab77267}")


# TODO: Methods
#     STDMETHOD(DrawGlyphRun)(
#         FLOAT baselineOriginX,
#         FLOAT baselineOriginY,
#         DWRITE_MEASURING_MODE measuringMode,
#         _In_ DWRITE_GLYPH_RUN const* glyphRun,
#         _In_ IDWriteRenderingParams* renderingParams,
#         COLORREF textColor,
#         _Out_opt_ RECT* blackBoxRect = NULL
#         ) PURE;
#     STDMETHOD_(HDC, GetMemoryDC)() PURE;
#     STDMETHOD_(FLOAT, GetPixelsPerDip)() PURE;
#     STDMETHOD(SetPixelsPerDip)(
#         FLOAT pixelsPerDip
#         ) PURE;
#     STDMETHOD(GetCurrentTransform)(
#         _Out_ DWRITE_MATRIX* transform
#         ) PURE;
#     STDMETHOD(SetCurrentTransform)(
#         _In_opt_ DWRITE_MATRIX const* transform
#         ) PURE;
#     STDMETHOD(GetSize)(
#         _Out_ SIZE* size
#         ) PURE;
#     STDMETHOD(Resize)(
#         UINT32 width,
#         UINT32 height
#         ) PURE;
# };


class IDWriteGdiInterop(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{1edd9491-9853-4299-898f-6432983b6f3a}")


#     STDMETHOD(CreateFontFromLOGFONT)(
#         _In_ LOGFONTW const* logFont,
#         _COM_Outptr_ IDWriteFont** font
#         ) PURE;
#     STDMETHOD(ConvertFontToLOGFONT)(
#         _In_ IDWriteFont* font,
#         _Out_ LOGFONTW* logFont,
#         _Out_ BOOL* isSystemFont
#         ) PURE;
#     STDMETHOD(ConvertFontFaceToLOGFONT)(
#         _In_ IDWriteFontFace* font,
#         _Out_ LOGFONTW* logFont
#         ) PURE;
#     STDMETHOD(CreateFontFaceFromHdc)(
#         HDC hdc,
#         _COM_Outptr_ IDWriteFontFace** fontFace
#         ) PURE;
#     STDMETHOD(CreateBitmapRenderTarget)(
#         _In_opt_ HDC hdc,
#         UINT32 width,
#         UINT32 height,
#         _COM_Outptr_ IDWriteBitmapRenderTarget** renderTarget
#         ) PURE;
# };


class IDWriteGlyphRunAnalysis(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{7d97dbf7-e085-42d4-81e3-6a883bded118}")


#     STDMETHOD(GetAlphaTextureBounds)(
#         DWRITE_TEXTURE_TYPE textureType,
#         _Out_ RECT* textureBounds
#         ) PURE;

#     STDMETHOD(CreateAlphaTexture)(
#         DWRITE_TEXTURE_TYPE textureType,
#         _In_ RECT const* textureBounds,
#         _Out_writes_bytes_(bufferSize) BYTE* alphaValues,
#         UINT32 bufferSize
#         ) PURE;

#     STDMETHOD(GetAlphaBlendParams)(
#         _In_ IDWriteRenderingParams* renderingParams,
#         _Out_ FLOAT* blendGamma,
#         _Out_ FLOAT* blendEnhancedContrast,
#         _Out_ FLOAT* blendClearTypeLevel
#         ) PURE;
# };


class IDWriteFactory(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{b859ee5a-d838-4b5b-a2e8-1adc7d93db48}")
