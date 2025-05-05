from comtypes import STDMETHOD, c_uint64
from powc.datetime import FILETIME

from ..d2d.types import IDWriteGeometrySink
from .types import *

DWriteGlyphRun._fields_ = (
    ("font_face", POINTER(IDWriteFontFace)),
    ("font_em_size", c_float),
    ("glyph_count", c_uint32),
    ("glyph_indices", POINTER(c_uint16)),
    ("glyph_advances", POINTER(c_float)),
    ("glyph_offsets", POINTER(DWriteGlyphOffset)),
    ("is_sideways", c_int32),
    ("bidi_level", c_uint32),
)

IDWriteFontFileLoader._methods_ = [
    STDMETHOD(c_int32, "CreateStreamFromKey", (c_void_p, c_uint32, POINTER(POINTER(IDWriteFontFileStream)))),
]

IDWriteLocalFontFileLoader._methods_ = [
    *IDWriteFontFileLoader._methods_,
    STDMETHOD(c_int32, "GetFilePathLengthFromKey", (c_void_p, c_uint32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetFilePathFromKey", (c_void_p, c_uint32, c_wchar_p, c_uint32)),
    STDMETHOD(c_int32, "GetLastWriteTimeFromKey", (c_void_p, c_uint32, POINTER(FILETIME))),
]
IDWriteFontFileStream._methods_ = [
    STDMETHOD(c_int32, "ReadFileFragment", (POINTER(c_void_p), c_uint64, c_uint64, POINTER(c_void_p))),
    STDMETHOD(c_int32, "ReleaseFileFragment", (c_void_p,)),
    STDMETHOD(c_int32, "GetFileSize", (POINTER(c_uint64),)),
    STDMETHOD(c_int32, "GetLastWriteTime", (POINTER(c_uint64),)),
]
IDWriteRenderingParams._methods_ = [
    STDMETHOD(c_float, "GetGamma", ()),
    STDMETHOD(c_float, "GetEnhancedContrast", ()),
    STDMETHOD(c_float, "GetClearTypeLevel", ()),
    STDMETHOD(c_int32, "GetPixelGeometry", ()),  # DWRITE_PIXEL_GEOMETRY
    STDMETHOD(c_int32, "GetRenderingMode", ()),  # DWRITE_RENDERING_MODE
]
IDWriteFontFace._methods_ = [
    STDMETHOD(c_int32, "GetType", ()),  # DWRITE_FONT_FACE_TYPE
    STDMETHOD(c_int32, "GetFiles", (POINTER(c_uint32), POINTER(POINTER(IDWriteFontFile)))),
    STDMETHOD(c_uint32, "GetIndex", ()),
    STDMETHOD(c_int32, "GetSimulations", ()),  # DWRITE_FONT_SIMULATIONS
    STDMETHOD(c_int32, "IsSymbolFont", ()),
    STDMETHOD(None, "GetMetrics", (POINTER(DWriteFontMetrics),)),
    STDMETHOD(c_uint16, "GetGlyphCount", ()),
    STDMETHOD(c_int32, "GetDesignGlyphMetrics", (POINTER(c_uint16), c_uint32, POINTER(DWriteGlyphMetrics), c_int32)),
    STDMETHOD(c_int32, "GetGlyphIndices", (POINTER(c_uint32), c_uint32, POINTER(c_uint16))),
    STDMETHOD(
        c_int32,
        "TryGetFontTable",
        (c_uint32, POINTER(c_void_p), POINTER(c_uint32), POINTER(c_void_p), POINTER(c_int32)),
    ),
    STDMETHOD(None, "ReleaseFontTable", (c_void_p,)),
    STDMETHOD(
        c_int32,
        "GetGlyphRunOutline",
        (
            c_float,
            POINTER(c_float),
            POINTER(DWriteGlyphOffset),
            c_uint32,
            c_int32,
            c_int32,
            POINTER(IDWriteGeometrySink),
        ),
    ),
    STDMETHOD(
        c_int32,
        "GetRecommendedRenderingMode",
        (c_float, c_float, c_int32, POINTER(IDWriteRenderingParams), POINTER(c_int32)),
    ),
    STDMETHOD(
        c_int32, "GetGdiCompatibleMetrics", (c_float, c_float, POINTER(DWriteMatrix), POINTER(DWriteFontMetrics))
    ),
    STDMETHOD(
        c_int32,
        "GetGdiCompatibleGlyphMetrics",
        (
            c_float,
            c_float,
            POINTER(DWriteMatrix),
            c_int32,
            POINTER(c_uint16),
            c_uint32,
            POINTER(DWriteGlyphMetrics),
            c_int32,
        ),
    ),
]
IDWriteFontCollectionLoader._methods_ = [
    STDMETHOD(
        c_int32,
        "CreateEnumeratorFromKey",
        (POINTER(IDWriteFactory), c_void_p, c_int32, POINTER(POINTER(IDWriteFontFileEnumerator))),
    ),
]
IDWriteFontFileEnumerator._methods_ = [
    STDMETHOD(c_int32, "MoveNext", (POINTER(c_int32),)),
    STDMETHOD(c_int32, "GetCurrentFontFile", (POINTER(POINTER(IDWriteFontFile)),)),
]
IDWriteLocalizedStrings._methods_ = [
    STDMETHOD(c_uint32, "GetCount", ()),
    STDMETHOD(c_int32, "FindLocaleName", (c_wchar_p, POINTER(c_uint32), POINTER(c_int32))),
    STDMETHOD(c_int32, "GetLocaleNameLength", (c_uint32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetLocaleName", (c_uint32, c_wchar_p, c_uint32)),
    STDMETHOD(c_int32, "GetStringLength", (c_uint32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetString", (c_uint32, c_wchar_p, c_uint32)),
]
IDWriteFontCollection._methods_ = [
    STDMETHOD(c_uint32, "GetFontFamilyCount", ()),
    STDMETHOD(c_int32, "GetFontFamily", (c_uint32, POINTER(POINTER(IDWriteFontFamily)))),
    STDMETHOD(c_int32, "FindFamilyName", (c_wchar_p, POINTER(c_uint32), POINTER(c_int32))),
    STDMETHOD(c_int32, "GetFontFromFontFace", (POINTER(IDWriteFontFace), POINTER(POINTER(IDWriteFont)))),
]
IDWriteFontList._methods_ = [
    STDMETHOD(
        c_int32,
        "GetFontCollection",
        (POINTER(POINTER(IDWriteFontCollection)),),
    ),
    STDMETHOD(c_uint32, "GetFontCount", ()),
    STDMETHOD(c_int32, "GetFont", (c_uint32, POINTER(POINTER(IDWriteFont)))),
]
IDWriteFontFamily._methods_ = [
    *IDWriteFontList._methods_,
    STDMETHOD(
        c_int32,
        "GetFamilyNames",
        (POINTER(POINTER(IDWriteLocalizedStrings)),),
    ),
    STDMETHOD(c_int32, "GetFirstMatchingFont", (c_int32, c_int32, c_int32, POINTER(POINTER(IDWriteFont)))),
]
IDWriteFont._methods_ = [
    STDMETHOD(c_int32, "GetFontFamily", (POINTER(POINTER(IDWriteFontFamily)),)),
    STDMETHOD(c_int32, "GetWeight", ()),
    STDMETHOD(c_int32, "GetStretch", ()),
    STDMETHOD(c_int32, "GetStyle", ()),
    STDMETHOD(c_int32, "IsSymbolFont", ()),
    STDMETHOD(c_int32, "GetFaceNames", (POINTER(POINTER(IDWriteLocalizedStrings)),)),
    STDMETHOD(
        c_int32, "GetInformationalStrings", (c_int32, POINTER(POINTER(IDWriteLocalizedStrings)), POINTER(c_int32))
    ),
    STDMETHOD(c_int32, "GetSimulations", ()),
    STDMETHOD(c_int32, "GetMetrics", (POINTER(DWriteFontMetrics),)),
    STDMETHOD(c_int32, "HasCharacter", (c_uint32, POINTER(c_int32))),
    STDMETHOD(c_int32, "CreateFontFace", (POINTER(POINTER(IDWriteFontFace)),)),
]
IDWriteFontFile._methods_ = [
    STDMETHOD(c_int32, "GetReferenceKey", (POINTER(c_void_p), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetLoader", (POINTER(POINTER(IDWriteFontFileLoader)),)),
    STDMETHOD(c_int32, "Analyze", (POINTER(c_int32), POINTER(c_int32), POINTER(c_int32), POINTER(c_uint32))),
]
IDWriteTextFormat._methods_ = [
    STDMETHOD(c_int32, "SetTextAlignment", (c_int32,)),
    STDMETHOD(c_int32, "SetParagraphAlignment", (c_int32,)),
    STDMETHOD(c_int32, "SetWordWrapping", (c_int32,)),
    STDMETHOD(c_int32, "SetReadingDirection", (c_int32,)),
    STDMETHOD(c_int32, "SetFlowDirection", (c_int32,)),
    STDMETHOD(c_int32, "SetIncrementalTabStop", (c_float,)),
    STDMETHOD(c_int32, "SetTrimming", (POINTER(DWriteTrimming), POINTER(IDWriteInlineObject))),
    STDMETHOD(c_int32, "SetLineSpacing", (c_int32, c_float, c_float)),
    STDMETHOD(c_int32, "GetTextAlignment", ()),
    STDMETHOD(c_int32, "GetParagraphAlignment", ()),
    STDMETHOD(c_int32, "GetWordWrapping", ()),
    STDMETHOD(c_int32, "GetReadingDirection", ()),
    STDMETHOD(c_int32, "GetFlowDirection", ()),
    STDMETHOD(c_float, "GetIncrementalTabStop", ()),
    STDMETHOD(c_int32, "GetTrimming", (POINTER(c_int32), POINTER(POINTER(IDWriteInlineObject)))),
    STDMETHOD(c_int32, "GetLineSpacing", (POINTER(c_int32), POINTER(c_float), POINTER(c_float))),
    STDMETHOD(c_int32, "GetFontCollection", (POINTER(POINTER(IDWriteFontCollection)),)),
    STDMETHOD(c_uint32, "GetFontFamilyNameLength", ()),
    STDMETHOD(c_int32, "GetFontFamilyName", (c_wchar_p, c_uint32)),
    STDMETHOD(c_int32, "GetFontWeight", ()),
    STDMETHOD(c_int32, "GetFontStyle", ()),
    STDMETHOD(c_int32, "GetFontStretch", ()),
    STDMETHOD(c_float, "GetFontSize", ()),
    STDMETHOD(c_uint32, "GetLocaleNameLength", ()),
    STDMETHOD(c_int32, "GetLocaleName", (c_wchar_p, c_uint32)),
]
IDWriteTypography._methods_ = [
    STDMETHOD(c_int32, "AddFontFeature", (DWriteFontFeature,)),
    STDMETHOD(c_uint32, "GetFontFeatureCount", ()),
    STDMETHOD(c_int32, "GetFontFeature", (c_uint32, POINTER(DWriteFontFeature))),
]
IDWriteTextAnalysisSource._methods_ = [
    STDMETHOD(c_int32, "GetTextAtPosition", (c_uint32, POINTER(c_wchar_p), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetTextBeforePosition", (c_uint32, POINTER(c_wchar_p), POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetParagraphReadingDirection", ()),
    STDMETHOD(c_int32, "GetLocaleName", (c_uint32, POINTER(c_uint32), POINTER(c_wchar_p))),
    STDMETHOD(
        c_int32, "GetNumberSubstitution", (c_uint32, POINTER(c_uint32), POINTER(POINTER(IDWriteNumberSubstitution)))
    ),
]
IDWriteTextAnalysisSink._methods_ = [
    STDMETHOD(c_int32, "SetScriptAnalysis", (c_uint32, c_uint32, POINTER(DWriteScriptAnalysis))),
    STDMETHOD(c_int32, "SetLineBreakpoints", (c_uint32, c_uint32, POINTER(DWriteLineBreakpoint))),
    STDMETHOD(c_int32, "SetBidiLevel", (c_uint32, c_uint32, c_uint8, c_uint8)),
    STDMETHOD(c_int32, "SetNumberSubstitution", (c_uint32, c_uint32, POINTER(IDWriteNumberSubstitution))),
    STDMETHOD(c_int32, "", ()),
    STDMETHOD(c_int32, "", ()),
]
IDWriteTextAnalyzer._methods_ = [
    STDMETHOD(
        c_int32,
        "AnalyzeScript",
        (POINTER(IDWriteTextAnalysisSource), c_uint32, c_uint32, POINTER(IDWriteTextAnalysisSink)),
    ),
    STDMETHOD(
        c_int32,
        "AnalyzeBidi",
        (POINTER(IDWriteTextAnalysisSource), c_uint32, c_uint32, POINTER(IDWriteTextAnalysisSink)),
    ),
    STDMETHOD(
        c_int32,
        "AnalyzeNumberSubstition",
        (POINTER(IDWriteTextAnalysisSource), c_uint32, c_uint32, POINTER(IDWriteTextAnalysisSink)),
    ),
    STDMETHOD(
        c_int32,
        "AnalyzeLineBreakpoints",
        (POINTER(IDWriteTextAnalysisSource), c_uint32, c_uint32, POINTER(IDWriteTextAnalysisSink)),
    ),
    STDMETHOD(
        c_int32,
        "GetGlyphs",
        (
            c_wchar_p,
            c_uint32,
            POINTER(IDWriteFontFace),
            c_int32,
            c_int32,
            POINTER(DWriteScriptAnalysis),
            c_wchar_p,
            POINTER(IDWriteNumberSubstitution),
            POINTER(DWriteTypegraphicFeatures),
            POINTER(c_uint32),
            c_uint32,
            c_uint32,
            POINTER(c_uint16),
            POINTER(DWriteShapingTextProps),
            POINTER(c_uint16),
            POINTER(DWriteShapingGlyphProps),
            POINTER(c_uint32),
        ),
    ),
    STDMETHOD(
        c_int32,
        "GetGlyphPlacements",
        (
            c_wchar_p,
            POINTER(c_uint16),
            POINTER(DWriteShapingTextProps),
            c_uint32,
            POINTER(c_uint16),
            POINTER(DWriteShapingGlyphProps),
            c_uint32,
            POINTER(IDWriteFontFace),
            c_float,
            c_int32,
            c_int32,
            POINTER(DWriteScriptAnalysis),
            c_wchar_p,
            POINTER(DWriteTypegraphicFeatures),
            POINTER(c_uint32),
            c_uint32,
            POINTER(c_float),
            POINTER(DWriteGlyphOffset),
        ),
    ),
    STDMETHOD(
        c_int32,
        "GetGdiCompatibleGlyphPlacements",
        (
            c_wchar_p,
            POINTER(c_uint16),
            POINTER(DWriteShapingTextProps),
            c_uint32,
            POINTER(c_uint16),
            POINTER(DWriteShapingGlyphProps),
            c_uint32,
            POINTER(IDWriteFontFace),
            c_float,
            c_float,
            POINTER(DWriteMatrix),
            c_int32,
            c_int32,
            c_int32,
            POINTER(DWriteScriptAnalysis),
            c_wchar_p,
            POINTER(POINTER(DWriteTypegraphicFeatures)),
            POINTER(c_uint32),
            c_uint32,
            POINTER(c_float),
            POINTER(DWriteGlyphOffset),
        ),
    ),
]
IDWriteInlineObject._methods_ = [
    STDMETHOD(
        c_int32, "Draw", (c_void_p, POINTER(IDWriteTextRenderer), c_float, c_float, c_int32, c_int32, POINTER(IUnknown))
    ),
    STDMETHOD(c_int32, "GetMetrics", (POINTER(DWriteInlineObjectMetrics),)),
    STDMETHOD(c_int32, "GetOverhangMetrics", (POINTER(DWriteOverhangMetrics),)),
    STDMETHOD(c_int32, "GetBreakConditions", (POINTER(c_int32), POINTER(c_int32))),
]
IDWritePixelSnapping._methods_ = [
    STDMETHOD(c_int32, "IsPixelSnappingDisabled", (c_void_p, POINTER(c_int32))),
    STDMETHOD(c_int32, "GetcurrentTransform", (c_void_p, POINTER(DWriteMatrix))),
    STDMETHOD(c_int32, "GetPixelsPerDip", (c_void_p, POINTER(c_float))),
]
IDWriteTextRenderer._methods_ = [
    *IDWritePixelSnapping._methods_,
    STDMETHOD(
        c_int32,
        "DrawGlyphRun",
        (
            c_void_p,
            c_float,
            c_float,
            c_int32,
            POINTER(DWriteGlyphRun),
            POINTER(DWriteGlyphRunDesc),
            POINTER(IUnknown),
        ),
    ),
    STDMETHOD(c_int32, "DrawUnderline", (c_void_p, c_float, c_float, POINTER(DWriteUnderline), POINTER(IUnknown))),
    STDMETHOD(
        c_int32, "DrawStrikethrough", (c_void_p, c_float, c_float, POINTER(DWriteStrikethrough), POINTER(IUnknown))
    ),
    STDMETHOD(
        c_int32,
        "DrawInlineObject",
        (c_void_p, c_float, c_float, POINTER(IDWriteInlineObject), c_int32, c_int32, POINTER(IUnknown)),
    ),
]
IDWriteTextLayout._methods_ = [
    *IDWriteTextFormat._methods_,
    STDMETHOD(c_int32, "SetMaxWidth", (c_float,)),
    STDMETHOD(c_int32, "SetMaxHeight", (c_float,)),
    STDMETHOD(c_int32, "SetFontCollection", (POINTER(IDWriteFontCollection), DWriteTextRange)),
    STDMETHOD(c_int32, "SetFontFamilyName", (c_wchar_p, DWriteTextRange)),
    STDMETHOD(c_int32, "SetFontWeight", (c_int32, DWriteTextRange)),
    STDMETHOD(c_int32, "SetFontStyle", (c_int32, DWriteTextRange)),
    STDMETHOD(c_int32, "SetFontStretch", (c_int32, DWriteTextRange)),
    STDMETHOD(c_int32, "SetFontSize", (c_float, DWriteTextRange)),
    STDMETHOD(c_int32, "SetUnderline", (c_int32, DWriteTextRange)),
    STDMETHOD(c_int32, "SetStrilethrough", (c_int32, DWriteTextRange)),
    STDMETHOD(c_int32, "SetDrawingEffect", (POINTER(IUnknown), DWriteTextRange)),
    STDMETHOD(c_int32, "SetInlineObject", (POINTER(IDWriteInlineObject), DWriteTextRange)),
    STDMETHOD(c_int32, "SetTypography", (POINTER(IDWriteTypography), DWriteTextRange)),
    STDMETHOD(c_int32, "SetLocaleName", (c_wchar_p, DWriteTextRange)),
    STDMETHOD(c_float, "GetMaxWidth", ()),
    STDMETHOD(c_float, "GetMaxHeight", ()),
    STDMETHOD(c_int32, "GetFontCollection", (c_uint32, POINTER(IDWriteFontCollection), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetFontFamilyNameLength", (c_uint32, POINTER(c_uint32), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetFontFamilyName", (c_uint32, c_void_p, c_uint32, POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetFontWeight", (c_uint32, POINTER(c_uint32), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetFontStyle", (c_uint32, POINTER(c_int32), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetFontStretch", (c_uint32, POINTER(c_int32), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetFontSize", (c_uint32, POINTER(c_float), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetUnderline", (c_uint32, POINTER(c_int32), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetStrikethrough", (c_uint32, POINTER(c_int32), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetDeawingEffect", (c_uint32, POINTER(POINTER(IUnknown)), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetInlineObject", (c_uint32, POINTER(POINTER(IDWriteInlineObject)), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetTypography", (c_uint32, POINTER(POINTER(IDWriteTypography)), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetLocaleNameLength", (c_uint32, POINTER(POINTER(c_uint32)), POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "GetLocaleName", (c_void_p, c_uint32, POINTER(DWriteTextRange))),
    STDMETHOD(c_int32, "Draw", (c_void_p, POINTER(IDWriteTextRenderer), c_float, c_float)),
    STDMETHOD(c_int32, "GetLineMetrics", (POINTER(DWriteLineMetrics), c_uint32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "GetMetrics", (POINTER(DWriteTextMetrics),)),
    STDMETHOD(c_int32, "GetOverhangMetrics", (POINTER(DWriteOverhangMetrics),)),
    STDMETHOD(c_int32, "GetClusterMetrics", (POINTER(DWriteClusterMetrics), c_uint32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "DetermineMinWidth", (POINTER(c_float),)),
    STDMETHOD(
        c_int32,
        "HitTestPoint",
        (c_float, c_float, POINTER(c_int32), POINTER(c_int32), POINTER(DWriteHitTestMetrics)),
    ),
    STDMETHOD(
        c_int32,
        "HitTestTextPosition",
        (c_uint32, c_int32, POINTER(c_float), POINTER(c_float), POINTER(DWriteHitTestMetrics)),
    ),
    STDMETHOD(
        c_int32,
        "HitTestTextRange",
        (c_uint32, c_uint32, c_float, c_float, POINTER(DWriteHitTestMetrics), c_uint32, POINTER(c_uint32)),
    ),
]
IDWriteFactory._methods_ = [
    STDMETHOD(c_int32, "GetSystemFontCollection", (POINTER(POINTER(IDWriteFontCollection)), c_int32)),
    STDMETHOD(
        c_int32,
        "CreateCustomFontCollection",
        (POINTER(IDWriteFontCollectionLoader), c_void_p, c_uint32, POINTER(POINTER(IDWriteFontCollection))),
    ),
    STDMETHOD(c_int32, "RegisterFontCollectionLoader", (POINTER(IDWriteFontCollectionLoader),)),
    STDMETHOD(c_int32, "UnregisterFontCollectionLoader", (POINTER(IDWriteFontCollectionLoader),)),
    STDMETHOD(c_int32, "CreateFontFileReference", (c_wchar_p, POINTER(FILETIME), POINTER(POINTER(IDWriteFontFile)))),
    STDMETHOD(
        c_int32,
        "CreateCustomFontFileReference",
        (c_void_p, c_uint32, POINTER(IDWriteFontFileLoader), POINTER(POINTER(IDWriteFontFile))),
    ),
    STDMETHOD(
        c_int32,
        "CreateFontFace",
        (
            c_int32,
            c_uint32,
            POINTER(POINTER(IDWriteFontFile)),
            c_uint32,
            c_int32,
            POINTER(POINTER(IDWriteFontFace)),
        ),
    ),
    STDMETHOD(c_int32, "CreateRenderingParams", (POINTER(POINTER(IDWriteRenderingParams)),)),
    STDMETHOD(
        c_int32,
        "CreateMonitorRenderingParams",
        (
            c_void_p,
            POINTER(POINTER(IDWriteRenderingParams)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "CreateCustomRenderingParams",
        (c_float, c_float, c_float, c_int32, c_int32, POINTER(POINTER(IDWriteRenderingParams))),
    ),
    STDMETHOD(c_int32, "RegisterFontFileLoader", (POINTER(IDWriteFontFileLoader),)),
    STDMETHOD(c_int32, "UnRegisterFontFileLoader", (POINTER(IDWriteFontFileLoader),)),
    STDMETHOD(
        c_int32,
        "CreateTextFormat",
        (
            c_wchar_p,
            POINTER(IDWriteFontCollection),
            c_int32,
            c_int32,
            c_int32,
            c_float,
            c_wchar_p,
            POINTER(POINTER(IDWriteTextFormat)),
        ),
    ),
    STDMETHOD(c_int32, "CreateTypography", (POINTER(POINTER(IDWriteTypography)),)),
    STDMETHOD(c_int32, "GetGdiInterop", (POINTER(POINTER(IDWriteGdiInterop)),)),
    STDMETHOD(
        c_int32,
        "CreateTextLayout",
        (c_wchar_p, c_uint32, POINTER(IDWriteTextFormat), c_float, c_float, POINTER(IDWriteTextLayout)),
    ),
    STDMETHOD(
        c_int32,
        "CreateGdiCompatibleTextLayout",
        (
            c_wchar_p,
            c_uint32,
            POINTER(IDWriteTextFormat),
            c_float,
            c_float,
            c_float,
            POINTER(DWriteMatrix),
            c_int32,
            POINTER(POINTER(IDWriteTextLayout)),
        ),
    ),
    STDMETHOD(
        c_int32, "CreateEllipsisTrimmingSign", (POINTER(IDWriteTextFormat), POINTER(POINTER(IDWriteInlineObject)))
    ),
    STDMETHOD(c_int32, "CreateTextAnalyzer", (POINTER(POINTER(IDWriteTextAnalyzer)),)),
    STDMETHOD(
        c_int32,
        "CreateNumberSubstitution",
        (c_int32, c_wchar_p, c_int32, POINTER(POINTER(IDWriteNumberSubstitution))),
    ),
    STDMETHOD(
        c_int32,
        "CreateGlyphRunAnalysis",
        (
            POINTER(c_int32),
            c_float,
            POINTER(DWriteMatrix),
            c_int32,
            c_int32,
            c_float,
            c_float,
            POINTER(POINTER(IDWriteGlyphRunAnalysis)),
        ),
    ),
]
