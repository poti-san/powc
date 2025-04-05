"""
VARIANT型関係の機能を提供します。
"""

from ctypes import (
    POINTER,
    Union,
    byref,
    c_byte,
    c_int16,
    c_int32,
    c_uint16,
    c_wchar_p,
    windll,
)
from enum import IntFlag

from .core import ComResult, cotaskmem, cr


class VARENUM(IntFlag):
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_R4 = 4
    VT_R8 = 5
    VT_CY = 6
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_ERROR = 10
    VT_BOOL = 11
    VT_VARIANT = 12
    VT_UNKNOWN = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_UI1 = 17
    VT_UI2 = 18
    VT_UI4 = 19
    VT_I8 = 20
    VT_UI8 = 21
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_PTR = 26
    VT_SAFEARRAY = 27
    VT_CARRAY = 28
    VT_USERDEFINED = 29
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_RECORD = 36
    VT_INT_PTR = 37
    VT_UINT_PTR = 38
    VT_FILETIME = 64
    VT_BLOB = 65
    VT_STREAM = 66
    VT_STORAGE = 67
    VT_STREAMED_OBJECT = 68
    VT_STORED_OBJECT = 69
    VT_BLOB_OBJECT = 70
    VT_CF = 71
    VT_CLSID = 72
    VT_VERSIONED_STREAM = 73
    VT_BSTR_BLOB = 0xFFF
    VT_VECTOR = 0x1000
    VT_ARRAY = 0x2000
    VT_BYREF = 0x4000
    VT_RESERVED = 0x8000
    VT_ILLEGAL = 0xFFFF
    VT_ILLEGALMASKED = 0xFFF
    VT_TYPEMASK = 0xFFF


class Variant(Union):
    """VARIANT構造体のラッパー。"""

    # VARTYPE vt;
    # PROPVAR_PAD1 wReserved1;
    # PROPVAR_PAD2 wReserved2;
    # PROPVAR_PAD3 wReserved3;
    # u
    # DECIMAL decVal;
    _fields_ = (("vt", c_int16), ("data", c_byte * 24))

    __slots__ = ("vt", "data")

    def clear(self) -> None:
        _VariantClear(self)

    def __del__(self) -> None:
        self.clear()

    def clone(self) -> "Variant":
        pv = Variant()
        _VariantCopy(pv, self)
        return pv

    def to_str_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(_VariantToStringAlloc(self, byref(p)), p.value or "")

    def to_str(self) -> str:
        return self.to_str_nothrow().value

    def __str__(self) -> str:
        s = self.to_str_nothrow()
        return s.value_unchecked or "<ERROR>" if s else ""

    def __repr__(self) -> str:
        return f"PropVariant({self.__str__()})"

    def __enter__(self):
        return self

    def __exit__(self):
        self.clear()
        return True

    @property
    def vartype(self) -> VARENUM:
        return VARENUM(self.vt)

    @property
    def vartype_elem(self) -> VARENUM:
        return VARENUM(self.vt & VARENUM.VT_TYPEMASK)

    @property
    def is_array(self) -> bool:
        return self.vt & VARENUM.VT_ARRAY != 0

    @property
    def is_vector(self) -> bool:
        return self.vt & VARENUM.VT_VECTOR != 0

    @property
    def data_memview(self) -> memoryview:
        return memoryview(self.data)[8:]

    def change_type_nothrow(self, vt: VARENUM) -> "ComResult[Variant]":
        pv = Variant()
        return cr(_VariantChangeType(byref(pv), byref(self), 0, vt.value), pv)

    def change_type(self, vt: VARENUM) -> "Variant":
        return self.change_type_nothrow(vt).value

    @property
    def is_empty(self) -> bool:
        return self.vartype == VARENUM.VT_EMPTY

    @property
    def is_null(self) -> bool:
        return self.vartype == VARENUM.VT_NULL


# oleauto32

_VariantClear = windll.oleaut32.VariantClear
_VariantClear.argtypes = (POINTER(Variant),)
_VariantClear.restype = c_int32

_VariantCopy = windll.oleaut32.VariantCopy
_VariantCopy.argtypes = (POINTER(Variant), POINTER(Variant))
_VariantCopy.restype = c_int32

_VariantCopyInd = windll.oleaut32.VariantCopyInd
_VariantCopyInd.argtypes = (POINTER(Variant), POINTER(Variant))
_VariantCopyInd.restype = c_int32

_VariantChangeType = windll.oleaut32.VariantChangeType
_VariantChangeType.argtypes = (POINTER(Variant), POINTER(Variant), c_uint16, c_int32)
_VariantChangeType.restype = c_int32

# propsys

_VariantToStringAlloc = windll.propsys.VariantToStringAlloc
_VariantToStringAlloc.argtypes = (POINTER(Variant), POINTER(c_wchar_p))
_VariantToStringAlloc.restype = c_int32
