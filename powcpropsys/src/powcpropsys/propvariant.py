"""プロパティの値。"""

from ctypes import (
    POINTER,
    Union,
    byref,
    c_byte,
    c_double,
    c_float,
    c_int,
    c_int8,
    c_int16,
    c_int32,
    c_int64,
    c_size_t,
    c_ssize_t,
    c_uint,
    c_uint8,
    c_uint16,
    c_uint32,
    c_uint64,
    c_void_p,
    c_wchar_p,
    cast,
    sizeof,
)
from datetime import datetime

from comtypes import GUID

from powc.core import ComResult, CoTaskMem, check_hresult, cotaskmem, cotaskmem_free, cr
from powc.datetime import FILETIME
from powc.variant import VARENUM

from . import _ole32, _propsys


class PropVariant(Union):
    """プロパティシステムのプロパティ値。PROPVARIANT構造体のラッパーです。"""

    # VARTYPE vt;
    # PROPVAR_PAD1 wReserved1;
    # PROPVAR_PAD2 wReserved2;
    # PROPVAR_PAD3 wReserved3;
    _fields_ = (("vt", c_uint16), ("data", c_byte * (24 if sizeof(c_void_p) == 8 else 16)))

    __slots__ = ()

    def clear(self) -> None:
        _PropVariantClear(self)

    def __del__(self) -> None:
        self.clear()

    def clone(self) -> "PropVariant":
        pv = PropVariant()
        _PropVariantCopy(pv, self)
        return pv

    def to_str_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(_PropVariantToStringAlloc(self, byref(p)), p.value or "")

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
        return memoryview(self.data)[PropVariant.data.offset :]

    def change_type_nothrow(self, vt: VARENUM) -> "ComResult[PropVariant]":
        global PropVariantChangeType
        pv = PropVariant()
        return cr(_PropVariantChangeType(byref(pv), byref(self), 0, vt.value), pv)

    def change_type(self, vt: VARENUM) -> "PropVariant":
        return self.change_type_nothrow(vt).value

    @property
    def is_empty(self) -> bool:
        return self.vartype == VARENUM.VT_EMPTY

    @property
    def is_null(self) -> bool:
        return self.vartype == VARENUM.VT_NULL

    # initializers

    @staticmethod
    def init_int8(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_I1
        c_int8.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_int16(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_I2
        c_int16.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_int32(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_I4
        c_int32.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_int64(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_I8
        c_int64.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_int(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_INT
        c_int.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_intptr(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_INT_PTR
        c_ssize_t.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_uint8(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_UI1
        c_uint8.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_uint16(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_UI2
        c_uint16.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_uint32(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_UI4
        c_uint32.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_uint64(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_UI8
        c_uint64.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_uint(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_UINT
        c_uint.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_uintptr(x: int):
        v = PropVariant()
        v.vt = VARENUM.VT_UINT_PTR
        c_size_t.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_float(x: float):
        v = PropVariant()
        v.vt = VARENUM.VT_R4
        c_float.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_double(x: float):
        v = PropVariant()
        v.vt = VARENUM.VT_R8
        c_double.from_buffer(v.data_memview).value = x
        return v

    @staticmethod
    def init_bool(x: bool):
        v = PropVariant()
        v.vt = VARENUM.VT_BOOL
        c_int32.from_buffer(v.data_memview).value = 1 if x else 0
        return v

    @staticmethod
    def init_wstr(x: str):
        v = PropVariant()
        v.vt = VARENUM.VT_LPWSTR
        p = CoTaskMem.alloc_unistr(x)
        c_void_p.from_buffer(v.data_memview).value = p.detatch()
        return v

    @staticmethod
    def init_filetime(x: datetime):
        v = PropVariant()
        v.vt = VARENUM.VT_FILETIME
        FILETIME.from_buffer(v.data_memview).datetime = x
        return v

    @staticmethod
    def init_clsid(x: GUID):
        v = PropVariant()
        v.vt = VARENUM.VT_CLSID
        (c_byte * 16).from_buffer(v.data_memview).value = bytes(x)
        return v

    # getters

    def get_int8(self):
        if self.vartype != VARENUM.VT_I1:
            raise TypeError
        return c_int8.from_buffer(self.data_memview).value

    def get_int16(self):
        if self.vartype != VARENUM.VT_I2:
            raise TypeError
        return c_int16.from_buffer(self.data_memview).value

    def get_int32(self):
        if self.vartype != VARENUM.VT_I4:
            raise TypeError
        return c_int32.from_buffer(self.data_memview).value

    def get_int64(self):
        if self.vartype != VARENUM.VT_I8:
            raise TypeError
        return c_int64.from_buffer(self.data_memview).value

    def get_wstr(self):
        if self.vartype != VARENUM.VT_LPWSTR:
            raise TypeError
        return c_wchar_p.from_buffer(self.data_memview).value or ""

    def get_float(self):
        if self.vartype != VARENUM.VT_R4:
            raise TypeError
        return c_float.from_buffer(self.data_memview).value

    def get_double(self):
        if self.vartype != VARENUM.VT_R8:
            raise TypeError
        return c_double.from_buffer(self.data_memview).value

    def get_bool(self) -> bool:
        if self.vartype != VARENUM.VT_BOOL:
            raise TypeError
        return c_int32.from_buffer(self.data_memview).value != 0

    # VT_CY = 6
    # VT_DATE = 7

    def get_bstr(self) -> str:
        if self.vartype != VARENUM.VT_BSTR:
            raise TypeError
        return c_wchar_p.from_buffer(self.data_memview).value or ""

    # VT_DISPATCH = 9
    # VT_ERROR = 10
    # VT_BOOL = 11
    # VT_VARIANT = 12
    # VT_UNKNOWN = 13
    # VT_DECIMAL = 14

    def get_uint8(self):
        if self.vartype != VARENUM.VT_UI1:
            raise TypeError
        return c_uint8.from_buffer(self.data_memview).value

    def get_uint16(self):
        if self.vartype != VARENUM.VT_UI2:
            raise TypeError
        return c_uint16.from_buffer(self.data_memview).value

    def get_uint32(self):
        if self.vartype != VARENUM.VT_UI4:
            raise TypeError
        return c_uint32.from_buffer(self.data_memview).value

    def get_uint64(self):
        if self.vartype != VARENUM.VT_UI8:
            raise TypeError
        return c_uint64.from_buffer(self.data_memview).value

    def get_int(self):
        if self.vartype != VARENUM.VT_INT:
            raise TypeError
        return c_uint.from_buffer(self.data_memview).value

    def get_uint(self):
        if self.vartype != VARENUM.VT_UINT:
            raise TypeError
        return c_uint.from_buffer(self.data_memview).value

    def get_intptr(self):
        if self.vartype != VARENUM.VT_INT_PTR:
            raise TypeError
        return c_ssize_t.from_buffer(self.data_memview).value

    def get_uintptr(self):
        if self.vartype != VARENUM.VT_UINT_PTR:
            raise TypeError
        return c_size_t.from_buffer(self.data_memview).value

    def get_filetime(self) -> datetime:
        if self.vartype != VARENUM.VT_FILETIME:
            raise TypeError
        return FILETIME.from_buffer(self.data_memview).datetime

    # VT_VOID = 24
    # VT_HRESULT = 25
    # VT_PTR = 26
    # VT_SAFEARRAY = 27
    # VT_CARRAY = 28
    # VT_USERDEFINED = 29
    # VT_LPSTR = 30
    # VT_RECORD = 36
    # VT_FILETIME = 64
    # VT_BLOB = 65
    # VT_STREAM = 66
    # VT_STORAGE = 67
    # VT_STREAMED_OBJECT = 68
    # VT_STORED_OBJECT = 69
    # VT_BLOB_OBJECT = 70
    # VT_CF = 71
    # VT_CLSID = 72
    # VT_VERSIONED_STREAM = 73
    # VT_BSTR_BLOB = 0xFFF
    # VT_VECTOR = 0x1000
    # VT_ARRAY = 0x2000
    # VT_BYREF = 0x4000

    @property
    def elemcount(self):
        return _PropVariantGetElementCount(byref(self))

    def get_elem(self, index: int) -> "PropVariant":
        pv = PropVariant()
        check_hresult(_InitPropVariantFromPropVariantVectorElem(byref(self), index, byref(pv)))
        return pv

    def to_strings(self) -> tuple[str, ...]:
        pp = c_void_p()
        len = c_uint32()
        check_hresult(_PropVariantToStringVectorAlloc(byref(self), byref(pp), byref(len)))

        parray = (c_void_p * len.value).from_address(pp.value or 0)
        try:
            return tuple(cast(parray[i], c_wchar_p).value or "" for i in range(len.value))
        finally:
            for i in range(len.value):
                cotaskmem_free(parray[i])
            cotaskmem_free(pp)


_PropVariantClear = _ole32.PropVariantClear
_PropVariantClear.argtypes = (POINTER(PropVariant),)
_PropVariantClear.restype = c_int32

_PropVariantCopy = _ole32.PropVariantCopy
_PropVariantCopy.argtypes = (POINTER(PropVariant), POINTER(PropVariant))
_PropVariantCopy.restype = c_int32

_PropVariantToStringAlloc = _propsys.PropVariantToStringAlloc
_PropVariantToStringAlloc.argtypes = (POINTER(PropVariant), POINTER(c_wchar_p))
_PropVariantToStringAlloc.restype = c_int32

_PropVariantChangeType = _propsys.PropVariantChangeType
_PropVariantChangeType.argtypes = (POINTER(PropVariant), POINTER(PropVariant), c_int32, c_int32)
_PropVariantChangeType.restype = c_int32

_PropVariantGetElementCount = _propsys.PropVariantGetElementCount
_PropVariantGetElementCount.argtypes = (POINTER(PropVariant),)
_PropVariantGetElementCount.restype = c_int32

_InitPropVariantFromPropVariantVectorElem = _propsys.InitPropVariantFromPropVariantVectorElem
_InitPropVariantFromPropVariantVectorElem.argtypes = (
    POINTER(PropVariant),
    POINTER(POINTER(c_wchar_p)),
    POINTER(c_uint32),
)
_InitPropVariantFromPropVariantVectorElem.restype = c_int32

_PropVariantToStringVectorAlloc = _propsys.PropVariantToStringVectorAlloc
_PropVariantToStringVectorAlloc.argtypes = (POINTER(PropVariant), POINTER(c_void_p), POINTER(c_uint32))
_PropVariantToStringVectorAlloc.restype = c_int32
_PropVariantToStringVectorAlloc.restype = c_int32
