"""プロパティキー。"""

from ctypes import POINTER, Structure, byref, c_int32, c_uint32, c_wchar_p

from comtypes import GUID

from powc.core import ComResult, cotaskmem, cr, guid_from_define

from . import _propsys


class PropertyKey(Structure):
    """プロパティシステムのプロパティキー。PROPERTYKEY構造体のラッパーです。"""

    _fields_ = (("fmtid", GUID), ("pid", c_uint32))

    __slots__ = ()

    @staticmethod
    def from_define(
        a: int, b: int, c: int, d: int, e: int, f: int, g: int, h: int, i: int, j: int, k: int, l: int  # noqa: E741
    ) -> "PropertyKey":
        return PropertyKey(guid_from_define(a, b, c, d, e, f, g, h, i, j, k), l)

    def __str__(self) -> str:
        return f"{self.fmtid} {self.pid}"

    def __repr__(self) -> str:
        cname = self.canonicalname_nothrow
        return f'PropertyKey({self.fmtid} {self.pid}, "{cname.value_unchecked if cname else ""}")'

    def __eq__(self, other) -> bool:
        if other is not PropertyKey:
            return False
        return self.fmtid == other.fmtid and self.pid == other.fmtid

    def __hash__(self) -> int:
        """ハッシュを計算します。mutableなのでハッシュは可変です。"""
        return hash(bytes(self))

    @staticmethod
    def from_canonicalname_nothrow(name: str) -> "ComResult[PropertyKey]":
        global _PSGetPropertyKeyFromName
        x = PropertyKey()
        return cr(_PSGetPropertyKeyFromName(name, byref(x)), x)

    @staticmethod
    def from_canonicalname(name: str) -> "PropertyKey":
        return PropertyKey.from_canonicalname_nothrow(name).value

    @property
    def canonicalname_nothrow(self) -> ComResult[str]:
        global _PSGetNameFromPropertyKey
        with cotaskmem(c_wchar_p()) as p:
            return cr(_PSGetNameFromPropertyKey(self, byref(p)), p.value or "")

    @property
    def canonicalname(self) -> str:
        return self.canonicalname_nothrow.value


_PSGetPropertyKeyFromName = _propsys.PSGetPropertyKeyFromName
_PSGetPropertyKeyFromName.argtypes = (c_wchar_p, POINTER(PropertyKey))
_PSGetPropertyKeyFromName.restype = c_int32

_PSGetNameFromPropertyKey = _propsys.PSGetNameFromPropertyKey
_PSGetNameFromPropertyKey.argtypes = (POINTER(PropertyKey), POINTER(c_wchar_p))
_PSGetNameFromPropertyKey.restype = c_int32
