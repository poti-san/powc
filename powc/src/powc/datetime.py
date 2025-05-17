"""Windowsの日時型とPythonの日時型の変換。"""

import datetime as _datetime
from ctypes import c_uint32
from datetime import datetime, timedelta

from comtypes import Structure


class FILETIME(Structure):
    """Win32のFILETIME型です。アライメントの都合で構造体のまま扱うことが推奨されています。"""

    _fields_ = (
        ("low", c_uint32),
        ("high", c_uint32),
    )

    __slots__ = ()

    def __int__(self) -> int:
        """64ビット整数に変換します。"""
        return self.low | (self.high << 32)

    @property
    def datetime(self) -> datetime:
        return filetimeint64_to_datetime(int(self))

    @datetime.setter
    def datetime(self, dt: _datetime.datetime) -> None:
        x = filetimeint64_from_datetime(dt)
        self.low = x & 0xFFFFFFFF
        self.high = (x & 0xFFFFFFFF00000000) >> 32

    @staticmethod
    def from_datetime(dt: _datetime.datetime) -> "FILETIME":
        ft = FILETIME()
        ft.datetime = dt
        return ft


def filetimeint64_to_datetime(ft: int) -> datetime:
    """int型のFILETIMEをdatetime.datetime型に変換します。"""
    return datetime(1601, 1, 1, 0, 0, 0) + timedelta(microseconds=ft / 10)


# NOTE 変数を減らすためにメソッド内へ移動することも考慮
_MACROSECONDS_160101010000 = datetime(1601, 1, 1, 0, 0, 0).microsecond


def filetimeint64_from_datetime(dt: datetime) -> int:
    """datetime.datetime型をint型のFILETIMEに変換します。"""
    return (dt.microsecond - _MACROSECONDS_160101010000) * 10
