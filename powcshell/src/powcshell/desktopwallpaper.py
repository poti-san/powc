"""デスクトップの壁紙。

主なクラスは :class:`DesktopWallpaper` です。
"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from ctypes.wintypes import RECT
from dataclasses import dataclass
from enum import IntEnum, IntFlag
from typing import Any, Iterator

from comtypes import GUID, STDMETHOD, CoCreateInstanceEx, IUnknown

from powc.core import ComResult, cotaskmem, cr, query_interface

from .shellitemarray import IShellItemArray, ShellItemArray


class DesktopSlideshowOption(IntFlag):
    """DESKTOP_SLIDESHOW_OPTIONS"""

    SHUFFLE_IMAGES = 0x1


class DesktopSlideshowState(IntFlag):
    """DESKTOP_SLIDESHOW_STATE"""

    ENABLED = 0x1
    SLIDESHOW = 0x2
    DISABLED_BY_REMOTE_SESSION = 0x4


class DesktopSlideshowDirection(IntEnum):
    """DESKTOP_SLIDESHOW_DIRECTION"""

    FORWARD = 0
    BACKWARD = 1


class DesktopWallpaperPosition(IntEnum):
    """DESKTOP_WALLPAPER_POSITION"""

    CENTER = 0
    TILE = 1
    STRETCH = 2
    FIT = 3
    FILL = 4
    SPAN = 5


class IDesktopWallpaper(IUnknown):
    _iid_ = GUID("{B92B56A9-8B55-4E14-9A89-0199BBB6F93B}")
    _methods_ = [
        STDMETHOD(c_int32, "SetWallpaper", (c_wchar_p, c_wchar_p)),
        STDMETHOD(c_int32, "GetWallpaper", (c_wchar_p, POINTER(c_wchar_p))),
        STDMETHOD(c_int32, "GetMonitorDevicePathAt", (c_uint32, POINTER(c_wchar_p))),
        STDMETHOD(c_int32, "GetMonitorDevicePathCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetMonitorRECT", (c_wchar_p, POINTER(RECT))),
        STDMETHOD(c_int32, "SetBackgroundColor", (c_uint32,)),
        STDMETHOD(c_int32, "GetBackgroundColor", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "SetPosition", (c_int32,)),
        STDMETHOD(c_int32, "GetPosition", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "SetSlideshow", (POINTER(IShellItemArray),)),
        STDMETHOD(c_int32, "GetSlideshow", (POINTER(POINTER(IShellItemArray)),)),
        STDMETHOD(c_int32, "SetSlideshowOptions", (c_int32, c_uint32)),
        STDMETHOD(c_int32, "GetSlideshowOptions", (POINTER(c_int32), POINTER(c_uint32))),
        STDMETHOD(c_int32, "AdvanceSlideshow", (c_wchar_p, c_int32)),
        STDMETHOD(c_int32, "GetStatus", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "Enable", (c_int32,)),
    ]
    __slots__ = ()


class DesktopWallpaper:
    """デスクトップ壁紙の設定。IDesktopWallpaperのラッパーです。

    Examples:
        >>> from powcshell.desktopwallpaper import DesktopWallpaper
        >>>
        >>> wallpaper = DesktopWallpaper.create()
        >>>
        >>> print(
        >>>     f\"\"\"
        >>>     壁紙:{wallpaper.get_wallpaper(0)}
        >>>     表示位置:{wallpaper.position}
        >>>     背景色:{wallpaper.bgcolor}
        >>>     モニタデバイスパス:{tuple(wallpaper.monitor_device_paths)}
        >>>     \"\"\"
        >>> )
    """

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDesktopWallpaper)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDesktopWallpaper)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "DesktopWallpaper":
        CLSID_DesktopWallpaper = GUID("{C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD}")
        return DesktopWallpaper(CoCreateInstanceEx(CLSID_DesktopWallpaper, IDesktopWallpaper))

    def set_wallpaper_nothrow(self, monitor_id_or_index: str | int, wallpaper: str) -> ComResult[None]:
        if isinstance(monitor_id_or_index, int):
            ret = self.get_monitor_device_path_at_nothrow(monitor_id_or_index)
            if not ret:
                return cr(ret.hr, None)
            monitor_id_or_index = ret.value_unchecked

        return cr(self.__o.SetWallpaper(monitor_id_or_index, wallpaper), None)

    def set_wallpaper(self, monitor_id_or_index: str | int, wallpaper: str) -> None:
        return self.set_wallpaper_nothrow(monitor_id_or_index, wallpaper).value

    @property
    def monitor_device_path_count_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetMonitorDevicePathCount(byref(x)), x.value)

    @property
    def monitor_device_path_count(self) -> int:
        return self.monitor_device_path_count_nothrow.value

    def get_wallpaper_nothrow(self, monitor_id_or_index: str | int) -> ComResult[str]:
        if isinstance(monitor_id_or_index, int):
            ret = self.get_monitor_device_path_at_nothrow(monitor_id_or_index)
            if not ret:
                return cr(ret.hr, "")
            monitor_id_or_index = ret.value_unchecked

        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetWallpaper(monitor_id_or_index, byref(p)), p.value or "")

    def get_wallpaper(self, monitor_id_or_index: str | int) -> str:
        return self.get_wallpaper_nothrow(monitor_id_or_index).value

    def get_monitor_device_path_at_nothrow(self, monitor_index: int) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetMonitorDevicePathAt(monitor_index, p), p.value or "")

    def get_monitor_device_path_at(self, monitor_index: int) -> str:
        return self.get_monitor_device_path_at_nothrow(monitor_index).value

    @property
    def monitor_device_paths(self) -> Iterator[str]:
        get_monitor_device_path_at = self.get_monitor_device_path_at
        return (get_monitor_device_path_at(i) for i in range(self.monitor_device_path_count))

    @property
    def wallpapers(self) -> Iterator[str]:
        get_wallpaper = self.get_wallpaper
        return (get_wallpaper(i) for i in range(self.monitor_device_path_count))

    def get_monitor_rect_nothrow(self, monitor_id_or_index: int | str) -> ComResult[RECT]:
        if isinstance(monitor_id_or_index, int):
            ret = self.get_monitor_device_path_at_nothrow(monitor_id_or_index)
            if not ret:
                return cr(ret.hr, RECT())
            monitor_id_or_index = ret.value_unchecked

        x = RECT()
        return cr(self.__o.GetMonitorRECT(monitor_id_or_index, byref(x)), x)

    def get_monitor_rect(self, monitor_id_or_index: int | str) -> RECT:
        return self.get_monitor_rect_nothrow(monitor_id_or_index).value

    @property
    def monitor_rects(self) -> Iterator[RECT]:
        get_monitor_rect = self.get_monitor_rect
        return (get_monitor_rect(i) for i in range(self.monitor_device_path_count))

    @property
    def bgcolor_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetBackgroundColor(byref(x)), x.value)

    @property
    def bgcolor(self) -> int:
        return self.bgcolor_nothrow.value

    def set_bgcolor_nothrow(self, value: int) -> ComResult[None]:
        return cr(self.__o.SetBackgroundColor(value), None)

    @bgcolor.setter
    def bgcolor(self, value: int) -> None:
        self.set_bgcolor_nothrow(value).value

    @property
    def position_nothrow(self) -> ComResult[DesktopWallpaperPosition]:
        x = c_int32()
        return cr(self.__o.GetPosition(byref(x)), DesktopWallpaperPosition(x.value))

    @property
    def position(self) -> DesktopWallpaperPosition:
        return self.position_nothrow.value

    def set_position_nothrow(self, value: DesktopWallpaperPosition | int) -> ComResult[None]:
        return self.__o.SetPosition(int(value))

    @position.setter
    def position(self, value: DesktopWallpaperPosition | int) -> None:
        return self.set_position_nothrow(value).value

    @property
    def slideshow_nothrow(self) -> ComResult[ShellItemArray]:
        x = POINTER(IShellItemArray)()
        return cr(self.__o.GetSlideshow(byref(x)), ShellItemArray(x))

    @property
    def slideshow(self) -> ShellItemArray:
        return self.slideshow_nothrow.value

    def set_slideshow_nothrow(self, value: ShellItemArray) -> ComResult[None]:
        return self.__o.SetSlideshow(value.wrapped_obj)

    @slideshow.setter
    def slideshow(self, value: ShellItemArray) -> None:
        return self.set_slideshow_nothrow(value).value

    @dataclass(frozen=True)
    class SlideshowOptions:
        options: int | DesktopSlideshowState
        slideshow_tick: int

    @property
    def slideshow_options_nothrow(self) -> ComResult[SlideshowOptions]:
        x1 = c_int32()
        x2 = c_uint32()
        return cr(
            self.__o.GetSlideshowOptions(byref(x1), byref(x2)), DesktopWallpaper.SlideshowOptions(x1.value, x2.value)
        )

    @property
    def slideshow_options(self) -> SlideshowOptions:
        return self.slideshow_options_nothrow.value

    def set_slideshow_options_nothrow(self, value: SlideshowOptions) -> ComResult[None]:
        return self.__o.SetSlideshowOptions(int(value.options), value.slideshow_tick)

    @slideshow_options.setter
    def slideshow_options(self, value: SlideshowOptions) -> None:
        return self.set_slideshow_options_nothrow(value).value

    def advance_slideshow_nothrow(
        self, monitor_id_or_index: int | str, direction: DesktopSlideshowDirection
    ) -> ComResult[None]:
        if isinstance(monitor_id_or_index, int):
            ret = self.get_monitor_device_path_at_nothrow(monitor_id_or_index)
            if not ret:
                return cr(ret.hr, None)
            monitor_id_or_index = ret.value_unchecked

        return cr(self.__o.AdvanceSlideshow(monitor_id_or_index, int(direction)), None)

    @property
    def status_nothrow(self) -> ComResult[DesktopSlideshowState]:
        x = c_int32()
        return cr(self.__o.GetStatus(byref(x)), DesktopSlideshowState(x.value))

    @property
    def status(self) -> DesktopSlideshowState:
        return self.status_nothrow.value

    def set_enable_nothrow(self, value: bool) -> ComResult[None]:
        x = c_int32()
        return cr(self.__o.Enable(byref(x), 1 if value else 0), None)

    def set_enable(self, value: bool) -> None:
        return self.set_enable_nothrow(value).value
        return self.set_enable_nothrow(value).value
