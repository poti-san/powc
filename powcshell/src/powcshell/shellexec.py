"""パスまたはアイテムIDリストのシェル操作。

主なクラスは :class:`ShellExecute` です。"""

from ctypes import (
    POINTER,
    Structure,
    WinError,
    byref,
    c_int32,
    c_uint32,
    c_void_p,
    c_wchar_p,
    get_last_error,
    sizeof,
)
from enum import IntEnum
from os import PathLike

from . import _shell32


class ShowCommand(IntEnum):
    HIDE = 0
    SHOWNORMAL = 1
    NORMAL = 1
    SHOWMINIMIZED = 2
    SHOWMAXIMIZED = 3
    MAXIMIZE = 3
    SHOWNOACTIVATE = 4
    SHOW = 5
    MINIMIZE = 6
    SHOWMINNOACTIVE = 7
    SHOWNA = 8
    RESTORE = 9
    SHOWDEFAULT = 10
    FORCEMINIMIZE = 11


# _SEE_MASK_DEFAULT = 0x00000000
_SEE_MASK_CLASSNAME = 0x00000001
_SEE_MASK_CLASSKEY = 0x00000003
_SEE_MASK_IDLIST = 0x00000004
_SEE_MASK_INVOKEIDLIST = 0x0000000C
_SEE_MASK_HOTKEY = 0x00000020
_SEE_MASK_NOCLOSEPROCESS = 0x00000040
_SEE_MASK_CONNECTNETDRV = 0x00000080
_SEE_MASK_NOASYNC = 0x00000100
# _SEE_MASK_FLAG_DDEWAIT = _SEE_MASK_NOASYNC
_SEE_MASK_DOENVSUBST = 0x00000200
_SEE_MASK_FLAG_NO_UI = 0x00000400
# _SEE_MASK_UNICODE = 0x00004000
_SEE_MASK_NO_CONSOLE = 0x00008000
_SEE_MASK_ASYNCOK = 0x00100000
_SEE_MASK_HMONITOR = 0x00200000
_SEE_MASK_NOZONECHECKS = 0x00800000
# _SEE_MASK_NOQUERYCLASSSTORE = 0x01000000
_SEE_MASK_WAITFORINPUTIDLE = 0x02000000
_SEE_MASK_FLAG_LOG_USAGE = 0x04000000


class ShellExecuteOption(IntEnum):
    DEFAULT = 0
    NO_CLOSE_PROCESS = _SEE_MASK_NOCLOSEPROCESS
    CONNECT_NETWORKDRIVE = _SEE_MASK_CONNECTNETDRV
    NO_ASYNC = _SEE_MASK_NOASYNC
    EXPAND_ENVVARS = _SEE_MASK_DOENVSUBST
    NO_UI = _SEE_MASK_FLAG_NO_UI
    NO_CONSOLE = _SEE_MASK_NO_CONSOLE
    ASYNC_OK = _SEE_MASK_ASYNCOK
    NO_ZONECHECKS = _SEE_MASK_NOZONECHECKS
    WAIT_FOR_INPUTIDLE = _SEE_MASK_WAITFORINPUTIDLE
    LOG_USAGE = _SEE_MASK_FLAG_LOG_USAGE


class _SHELLEXECUTEINFOW(Structure):
    __slots__ = ()
    _fields_ = (
        ("cbSize", c_uint32),
        ("fMask", c_uint32),
        ("hwnd", c_void_p),
        ("lpVerb", c_wchar_p),
        ("lpFile", c_wchar_p),
        ("lpParameters", c_wchar_p),
        ("lpDirectory", c_wchar_p),
        ("nShow", c_int32),
        ("hInstApp", c_void_p),
        ("lpIDList", c_void_p),
        ("lpClass", c_wchar_p),
        ("hkeyClass", c_void_p),
        ("dwHotKey", c_uint32),
        ("hMonitor", c_void_p),
        ("hProcess", c_void_p),
    )


_ShellExecuteExW = _shell32.ShellExecuteExW
_ShellExecuteExW.argtypes = (POINTER(_SHELLEXECUTEINFOW),)
_ShellExecuteExW.restype = c_int32


class ShellExecute:
    """パスまたはアイテムIDリストのシェル操作機能を提供します。"""

    @staticmethod
    def execute_pidl(
        pidl: int,
        verb: str,
        invokes: bool = True,
        params: str | None = None,
        dir: str | None = None,
        showcmd: ShowCommand | int = ShowCommand.SHOWNORMAL,
        hotkey: int | None = None,
        monitor_handle: int | None = None,
        options: ShellExecuteOption = ShellExecuteOption.DEFAULT,
    ) -> int:

        see = _SHELLEXECUTEINFOW()
        see.cbSize = sizeof(see)
        see.fMask = (
            _SEE_MASK_IDLIST
            | (_SEE_MASK_INVOKEIDLIST if invokes else 0)
            | (_SEE_MASK_HOTKEY if hotkey else 0)
            | (_SEE_MASK_HMONITOR if monitor_handle else 0)
            | options
        )
        see.lpVerb = verb
        see.lpParameters = params
        see.lpDirectory = dir
        see.nShow = int(showcmd)
        see.lpIDList = pidl
        see.hMonitor = monitor_handle if monitor_handle else 0
        see.dwHotKey = hotkey if hotkey else 0
        if not _ShellExecuteExW(byref(see)):
            raise WinError(get_last_error())
        return see.hProcess

    @staticmethod
    def execute_path(
        path: str | PathLike,
        verb: str,
        invokes: bool = True,
        params: str | None = None,
        dir: str | None = None,
        showcmd: ShowCommand | int = ShowCommand.SHOWNORMAL,
        hotkey: int | None = None,
        monitor_handle: int | None = None,
        options: ShellExecuteOption = ShellExecuteOption.DEFAULT,
    ) -> int:

        see = _SHELLEXECUTEINFOW()
        see.cbSize = sizeof(see)
        see.fMask = (
            (_SEE_MASK_INVOKEIDLIST if invokes else 0)
            | (_SEE_MASK_HOTKEY if hotkey else 0)
            | (_SEE_MASK_HMONITOR if monitor_handle else 0)
            | options
        )
        see.lpFile = str(path)
        see.lpVerb = verb
        see.lpParameters = params
        see.lpDirectory = dir
        see.nShow = int(showcmd)
        see.hMonitor = monitor_handle if monitor_handle else 0
        see.dwHotKey = hotkey if hotkey else 0
        if not _ShellExecuteExW(byref(see)):
            raise WinError(get_last_error())
        return see.hProcess

    @staticmethod
    def execute_class(
        classname: str,
        verb: str,
        invokes: bool = True,
        params: str | None = None,
        dir: str | None = None,
        showcmd: ShowCommand | int = ShowCommand.SHOWNORMAL,
        hotkey: int | None = None,
        monitor_handle: int | None = None,
        options: ShellExecuteOption = ShellExecuteOption.DEFAULT,
    ) -> int:

        see = _SHELLEXECUTEINFOW()
        see.cbSize = sizeof(see)
        see.fMask = (
            _SEE_MASK_CLASSNAME
            | (_SEE_MASK_INVOKEIDLIST if invokes else 0)
            | (_SEE_MASK_HOTKEY if hotkey else 0)
            | (_SEE_MASK_HMONITOR if monitor_handle else 0)
            | options
        )
        see.lpClass = classname
        see.lpVerb = verb
        see.lpParameters = params
        see.lpDirectory = dir
        see.nShow = int(showcmd)
        see.hMonitor = monitor_handle if monitor_handle else 0
        see.dwHotKey = hotkey if hotkey else 0
        if not _ShellExecuteExW(byref(see)):
            raise WinError(get_last_error())
        return see.hProcess

    @staticmethod
    def execute_regkey(
        regkey_handle: int,
        verb: str,
        invokes: bool = True,
        params: str | None = None,
        dir: str | None = None,
        showcmd: ShowCommand | int = ShowCommand.SHOWNORMAL,
        hotkey: int | None = None,
        monitor_handle: int | None = None,
        options: ShellExecuteOption = ShellExecuteOption.DEFAULT,
    ) -> int:

        see = _SHELLEXECUTEINFOW()
        see.cbSize = sizeof(see)
        see.fMask = (
            _SEE_MASK_CLASSKEY
            | (_SEE_MASK_INVOKEIDLIST if invokes else 0)
            | (_SEE_MASK_HOTKEY if hotkey else 0)
            | (_SEE_MASK_HMONITOR if monitor_handle else 0)
            | options
        )
        see.hkeyClass = regkey_handle
        see.lpVerb = verb
        see.lpParameters = params
        see.lpDirectory = dir
        see.nShow = int(showcmd)
        see.hMonitor = monitor_handle if monitor_handle else 0
        see.dwHotKey = hotkey if hotkey else 0
        if not _ShellExecuteExW(byref(see)):
            raise WinError(get_last_error())
        return see.hProcess
