from contextlib import contextmanager
from ctypes import c_size_t, c_uint32, c_void_p
from enum import IntFlag
from typing import Any, Iterator

from . import _kernel32


class GlobalHandleFlag(IntFlag):
    """グローバルハンドルのフラグ。"""

    FIXED = 0x0000
    MOVEABLE = 0x0002
    ZEROINIT = 0x0040
    HANDLE = MOVEABLE | ZEROINIT
    POINTER = FIXED | ZEROINIT


_GlobalAlloc = _kernel32.GlobalAlloc
_GlobalAlloc.argtypes = (c_uint32, c_size_t)
_GlobalAlloc.restype = c_void_p


_GlobalFree = _kernel32.GlobalFree
_GlobalFree.argtypes = (c_void_p,)
_GlobalFree.restype = c_void_p


@contextmanager
def globalmem[T](p: T) -> Iterator[T]:
    """グローバルメモリをスコープ管理します。

    Examples:
    >>> from comtypes import c_void_p
    >>> with globalmem(c_void_p()) as p:
    >>>     # pの割り当てや使用
    """
    try:
        yield p
    finally:
        _GlobalFree(p)


def globalmemalloc(size: int) -> c_void_p:
    """グローバルメモリを確保します。

    Args:
        size (int): バイト数。

    Returns:
        c_void_p: 確保したメモリ。
    """
    return _GlobalAlloc(size)


def globalmemfree(p: int | c_void_p | Any) -> None:
    """グローバルメモリを解放します。

    Args:
        p (int | c_void_p | Any): グローバルメモリのポインタ。
    """
    return _GlobalFree(p)
