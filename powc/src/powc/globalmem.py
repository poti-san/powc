"""グローバルメモリの管理。"""

from contextlib import contextmanager
from ctypes import _SimpleCData, c_size_t, c_uint32, c_void_p
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

_GlobalLock = _kernel32.GlobalLock
_GlobalLock.argtypes = (c_void_p,)
_GlobalLock.restype = c_void_p

_GlobalUnlock = _kernel32.GlobalUnlock
_GlobalUnlock.argtypes = (c_void_p,)
_GlobalUnlock.restype = c_void_p

_GlobalSize = _kernel32.GlobalSize
_GlobalSize.argtypes = (c_void_p,)
_GlobalSize.restype = c_size_t


@contextmanager
def globalmem[T](p: T) -> Iterator[T]:
    """グローバルメモリの寿命をスコープ管理します。

    Examples:
    >>> from comtypes import c_void_p
    >>> with globalmem(c_void_p()) as p:
    >>>     # pの割り当てや使用
    """
    try:
        yield p
    finally:
        _GlobalFree(p)


@contextmanager
def globalmem_lock[T: _SimpleCData](handle: _SimpleCData | int, t: type[T]) -> Iterator[tuple[T, int]]:
    """グローバルメモリのロックをスコープ管理します。"""
    p = 0
    try:
        p: int = _GlobalLock(handle.value if isinstance(handle, _SimpleCData) else int(handle))
        yield (t(p), _GlobalSize(p))
    finally:
        if p != 0:
            _GlobalUnlock(handle)


def globalmem_alloc(size: int) -> c_void_p:
    """グローバルメモリを確保します。

    Args:
        size (int): バイト数。

    Returns:
        c_void_p: 確保したメモリ。
    """
    return _GlobalAlloc(size)


def globalmem_free(p: int | c_void_p | Any) -> None:
    """グローバルメモリを解放します。

    Args:
        p (int | c_void_p | Any): グローバルメモリのポインタ。
    """
    return _GlobalFree(p)
