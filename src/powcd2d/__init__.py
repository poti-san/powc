"""powcd2dライブラリの初期化。

DXGI、Direct2D、Direct2D、WICはインターフェイスを共有するため、先に全て初期化します。
ライブラリ分類は索引性のためです"""

from ctypes import WinDLL

_dwrite = WinDLL("dwrite.dll")
_dxgi = WinDLL("dxgi.dll")

# fmt: off
# isort: off
from .d2d.interfacebody import *
#from .dcommon.interfacebody import *
from .dxgi.interfacebody import *
from .dwrite.interfacebody import *
from .wic.interfacebody import *

from .d2d.types import *
from .dcommon.types import *
from .dxgi.types import *
from .dwrite.types import *
from .wic.types import *
