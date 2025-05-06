"""COM基本機能。多くは他のライブラリの内部で使用されます。"""

from ctypes import WinDLL

_ole32 = WinDLL("ole32.dll")
_kernel32 = WinDLL("kernel32.dll")
_user32 = WinDLL("user32.dll")
_oleaut32 = WinDLL("oleaut32.dll")
_propsys = WinDLL("propsys.dll")
_shlwapi = WinDLL("shlwapi.dll")
