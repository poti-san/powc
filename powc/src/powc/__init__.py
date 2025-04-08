from ctypes import WinDLL

_ole32 = WinDLL("ole32.dll")
_kernel32 = WinDLL("kernel32.dll")
_oleaut32 = WinDLL("oleaut32.dll")
_propsys = WinDLL("propsys.dll")
_shlwapi = WinDLL("shlwapi.dll")
