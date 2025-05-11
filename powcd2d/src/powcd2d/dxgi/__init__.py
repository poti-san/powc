from collections.abc import Buffer
from ctypes import byref, c_int64
from typing import Any, Iterator, Sequence

from powc.core import (
    ComResult,
    IUnknownPointer,
    IUnknownWrapper,
    cr,
    hr,
    query_interface,
)

from .. import _dxgi
from .types import *


class DXGIObject:
    """DXGIオブジェクト。IDXGIObjectインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDXGIObject)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IDXGIObject)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def set_privatedata(self, name: GUID, data: Buffer) -> ComResult[None]:
        if isinstance(data, bytes):
            return cr(self.__o.SetPrivateData(name, len(data), data), None)
        else:
            mv = memoryview(data)
            return cr(self.__o.SetPrivateData(name, len(mv), (c_byte * len(mv))(mv)), None)

    def set_privatedata_interface(self, name: GUID, data: IUnknownWrapper | IUnknownPointer) -> ComResult[None]:
        if isinstance(data, IUnknownWrapper):
            return cr(self.__o.SetPrivateDataInterface(name, data.wrapped_obj), None)
        elif isinstance(data, POINTER(IUnknown)):
            return cr(self.__o.SetPrivateDataInterface(name, data), None)
        else:
            raise TypeError

    # TODO
    def get_privatedatalen_nothrow(self, name: GUID) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetPrivateData(name, byref(x), None), x.value)

    def get_privatedatalen(self, name: GUID) -> int:
        return self.get_privatedatalen_nothrow(name).value

    def get_privatedata_nothrow(self, name: GUID) -> ComResult[bytes]:
        x = (c_byte * self.get_privatedatalen(name))()
        return cr(self.__o.GetPrivateData(name, byref(x), x), x.value)

    def get_privatedata(self, name: GUID) -> bytes:
        return self.get_privatedata_nothrow(name).value

    # TODO: get_privatedata as iunknown

    @property
    def parent_nothrow(self) -> "ComResult[DXGIObject]":
        x = POINTER(IDXGIObject)()
        return cr(self.__o.GetParent(IDXGIObject._iid_, byref(x)), DXGIObject(x))

    @property
    def parent(self) -> "DXGIObject":
        return self.parent_nothrow.value


class DXGIAdapter(DXGIObject):
    """DXGIアダプタ。IDXGIAdapterインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDXGIAdaptor)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IDXGIAdapter)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def get_output_nothrow(self, index: int) -> ComResult[DXGIObject]:
        x = POINTER(IDXGIObject)()
        return cr(self.__o.EnumOutputs(index, byref(x)), DXGIObject(x))

    def get_output(self, index: int) -> DXGIObject:
        ret = self.get_output_nothrow(index)
        if ret.hr == _DXGI_ERROR_NOT_FOUND:
            raise StopIteration
        return ret.value

    @property
    def output_iter(self) -> Iterator[DXGIObject]:
        for i in range(0xFFFFFFFF):
            try:
                yield self.get_output(i)
            except StopIteration:
                return
        raise OverflowError

    @property
    def desc_nothrow(self) -> ComResult[DXGIAdapterDesc]:
        x = DXGIAdapterDesc()
        return cr(self.__o.GetDesc(byref(x)), x)

    @property
    def desc(self) -> DXGIAdapterDesc:
        return self.desc_nothrow.value

    def check_interfacesupport_nothrow(self, iid: GUID) -> ComResult[bool]:
        x = c_int64()
        hr = self.__o.CheckInterfaceSupport(iid, byref(x))
        if hr == 0 or hr == _DXGI_ERROR_UNSUPPORTED:
            return cr(0, hr == 0)
        else:
            return cr(hr, x.value != 0)

    def check_interfacesupport(self, iid: GUID) -> bool:
        return self.check_interfacesupport_nothrow(iid).value


_DXGI_ERROR_NOT_FOUND = hr(0x887A0002)
_DXGI_ERROR_UNSUPPORTED = hr(0x887A0004)


class DXGIFactory(DXGIObject):
    """DXGIファクトリークラス。IDXGIFactoryインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDXGIFactory)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IDXGIFactory)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create_nothrow() -> "ComResult[DXGIFactory]":
        x = POINTER(IDXGIFactory)()
        return cr(_CreateDXGIFactory(IDXGIFactory._iid_, byref(x)), DXGIFactory(x))

    @staticmethod
    def create() -> "DXGIFactory":
        return DXGIFactory.create_nothrow().value

    def get_adapter_nothrow(self, index: int) -> ComResult[DXGIAdapter]:
        x = POINTER(IDXGIAdapter)()
        return cr(self.__o.EnumAdapters(index, byref(x)), DXGIAdapter(x))

    def get_adapter(self, index: int) -> DXGIAdapter:
        ret = self.get_adapter_nothrow(index)
        if ret.hr == _DXGI_ERROR_NOT_FOUND:
            raise StopIteration
        return ret.value

    @property
    def adapter_iter(self) -> Iterator[DXGIAdapter]:
        for i in range(0xFFFFFFFF):
            try:
                yield self.get_adapter(i)
            except StopIteration:
                return
        raise OverflowError


# STDMETHOD(c_int32, "MakeWindowAssociation", (c_void_p, c_uint32)),
# STDMETHOD(c_int32, "GetWindowAssociation", (POINTER(c_void_p),)),
# STDMETHOD(
#     c_int32, "CreateSwapChain", (POINTER(IUnknown), POINTER(DXGI_SWAP_CHAIN_DESC), POINTER(POINTER(IDXGISwapChain)))
# ),
# STDMETHOD(c_int32, "CreateSoftwareAdapter", (c_void_p, POINTER(POINTER(IDXGIAdapter)))),


class DXGIDevice(DXGIObject):
    """DXGIデバイス。IDXGIDeviceインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IDXGIDevice)

    def __init__(self, o: Any) -> None:
        super().__init__(o)
        self.__o = query_interface(o, IDXGIDevice)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def adapter_nothrow(self) -> ComResult[DXGIAdapter]:
        x = POINTER(IDXGIAdapter)()
        return cr(self.__o.GetAdapter(byref(x)), DXGIAdapter(x))

    @property
    def adapter(self) -> DXGIAdapter:
        return self.adapter_nothrow.value

    # STDMETHOD(
    #     c_int32,
    #     "CreateSurface",
    #     (POINTER(DXGI_SURFACE_DESC), c_uint32, c_int32, POINTER(DXGI_SHARED_RESOURCE), POINTER(POINTER(IDXGISurface))),
    # ),

    def query_resourceresidency_nothrow(
        self, resources: Sequence[IUnknownWrapper]
    ) -> ComResult[tuple[DXGIResidency, ...]]:
        l = len(resources)
        a = (POINTER(IUnknown) * l)(resource.wrapped_obj for resource in resources)
        x = (c_int32 * l)()
        return cr(self.__o.QueryResourceResidency(a, x, l), tuple(DXGIResidency(i) for i in x))

    def query_resourceresidency(self, resources: Sequence[IUnknownWrapper]) -> tuple[DXGIResidency, ...]:
        return self.query_resourceresidency_nothrow(resources).value

    def set_gputhreadpriority_nothrow(self, value: DXGIResourcePriority) -> ComResult[None]:
        return cr(self.__o.SetGPUThreadPriority(int(value)), None)

    @property
    def gputhreadpriority_nothrow(self) -> ComResult[DXGIResourcePriority]:
        x = c_int32()
        return cr(self.__o.GetGPUThreadPriority(byref(x)), DXGIResourcePriority(x.value))

    @property
    def gputhreadpriority(self) -> DXGIResourcePriority:
        return self.gputhreadpriority_nothrow.value

    @gputhreadpriority.setter
    def gputhreadpriority(self, value: DXGIResourcePriority) -> None:
        self.set_gputhreadpriority_nothrow(value)


_CreateDXGIFactory = _dxgi.CreateDXGIFactory
_CreateDXGIFactory.argtypes = (POINTER(GUID), POINTER(POINTER(IUnknown)))

_CreateDXGIFactory1 = _dxgi.CreateDXGIFactory1
_CreateDXGIFactory1.argtypes = (POINTER(GUID), POINTER(POINTER(IUnknown)))
