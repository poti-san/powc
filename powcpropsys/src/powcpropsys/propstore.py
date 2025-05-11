"""プロパティストア。"""

from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p
from enum import IntFlag
from typing import Any, Iterable, Iterator, OrderedDict

from comtypes import GUID, STDMETHOD, IUnknown

from powc.core import ComResult, cr, query_interface

from . import _propsys
from .propkey import PropertyKey
from .propsys import PropertySystem
from .propvariant import PropVariant


class GetPropertyStoreFlag(IntFlag):
    """GETPROPERTYSTOREFLAGS"""

    DEFAULT = 0
    HANDLER_PROPERTIES_ONLY = 0x1
    READ_WRITE = 0x2
    TEMPORARY = 0x4
    FAST_PROPERTIES_ONLY = 0x8
    OPEN_SLOW_ITEM = 0x10
    DELAY_CREATION = 0x20
    BEST_EFFORT = 0x40
    NO_OPLOCK = 0x80
    PREFER_QUERY_PROPERTIES = 0x100
    EXTRINSIC_PROPERTIES = 0x200
    EXTRINSIC_PROPERTIES_ONLY = 0x400
    VOLATILE_PROPERTIES = 0x800
    VOLATILE_PROPERTIES_ONLY = 0x1000


class IPropertyStore(IUnknown):
    """"""

    _iid_ = GUID("{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}")
    _methods_ = [
        STDMETHOD(c_int32, "GetCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetAt", (c_uint32, POINTER(PropertyKey))),
        STDMETHOD(c_int32, "GetValue", (POINTER(PropertyKey), POINTER(PropVariant))),
        STDMETHOD(c_int32, "SetValue", (POINTER(PropertyKey), POINTER(PropVariant))),
        STDMETHOD(c_int32, "Commit", ()),
    ]
    __slots__ = ()


class PropertyStore:
    """プロパティストア。IPropertyStoreインターフェイスのラッパーです。"""

    __o: Any  # IPropertyStore

    __slots__ = ("__o",)

    def __init__(self, o: Any):
        self.__o = query_interface(o, IPropertyStore)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def createonmem_nothrow() -> "ComResult[PropertyStore]":
        """メモリ内プロパティストアを作成します。

        このプロパティストアは以下のインターフェイスを実装しています。
        IPropertyStore、INamedPropertyStore、IPropertyStoreCache、
        IPersistStream、IPropertyBag、IPersistSerializedPropStorage
        """
        p = POINTER(IPropertyStore)()
        return cr(_PSCreateMemoryPropertyStore(IPropertyStore._iid_, byref(p)), PropertyStore(p))

    @staticmethod
    def createonmem() -> "PropertyStore":
        return PropertyStore.createonmem_nothrow().value

    createonmem.__doc__ = createonmem_nothrow.__doc__

    @property
    def count_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetCount(byref(x)), x.value)

    @property
    def count(self) -> int:
        return self.count_nothrow.value

    def get_key_at_nothrow(self, index: int) -> ComResult[PropertyKey]:
        x = PropertyKey()
        return cr(self.__o.GetAt(index, byref(x)), x)

    def get_key_at(self, index: int) -> PropertyKey:
        return self.get_key_at_nothrow(index).value

    def get_value_nothrow(self, key: PropertyKey) -> ComResult[PropVariant]:
        x = PropVariant()
        return cr(self.__o.GetValue(byref(key), byref(x)), x)

    def get_value(self, key: PropertyKey) -> PropVariant:
        return self.get_value_nothrow(key).value

    def set_value_nothrow(self, key: PropertyKey, value: PropVariant) -> ComResult[None]:
        return cr(self.__o.SetValue(byref(key), byref(value)), None)

    def set_value(self, key: PropertyKey, value: PropVariant) -> None:
        return self.set_value_nothrow(key, value).value

    def commit_nothrow(self) -> ComResult[None]:
        return cr(self.__o.Commit(), None)

    def commit(self) -> None:
        return self.commit_nothrow().value

    def __enter__(self):
        return self

    def __exit__(self, ex_type, ex_value, trace):
        self.commit()

    def iter_keys(self) -> Iterator[PropertyKey]:
        return (self.get_key_at(i) for i in range(self.count))

    @property
    def keys(self) -> tuple[PropertyKey, ...]:
        return tuple(self.iter_keys())

    def iter_items(self) -> Iterator[tuple[PropertyKey, PropVariant]]:
        return ((key, self.get_value(key)) for key in self.iter_keys())

    @property
    def items(self) -> tuple[tuple[PropertyKey, PropVariant], ...]:
        return tuple(self.iter_items())

    @property
    def itemdict(self) -> OrderedDict[PropertyKey, PropVariant]:
        return OrderedDict(self.iter_items())

    def iter_keys_in_propsystem(self, propsys: PropertySystem | None = None) -> Iterator[PropertyKey]:
        """
        プロパティシステムに登録されたキーでプロパティストアでも有効なキーのイテレーターを返します。
        返されるキーにはプロパティストア自体からは列挙できないキーも含まれます。
        """

        propsys = propsys if propsys else PropertySystem.create()
        keys = propsys.get_propkeys_system()
        return self.iter_keys_in_keys(keys)

    def iter_keys_in_keys(self, keys: Iterable[PropertyKey]) -> Iterator[PropertyKey]:
        """
        キーに対応するキーのイテレーターを返します。
        返されるキーにはプロパティストア自体からは列挙できないキーも含まれます。
        """
        return (key for key, value in ((key, self.get_value_nothrow(key)) for key in keys) if not value.value.is_empty)

    def iter_items_in_propsystem(
        self, propsys: PropertySystem | None = None
    ) -> Iterable[tuple[PropertyKey, PropVariant]]:
        """
        プロパティシステムに登録されたキーに対応する項目のイテレーターを返します。
        返される項目にはプロパティストア自体に含まれるキーでは取得できない項目も含まれます。
        """

        propsys = propsys if propsys else PropertySystem.create()
        keys = propsys.get_propkeys_system()
        return self.iter_items_in_keys(keys)

    def iter_items_in_keys(self, keys: Iterable[PropertyKey]) -> Iterable[tuple[PropertyKey, PropVariant]]:
        """
        キーに対応する項目のイテレーターを返します。
        返される項目にはプロパティストア自体に含まれるキーでは取得できない項目も含まれます。
        """

        return (
            (key, value.value_unchecked)
            for key, value in ((key, self.get_value_nothrow(key)) for key in keys)
            if not value.value.is_empty
        )


_PSCreateMemoryPropertyStore = _propsys.PSCreateMemoryPropertyStore
_PSCreateMemoryPropertyStore.argtypes = (POINTER(GUID), POINTER(POINTER(IUnknown)))
_PSCreateMemoryPropertyStore.restype = c_int32
_PSCreateMemoryPropertyStore.restype = c_int32
