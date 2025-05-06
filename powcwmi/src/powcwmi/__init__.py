"""WMI。 :class:`WBEMLocator` 経由で :class:`WBEMServices` を取得した後、
WMIクラスやインスタンスの情報を取得できます。"""

from csv import Error
from ctypes import POINTER, WinDLL, byref, c_int32, c_uint32, c_void_p
from dataclasses import dataclass
from enum import IntEnum, IntFlag
from typing import Any, Iterator, Literal, NamedTuple, OrderedDict

from comtypes import BSTR, GUID, CoCreateInstance
from powc.comsec import com_init_security, com_set_securityblanket
from powc.core import ComResult, check_hresult, cr, query_interface
from powc.safearray import SafeArrayPtr
from powc.variant import Variant

from .comtypes import *
from .wbemflags import *
from .wbemstatus import *


class WBEMClassObjectGetMethodException(Error):
    def __init__(self) -> None:
        super().__init__("メソッドの取得に失敗しました。メソッドはWBEMクラスからのみ取得できます。")


class WBEMGenus(IntEnum):
    """WBEM_GENUS_TYPE"""

    CLASS = 1
    INSTANCE = 2


class WBEMStatus(IntFlag):
    """WBEM_STATUS"""

    COMPLETE = 0
    REQUIREMENTS = 1
    PROGRESS = 2
    LOGGING_INFORMATION = 0x100
    LOGGING_INFORMATION_PROVIDER = 0x200
    LOGGING_INFORMATION_HOST = 0x400
    LOGGING_INFORMATION_REPOSITORY = 0x800
    LOGGING_INFORMATION_ESS = 0x1000


class WBEMTimeout(IntEnum):
    """WBEM_TIMEOUT_TYPE"""

    NO_WAIT = 0
    INFINITE = -1


class WBEMFlavor(IntFlag):
    """WBEM_FLAVOR"""

    DONT_PROPAGATE = 0
    FLAG_PROPAGATE_TO_INSTANCE = 0x1
    FLAG_PROPAGATE_TO_DERIVED_CLASS = 0x2
    MASK_PROPAGATION = 0xF
    OVERRIDABLE = 0
    NOT_OVERRIDABLE = 0x10
    MASK_PERMISSIONS = 0x10
    ORIGIN_LOCAL = 0
    ORIGIN_PROPAGATED = 0x20
    ORIGIN_SYSTEM = 0x40
    MASK_ORIGIN = 0x60
    NOT_AMENDED = 0
    AMENDED = 0x80
    MASK_AMENDED = 0x80


class CimType(IntFlag):
    """CIMTYPE_ENUMERATION"""

    ILLEGAL = 0xFFF
    EMPTY = 0
    SINT8 = 16
    UINT8 = 17
    SINT16 = 2
    UINT16 = 18
    SINT32 = 3
    UINT32 = 19
    SINT64 = 20
    UINT64 = 21
    REAL32 = 4
    REAL64 = 5
    BOOLEAN = 11
    STRING = 8
    DATETIME = 101
    REFERENCE = 102
    CHAR16 = 103
    OBJECT = 13
    ARRAY = 0x2000


class WBEMStatusFormat(IntEnum):
    """WBEMSTATUS_FORMAT"""

    NEWLINE = 0
    NO_NEWLINE = 1


class WBEMLimits:
    """WBEM_LIMITS"""

    MAX_IDENTIFIER = 0x1000
    MAX_QUERY = 0x4000
    MAX_PATH = 0x2000
    MAX_OBJECT_NESTING = 64
    MAX_USER_PROPERTIES = 1024


class WBEMExtraReturnCode(IntEnum):
    """WBEM_EXTRA_RETURN_CODES"""

    S_INITIALIZED = 0
    S_LIMITED_SERVICE = 0x43001
    S_INDIRECTLY_UPDATED = S_LIMITED_SERVICE + 1
    S_SUBJECT_TO_SDS = S_INDIRECTLY_UPDATED + 1
    E_RETRY_LATER = 0x80043001
    E_RESOURCE_CONTENTION = E_RETRY_LATER + 1


class WMIObjectText(IntEnum):
    """WMI_OBJ_TEXT"""

    CIM_DTD_2_0 = 1
    WMI_DTD_2_0 = 2


class WBEMLocator:
    """WBEMロケーター。IWbemLocatorインターフェイスのラッパーです。

    :func:`connect_server` で :class:`WBEMServices` を作成できます。
    """

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWbemLocator)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWbemLocator)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "WBEMLocator":
        """WBEMロケーターを作成します。同時にCOMセキュリティをWMI用に初期化します。"""
        com_init_security()
        return WBEMLocator(CoCreateInstance(GUID("{4590f811-1d3a-11d0-891f-00aa004b2e24}"), IWbemLocator))

    @staticmethod
    def create_nosecinit() -> "WBEMLocator":
        return WBEMLocator(CoCreateInstance(GUID("{4590f811-1d3a-11d0-891f-00aa004b2e24}"), IWbemLocator))

    # TODO pCtx対応
    def connect_server_nothrow(
        self,
        netres: str,
        user: str | None = None,
        password: str | None = None,
        locale: str | None = None,
        waits_max: bool = False,
        authority: str | None = None,
    ) -> "ComResult[WBEMServices]":
        flags = WBEM_FLAG_CONNECT_USE_MAX_WAIT if waits_max else 0
        x = POINTER(IWbemServices)()
        return cr(
            self.__o.ConnectServer(netres, user, password, locale, flags, authority, None, byref(x)),
            WBEMServices(x),
        )

    def connect_server(
        self,
        network_resource: str | Literal["root\\default", "root\\cimv2"],
        user: str | None = None,
        password: str | None = None,
        locale: str | None = None,
        waits_max: bool = False,
        authority: str | None = None,
    ) -> "WBEMServices":
        return self.connect_server_nothrow(network_resource, user, password, locale, waits_max, authority).value


# TODO: async系メソッドの追加
class WBEMServices:
    """WBEMサービス。IWbemServicesインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWbemServices)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWbemServices)
        com_set_securityblanket(self.wrapped_obj)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def open_namespace_nothrow(self, namespace: str) -> "ComResult[WBEMServices]":
        x = POINTER(IWbemServices)()
        return cr(self.__o.OpenNamespace(namespace, 0, None, byref(x), None), WBEMServices(x))

    def open_namespace(self, namespace: str) -> "WBEMServices":
        return self.open_namespace_nothrow(namespace).value

    def get_object_nothrow(
        self,
        path: str,
        uses_amended_qualifiers: bool = True,
        direct_read: bool = True,
    ) -> "ComResult[WBEMClassObject]":
        """関連付けられた名前空間に存在するクラスまたはインスタンスを取得します。"""
        flags = (WBEM_FLAG_USE_AMENDED_QUALIFIERS if uses_amended_qualifiers else 0) | (
            WBEM_FLAG_DIRECT_READ if direct_read else 0
        )
        x = POINTER(IWbemClassObject)()
        return cr(self.__o.GetObject(path, flags, None, byref(x), None), WBEMClassObject(x))

    def get_object(
        self,
        path: str,
        uses_amended_qualifiers: bool = True,
        direct_read: bool = True,
    ) -> "WBEMClassObject":
        """関連付けられた名前空間に存在するクラスまたはインスタンスを取得します。"""
        return self.get_object_nothrow(path, uses_amended_qualifiers, direct_read).value

    class UpdateMode(IntEnum):
        COMPATIBLE = WBEM_FLAG_UPDATE_COMPATIBLE
        SAFE_MODE = WBEM_FLAG_UPDATE_SAFE_MODE
        FORCE_MODE = WBEM_FLAG_UPDATE_FORCE_MODE

    class UpdateCreateMode(IntEnum):
        BOTH = 0
        UPDATE_ONLY = WBEM_FLAG_UPDATE_ONLY
        CREATE_ONLY = WBEM_FLAG_CREATE_ONLY

    def put_class_nothrow(
        self,
        class_: "WBEMClassObject",
        uses_amended_qualifiers: bool = True,
        updatecreate: UpdateCreateMode = UpdateCreateMode.BOTH,
        owner_update: bool = True,
        updatemode: UpdateMode = UpdateMode.COMPATIBLE,
    ) -> ComResult[None]:
        flags = (
            (WBEM_FLAG_USE_AMENDED_QUALIFIERS if uses_amended_qualifiers else 0)
            | int(updatecreate)
            | (WBEM_FLAG_OWNER_UPDATE if owner_update else 0)
            | int(updatemode)
        )
        return cr(self.__o.PutClass(class_.wrapped_obj, flags, None, None), None)

    def put_class(
        self,
        class_: "WBEMClassObject",
        uses_amended_qualifiers: bool = True,
        updatecreate: UpdateCreateMode = UpdateCreateMode.BOTH,
        owner_update: bool = True,
        updatemode: UpdateMode = UpdateMode.COMPATIBLE,
    ) -> None:
        return self.put_class_nothrow(class_, uses_amended_qualifiers, updatecreate, owner_update, updatemode).value

    def delete_class_nothrow(self, class_name: str, owner_update: bool = True) -> ComResult[None]:
        flags = WBEM_FLAG_OWNER_UPDATE if owner_update else 0
        return cr(self.__o.DeleteClass(class_name, flags, None, None), None)

    def delete_class(self, class_name: str, owner_update: bool = True) -> None:
        return self.delete_class_nothrow(class_name, owner_update).value

    def create_classenum_nothrow(
        self,
        superclass: str | None,
        uses_amended_qualifiers: bool = True,
        shallow: bool = True,
        forward_only: bool = True,
    ) -> "ComResult[WBEMClassObjectEnumerator]":
        flags = (
            (WBEM_FLAG_USE_AMENDED_QUALIFIERS if uses_amended_qualifiers else 0)
            | (WBEM_FLAG_SHALLOW if shallow else WBEM_FLAG_DEEP)
            | (WBEM_FLAG_FORWARD_ONLY if forward_only else WBEM_FLAG_BIDIRECTIONAL)
        )
        x = POINTER(IEnumWbemClassObject)()
        return cr(self.__o.CreateClassEnum(superclass, flags, None, byref(x)), WBEMClassObjectEnumerator(x))

    def create_classenum(
        self,
        superclass: str | None,
        uses_amended_qualifiers: bool = True,
        shallow: bool = True,
        forward_only: bool = True,
    ) -> "WBEMClassObjectEnumerator":
        return self.create_classenum_nothrow(superclass, uses_amended_qualifiers, shallow, forward_only).value

    def create_classenum_toplevel_nothrow(self) -> "ComResult[WBEMClassObjectEnumerator]":
        return self.create_classenum_nothrow(None, shallow=True)

    def create_classenum_toplevel(self) -> "WBEMClassObjectEnumerator":
        return self.create_classenum_toplevel_nothrow().value

    def put_instance_nothrow(
        self,
        instance: "WBEMClassObject",
        updatecreate: UpdateCreateMode = UpdateCreateMode.BOTH,
        uses_amended_qualifiers: bool = True,
    ) -> ComResult[None]:
        flags = int(updatecreate) | (WBEM_FLAG_USE_AMENDED_QUALIFIERS if uses_amended_qualifiers else 0)
        return cr(self.__o.PutInstance(instance.wrapped_obj, flags, None, None), None)

    def put_instance(
        self,
        instance: "WBEMClassObject",
        updatecreate: UpdateCreateMode = UpdateCreateMode.BOTH,
        uses_amended_qualifiers: bool = True,
    ) -> None:
        return self.put_instance_nothrow(instance, updatecreate, uses_amended_qualifiers).value

    def delete_instance_nothrow(self, path: str) -> ComResult[None]:
        return cr(self.__o.DeleteInstance(path, 0, None, None), None)

    def delete_instance(self, path: str) -> None:
        return self.delete_instance_nothrow(path).value

    def create_instanceenum_nothrow(
        self,
        filter: str | None,
        uses_amended_qualifiers: bool = True,
        shallow: bool = True,
        forward_only: bool = True,
        direct_read: bool = False,
    ) -> "ComResult[WBEMClassObjectEnumerator]":
        flags = (
            (WBEM_FLAG_USE_AMENDED_QUALIFIERS if uses_amended_qualifiers else 0)
            | (WBEM_FLAG_SHALLOW if shallow else WBEM_FLAG_DEEP)
            | (WBEM_FLAG_FORWARD_ONLY if forward_only else WBEM_FLAG_BIDIRECTIONAL)
            | (WBEM_FLAG_DIRECT_READ if direct_read else 0)
        )
        x = POINTER(IEnumWbemClassObject)()
        return cr(self.__o.CreateInstanceEnum(filter, flags, None, byref(x)), WBEMClassObjectEnumerator(x))

    def create_instanceenum(
        self,
        filter: str | None,
        uses_amended_qualifiers: bool = True,
        shallow: bool = True,
        forward_only: bool = True,
        direct_read: bool = False,
    ) -> "WBEMClassObjectEnumerator":
        return self.create_instanceenum_nothrow(
            filter, uses_amended_qualifiers, shallow, forward_only, direct_read
        ).value

    def exec_query_nothrow(
        self,
        query: str | None,
        uses_amended_qualifiers: bool = True,
        forward_only: bool = True,
        direct_read: bool = False,
        ensures_locatable: bool = False,
        prototype: bool = False,
    ) -> "ComResult[WBEMClassObjectEnumerator]":
        flags = (
            (WBEM_FLAG_USE_AMENDED_QUALIFIERS if uses_amended_qualifiers else 0)
            | (WBEM_FLAG_FORWARD_ONLY if forward_only else WBEM_FLAG_BIDIRECTIONAL)
            | (WBEM_FLAG_DIRECT_READ if direct_read else 0)
            | (WBEM_FLAG_ENSURE_LOCATABLE if ensures_locatable else 0)
            | (WBEM_FLAG_PROTOTYPE if prototype else 0)
        )
        x = POINTER(IEnumWbemClassObject)()
        return cr(self.__o.ExecQuery("WQL", query, flags, None, byref(x)), WBEMClassObjectEnumerator(x))

    def exec_query(
        self,
        query: str | None,
        uses_amended_qualifiers: bool = True,
        forward_only: bool = True,
        direct_read: bool = False,
        ensures_locatable: bool = False,
        prototype: bool = False,
    ) -> "WBEMClassObjectEnumerator":
        return self.exec_query_nothrow(
            query,
            uses_amended_qualifiers,
            forward_only,
            direct_read,
            ensures_locatable,
            prototype,
        ).value

    def exec_notificationquery_nothrow(self, query: str) -> ComResult["WBEMClassObjectEnumerator"]:
        x = POINTER(IEnumWbemClassObject)()
        return cr(
            self.__o.ExecNotificationQuery(
                "WQL", query, WBEM_FLAG_RETURN_IMMEDIATELY | WBEM_FLAG_FORWARD_ONLY, None, byref(x)
            ),
            WBEMClassObjectEnumerator(x),
        )

    def exec_notificationquery(self, query: str) -> "WBEMClassObjectEnumerator":
        return self.exec_notificationquery_nothrow(query).value

    def exec_method_nothrow(
        self, path: str, methodname: str, inparams: "WBEMClassObject"
    ) -> ComResult["WBEMClassObject"]:
        x = POINTER(IWbemClassObject)()
        return cr(
            self.__o.ExecMethod(path, methodname, 0, None, inparams.wrapped_obj, byref(x), None), WBEMClassObject(x)
        )

    def exec_method(self, path: str, methodname: str, inparams: "WBEMClassObject") -> "WBEMClassObject":
        return self.exec_method_nothrow(path, methodname, inparams).value


class WBEMClassObjectEnumerator:
    """WBEMクラスオブジェクト列挙子。IEnumWbemClassObjectインターフェイスのラッパーです。"""

    __slots__ = ("__o", "__timeout")
    __o: Any  # POINTER(IEnumWbemClassObject)
    __timeout: int

    def __init__(self, o: Any, timeout: int = WBEMTimeout.INFINITE) -> None:
        self.__o = query_interface(o, IEnumWbemClassObject)
        self.__timeout = timeout

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __iter__(self) -> "Iterator[WBEMClassObject]":
        ret = c_uint32()
        while True:
            x = POINTER(IWbemClassObject)()
            hr = self.__o.Next(self.__timeout, 1, byref(x), byref(ret))
            if hr != 0:
                if hr == 1:
                    break
                check_hresult(hr)
            yield WBEMClassObject(x)


class WBEMClassObject:
    """WBEMクラスオブジェクト。IWbemClassObjectインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWbemClassObject)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWbemClassObject)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def qualifierset_nothrow(self) -> "ComResult[WBEMQualifierSet]":
        x = POINTER(IWbemQualifierSet)()
        return cr(self.__o.GetQualifierSet(byref(x)), WBEMQualifierSet(x))

    @property
    def qualifierset(self) -> "WBEMQualifierSet":
        return self.qualifierset_nothrow.value

    class Property(NamedTuple):
        value: Variant
        type: CimType
        flavor: WBEMFlavor

    def get_nothrow(self, name: str) -> ComResult[Property]:
        x1 = Variant()
        x2 = c_int32()
        x3 = c_int32()
        return cr(
            self.__o.Get(name, 0, byref(x1), byref(x2), byref(x2)),
            WBEMClassObject.Property(x1, CimType(x2.value), WBEMFlavor(x3.value)),
        )

    def get(self, name: str) -> Property:
        return self.get_nothrow(name).value

    def put_nothrow(self, name: str, value: Variant | None, type: CimType) -> ComResult[None]:
        return cr(self.__o.Put(name, 0, value, int(type)), None)

    def put(self, name: str, value: Variant | None, type: CimType) -> None:
        return self.put_nothrow(name, value, type).value

    def delete_nothrow(self, name: str) -> ComResult[None]:
        return cr(self.__o.Delete(name), None)

    def delete(self, name: str) -> None:
        return self.delete_nothrow(name).value

    class GetNamesFlag1(IntEnum):
        ALL_PROPS = WBEM_FLAG_ALWAYS
        QUALNAME_ONLY_IF_TRUE = WBEM_FLAG_ONLY_IF_TRUE
        QUALNAME_ONLY_IF_FALSE = WBEM_FLAG_ONLY_IF_FALSE
        QUALNAME_AND_QUALVAL_ONLY_IF_IDENTICAL = WBEM_FLAG_ONLY_IF_IDENTICAL

    class GetNamesFlag2(IntEnum):
        ALL = 0
        KEYS_ONLY = WBEM_FLAG_KEYS_ONLY
        REFS_ONLY = WBEM_FLAG_REFS_ONLY

    class GetNamesFlag3(IntEnum):
        ALL = 0
        LOCAL_ONLY = WBEM_FLAG_LOCAL_ONLY
        PROPAGATED_ONLY = WBEM_FLAG_PROPAGATED_ONLY
        SYSTEM_ONLY = WBEM_FLAG_SYSTEM_ONLY
        NONSYSTEM_ONLY = WBEM_FLAG_NONSYSTEM_ONLY

    def get_names_nothrow(
        self,
        flag1: GetNamesFlag1,
        flag2: GetNamesFlag2,
        flag3: GetNamesFlag3,
        qualifier_name: str | None = None,
        qualifier_value: Variant | None = None,
    ) -> ComResult[tuple[str, ...]]:
        sa = SafeArrayPtr()
        hr = self.__o.GetNames(qualifier_name, int(flag1) | int(flag2) | int(flag3), qualifier_value, byref(sa))
        if hr != 0:
            return cr(hr, ())
        return cr(hr, sa.to_bstrarray())

    def get_names(
        self,
        flag1: GetNamesFlag1,
        flag2: GetNamesFlag2,
        flag3: GetNamesFlag3,
        qualifier_name: str | None = None,
        qualifier_value: Variant | None = None,
    ) -> tuple[str, ...]:
        return self.get_names_nothrow(flag1, flag2, flag3, qualifier_name, qualifier_value).value

    def get_propnames(self, flag1: GetNamesFlag1, flag2: GetNamesFlag2, flag3: GetNamesFlag3) -> tuple[str, ...]:
        check_hresult(self.__o.BeginEnumeration(int(flag1) | int(flag2) | int(flag3)))
        try:
            list = []
            while True:
                x = BSTR()
                hr = self.__o.Next(0, byref(x), None, None, None)
                if hr != 0:
                    if hr == WBEM_S_NO_MORE_DATA:
                        break
                    check_hresult(hr)
                list.append(x.value)
            return tuple(list)
        finally:
            self.__o.EndEnumeration()

    @property
    def propnames_all(self) -> tuple[str, ...]:
        return self.get_propnames(
            WBEMClassObject.GetNamesFlag1.ALL_PROPS,
            WBEMClassObject.GetNamesFlag2.ALL,
            WBEMClassObject.GetNamesFlag3.ALL,
        )

    @property
    def propnames_system(self) -> tuple[str, ...]:
        return self.get_propnames(
            WBEMClassObject.GetNamesFlag1.ALL_PROPS,
            WBEMClassObject.GetNamesFlag2.ALL,
            WBEMClassObject.GetNamesFlag3.SYSTEM_ONLY,
        )

    @property
    def propnames_nonsystem(self) -> tuple[str, ...]:
        return self.get_propnames(
            WBEMClassObject.GetNamesFlag1.ALL_PROPS,
            WBEMClassObject.GetNamesFlag2.ALL,
            WBEMClassObject.GetNamesFlag3.NONSYSTEM_ONLY,
        )

    def get_props(
        self,
        flag1: GetNamesFlag1,
        flag2: GetNamesFlag2,
        flag3: GetNamesFlag3,
    ) -> OrderedDict[str, Property]:
        check_hresult(self.__o.BeginEnumeration(int(flag1) | int(flag2) | int(flag3)))
        try:
            d = OrderedDict()
            while True:
                name = BSTR()
                val = Variant()
                type = c_int32()
                flavor = c_int32()
                hr = self.__o.Next(0, byref(name), byref(val), byref(type), byref(flavor))
                if hr != 0:
                    if hr == WBEM_S_NO_MORE_DATA:
                        break
                    check_hresult(hr)
                d[name.value] = (WBEMClassObject.Property(val, CimType(type.value), WBEMFlavor(flavor.value)),)
            return d
        finally:
            self.__o.EndEnumeration()

    @property
    def props_all(self) -> OrderedDict[str, Property]:
        return self.get_props(
            WBEMClassObject.GetNamesFlag1.ALL_PROPS,
            WBEMClassObject.GetNamesFlag2.ALL,
            WBEMClassObject.GetNamesFlag3.ALL,
        )

    @property
    def props_system(self) -> OrderedDict[str, Property]:
        return self.get_props(
            WBEMClassObject.GetNamesFlag1.ALL_PROPS,
            WBEMClassObject.GetNamesFlag2.ALL,
            WBEMClassObject.GetNamesFlag3.SYSTEM_ONLY,
        )

    @property
    def props_nonsystem(self) -> OrderedDict[str, Property]:
        return self.get_props(
            WBEMClassObject.GetNamesFlag1.ALL_PROPS,
            WBEMClassObject.GetNamesFlag2.ALL,
            WBEMClassObject.GetNamesFlag3.NONSYSTEM_ONLY,
        )

    def get_propqualifierset_nothrow(self, name: str) -> "ComResult[WBEMQualifierSet]":
        x = POINTER(IWbemQualifierSet)()
        return cr(self.__o.GetPropertyQualifierSet(name, byref(x)), WBEMQualifierSet(x))

    def get_propqualifierset(self, name: str) -> "WBEMQualifierSet":
        return self.get_propqualifierset_nothrow(name).value

    @property
    def propqualifiersetdict_all(self) -> OrderedDict[str, "WBEMQualifierSet"]:
        return OrderedDict(((name, self.get_propqualifierset(name)) for name in self.propnames_all))

    @property
    def propqualifiersetdict_system(self) -> OrderedDict[str, "WBEMQualifierSet"]:
        return OrderedDict(((name, self.get_propqualifierset(name)) for name in self.props_system))

    @property
    def propqualifiersetdict_nonsystem(self) -> OrderedDict[str, "WBEMQualifierSet"]:
        return OrderedDict(((name, self.get_propqualifierset(name)) for name in self.propnames_nonsystem))

    def clone_nothrow(self) -> "ComResult[WBEMClassObject]":
        x = POINTER(IWbemClassObject)()
        return cr(self.__o.Clone(byref(x)), WBEMClassObject(x))

    def clone(self) -> "WBEMClassObject":
        return self.clone_nothrow().value

    @property
    def objtext_nothrow(self) -> ComResult[str]:
        x = BSTR()
        return cr(self.__o.GetObjectText(0, byref(x)), x.value)

    @property
    def objtext(self) -> str:
        return self.objtext_nothrow.value

    @property
    def objtext_noflavors_nothrow(self) -> ComResult[str]:
        x = BSTR()
        return cr(self.__o.GetObjectText(WBEM_FLAG_NO_FLAVORS, byref(x)), x.value)

    @property
    def objtext_noflavors(self) -> str:
        return self.objtext_noflavors_nothrow.value

    def spawn_derivedcls_nothrow(self) -> "ComResult[WBEMClassObject]":
        x = POINTER(IWbemClassObject)()
        return cr(self.__o.SpawnDerivedClass(0, byref(x)), WBEMClassObject(x))

    def spawn_derivedcls(self) -> "WBEMClassObject":
        return self.spawn_derivedcls_nothrow().value

    def spawn_derivedinstance_nothrow(self) -> "ComResult[WBEMClassObject]":
        x = POINTER(IWbemClassObject)()
        return cr(self.__o.SpawnInstance(0, byref(x)), WBEMClassObject(x))

    def spawn_derivedinstance(self) -> "WBEMClassObject":
        return self.spawn_derivedinstance_nothrow().value

    class CompareFlag(IntFlag):
        ALL = 0
        IGNORE_OBJECT_SOURCE = WBEM_FLAG_IGNORE_OBJECT_SOURCE
        IGNORE_QUALIFIERS = WBEM_FLAG_IGNORE_QUALIFIERS
        IGNORE_DEFAULT_VALUES = WBEM_FLAG_IGNORE_DEFAULT_VALUES
        IGNORE_FLAVOR = WBEM_FLAG_IGNORE_FLAVOR
        IGNORE_CASE = WBEM_FLAG_IGNORE_CASE
        IGNORE_CLASS = WBEM_FLAG_IGNORE_CLASS

    def compare_to_nothrow(self, other: "WBEMClassObject", flags: CompareFlag = CompareFlag.ALL) -> ComResult[bool]:
        hr = self.__o.CompareTo(int(flags), other.wrapped_obj)
        return cr(hr, hr == 0)

    def compare_to(self, other: "WBEMClassObject", flags: CompareFlag = CompareFlag.ALL) -> bool:
        return self.compare_to_nothrow(other, flags).value

    def get_proporigin_nothrow(self, name: str) -> ComResult[str]:
        x = BSTR()
        return cr(self.__o.GetPropertyOrigin(name, byref(x)), x.value)

    def get_proporigin(self, name: str) -> str:
        return self.get_proporigin_nothrow(name).value

    def proporigins_all(self) -> OrderedDict[str, str]:
        return OrderedDict(((name, self.get_proporigin(name)) for name in self.propnames_all))

    def proporigins_system(self) -> OrderedDict[str, str]:
        return OrderedDict(((name, self.get_proporigin(name)) for name in self.propnames_system))

    def proporigins_nonsystem(self) -> OrderedDict[str, str]:
        return OrderedDict(((name, self.get_proporigin(name)) for name in self.propnames_nonsystem))

    def inherits_from_nothrow(self, ancester: str) -> ComResult[bool]:
        hr = self.__o.InheritsFrom(ancester)
        return cr(hr, hr == 0)

    def inherits_from(self, ancester: str) -> bool:
        return self.inherits_from_nothrow(ancester).value

    def get_method_nothrow(self, name: str) -> "ComResult[WBEMClassObject.MethodInfo]":
        x1 = POINTER(IWbemClassObject)()
        x2 = POINTER(IWbemClassObject)()
        return cr(
            self.__o.GetMethod(name, 0, byref(x1), byref(x2)),
            WBEMClassObject.MethodInfo(WBEMClassObject(x1), WBEMClassObject(x2)),
        )

    def get_method(self, name: str) -> "WBEMClassObject.MethodInfo":
        return self.get_method_nothrow(name).value

    def put_method_nothrow(self, name: str, insig: "WBEMClassObject", outsig: "WBEMClassObject") -> ComResult[None]:
        try:
            return cr(self.__o.GetMethod(name, 0, insig.wrapped_obj, outsig.wrapped_obj), None)
        except OSError:
            raise WBEMClassObjectGetMethodException

    def put_method(self, name: str, insig: "WBEMClassObject", outsig: "WBEMClassObject") -> None:
        try:
            return self.put_method_nothrow(name, insig, outsig).value
        except OSError:
            raise WBEMClassObjectGetMethodException

    def delete_method_nothrow(self, name: str) -> ComResult[None]:
        return cr(self.__o.DeleteMethod(name), None)

    def delete_method(self, name: str) -> None:
        try:
            return self.delete_method_nothrow(name).value
        except OSError:
            raise WBEMClassObjectGetMethodException

    class GetMethodFlag(IntEnum):
        ALL = 0
        LOCAL_ONLY = WBEM_FLAG_LOCAL_ONLY
        PROPAGATED_ONLY = WBEM_FLAG_PROPAGATED_ONLY

    def get_methodnames(self, flag1: GetMethodFlag) -> tuple[str, ...]:
        try:
            check_hresult(self.__o.BeginMethodEnumeration(int(flag1)))
        except OSError:
            raise WBEMClassObjectGetMethodException
        try:
            list = []
            while True:
                name = BSTR()
                insig = POINTER(IWbemClassObject)()
                outsig = POINTER(IWbemClassObject)()
                hr = self.__o.NextMethod(0, byref(name), byref(insig), byref(outsig))
                if hr != 0:
                    if hr == WBEM_S_NO_MORE_DATA:
                        break
                    check_hresult(hr)
                list.append(name.value)
            return tuple(list)
        finally:
            self.__o.EndMethodEnumeration()

    @property
    def methodnames_all(self) -> tuple[str, ...]:
        return self.get_methodnames(WBEMClassObject.GetMethodFlag.ALL)

    @property
    def methodnames_local(self) -> tuple[str, ...]:
        return self.get_methodnames(WBEMClassObject.GetMethodFlag.LOCAL_ONLY)

    @property
    def methodnames_propagated(self) -> tuple[str, ...]:
        return self.get_methodnames(WBEMClassObject.GetMethodFlag.PROPAGATED_ONLY)

    @dataclass
    class MethodInfo:
        inparams: "WBEMClassObject"
        outparams: "WBEMClassObject"

    def get_methoddict(self, flag1: GetMethodFlag) -> OrderedDict[str, MethodInfo]:
        try:
            check_hresult(self.__o.BeginMethodEnumeration(int(flag1)))
        except OSError:
            raise WBEMClassObjectGetMethodException
        try:
            d = OrderedDict()
            while True:
                name = BSTR()
                insig = POINTER(IWbemClassObject)()
                outsig = POINTER(IWbemClassObject)()
                hr = self.__o.NextMethod(0, byref(name), byref(insig), byref(outsig))
                if hr != 0:
                    if hr == WBEM_S_NO_MORE_DATA:
                        break
                    check_hresult(hr)
                d[name.value] = WBEMClassObject.MethodInfo(WBEMClassObject(insig), WBEMClassObject(outsig))
            return d
        finally:
            self.__o.EndMethodEnumeration()

    def get_methodqualifierset_nothrow(self, name: str) -> "ComResult[WBEMQualifierSet]":
        x = POINTER(IWbemQualifierSet)()
        return cr(self.__o.GetMethodQualifierSet(name, byref(x)), WBEMQualifierSet(x))

    def get_methodqualifierset(self, name: str) -> "WBEMQualifierSet":
        try:
            return self.get_methodqualifierset_nothrow(name).value
        except OSError:
            raise WBEMClassObjectGetMethodException

    @property
    def methodqualifiersets_all(self) -> "OrderedDict[str, WBEMQualifierSet]":
        try:
            return OrderedDict((name, self.get_methodqualifierset(name)) for name in self.methodnames_all)
        except OSError:
            raise WBEMClassObjectGetMethodException

    @property
    def methodqualifiersets_local(self) -> "OrderedDict[str, WBEMQualifierSet]":
        try:
            return OrderedDict((name, self.get_methodqualifierset(name)) for name in self.methodnames_local)
        except OSError:
            raise WBEMClassObjectGetMethodException

    @property
    def methodqualifiersets_propagated(self) -> "OrderedDict[str, WBEMQualifierSet]":
        try:
            return OrderedDict((name, self.get_methodqualifierset(name)) for name in self.methodnames_propagated)
        except OSError:
            raise WBEMClassObjectGetMethodException

    def get_methodorigin_nothrow(self, name: str) -> ComResult[str]:
        x = BSTR()
        return cr(self.__o.GetMethodOrigin(name, byref(x)), x.value)

    def get_methodorigin(self, name: str) -> str:
        try:
            return self.get_methodorigin_nothrow(name).value
        except OSError:
            raise WBEMClassObjectGetMethodException

    @property
    def methodorigins_all(self) -> OrderedDict[str, str]:
        try:
            return OrderedDict(((name, self.get_methodorigin(name)) for name in self.methodnames_all))
        except OSError:
            raise WBEMClassObjectGetMethodException

    @property
    def methodorigins_local(self) -> OrderedDict[str, str]:
        try:
            return OrderedDict(((name, self.get_methodorigin(name)) for name in self.methodnames_local))
        except OSError:
            raise WBEMClassObjectGetMethodException

    @property
    def methodorigins_propagated(self) -> OrderedDict[str, str]:
        try:
            return OrderedDict(((name, self.get_methodorigin(name)) for name in self.methodnames_propagated))
        except OSError:
            raise WBEMClassObjectGetMethodException

    #
    # システムプロパティ
    #

    @property
    def classname(self) -> str | None:
        x = self.get_nothrow("__Class")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def derivation(self) -> str | None:
        x = self.get_nothrow("__Derivation")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def dynasty(self) -> str | None:
        x = self.get_nothrow("__Dynasty")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def genus(self) -> WBEMGenus | None:
        x = self.get_nothrow("__Genus")
        if not x:
            return None
        t = x.value_unchecked.value.get_int32_or_none()
        return WBEMGenus(t) if t else None

    @property
    def is_instance(self) -> bool:
        return self.genus == WBEMGenus.INSTANCE

    @property
    def is_class(self) -> bool:
        return self.genus == WBEMGenus.CLASS

    @property
    def namespacename(self) -> str | None:
        x = self.get_nothrow("__Namespace")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def path(self) -> str | None:
        x = self.get_nothrow("__Path")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def propcount(self) -> int | None:
        x = self.get_nothrow("__Property_Count")
        if not x:
            return None
        return x.value_unchecked.value.get_int32_or_none()

    @property
    def relpath(self) -> str | None:
        x = self.get_nothrow("__Relpath")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def servername(self) -> str | None:
        x = self.get_nothrow("__Server")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()

    @property
    def superclassname(self) -> str | None:
        x = self.get_nothrow("__Superclass")
        if not x:
            return None
        return x.value_unchecked.value.get_bstr_or_none()


class WBEMQualifierSet:
    """IWbemQualifierSetインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IWbemQualifierSet)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IWbemQualifierSet)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def get_nothrow(self, name: str) -> ComResult[tuple[Variant, CimType]]:
        x1 = Variant()
        x2 = c_int32()
        return cr(self.__o.Get(name, 0, x1, x2), (x1, CimType(x2.value)))

    def get(self, name: str) -> tuple[Variant, CimType]:
        return self.get_nothrow(name).value

    def put_nothrow(self, name: str, value: Variant, type: CimType) -> ComResult[None]:
        return cr(self.__o.Get(name, value, type), None)

    def put(self, name: str, value: Variant, type: CimType) -> None:
        return self.put_nothrow(name, value, type).value

    def delete_nothrow(self, name: str) -> ComResult[None]:
        return cr(self.__o.Delete(name), None)

    def delete(self, name: str) -> None:
        return self.delete_nothrow(name).value

    @property
    def names_all_nothrow(self) -> ComResult[tuple[str, ...]]:
        x = SafeArrayPtr()
        hr = self.__o.GetNames(0, byref(x))
        if hr != 0:
            return cr(hr, ())
        return cr(hr, x.to_bstrarray())

    @property
    def names_all(self) -> tuple[str, ...]:
        return self.names_all_nothrow.value

    @property
    def names_local_nothrow(self) -> ComResult[tuple[str, ...]]:
        x = SafeArrayPtr()
        hr = self.__o.GetNames(WBEM_FLAG_LOCAL_ONLY, byref(x))
        if hr != 0:
            return cr(hr, ())
        return cr(hr, x.to_bstrarray())

    @property
    def names_local(self) -> tuple[str, ...]:
        return self.names_local_nothrow.value

    @property
    def names_propagated_nothrow(self) -> ComResult[tuple[str, ...]]:
        x = SafeArrayPtr()
        hr = self.__o.GetNames(WBEM_FLAG_PROPAGATED_ONLY, byref(x))
        if hr != 0:
            return cr(hr, ())
        return cr(hr, x.to_bstrarray())

    @property
    def names_propagated(self) -> tuple[str, ...]:
        return self.names_propagated_nothrow.value

    @dataclass
    class Qualifier:
        name: str
        value: Variant
        flavor: WBEMFlavor

    def __get_qualifiers_nothrow(self, flags: int) -> ComResult[tuple[Qualifier, ...]]:
        hr = self.__o.BeginEnumeration(flags)
        if hr != 0:
            return cr(hr, ())
        l = list[WBEMQualifierSet.Qualifier]()  # noqa E741
        while True:
            name = BSTR()
            value = Variant()
            flavor = c_int32()
            hr = self.__o.Next(0, byref(name), byref(value), byref(flavor))
            if hr != 0:
                if hr == WBEM_S_NO_MORE_DATA:
                    break
                return cr(hr, ())
            l.append(WBEMQualifierSet.Qualifier(name.value, value, WBEMFlavor(flavor.value)))
        return cr(hr, tuple(l))

    @property
    def qualifiers_all_nothrow(self) -> ComResult[tuple[Qualifier, ...]]:
        return self.__get_qualifiers_nothrow(0)

    @property
    def qualifiers_all(self) -> tuple[Qualifier, ...]:
        return self.__get_qualifiers_nothrow(0).value

    @property
    def qualifiers_local_nothrow(self) -> ComResult[tuple[Qualifier, ...]]:
        return self.__get_qualifiers_nothrow(WBEM_FLAG_LOCAL_ONLY)

    @property
    def qualifiers_local(self) -> tuple[Qualifier, ...]:
        return self.__get_qualifiers_nothrow(WBEM_FLAG_LOCAL_ONLY).value

    @property
    def qualifiers_propagated_nothrow(self) -> ComResult[tuple[Qualifier, ...]]:
        return self.__get_qualifiers_nothrow(WBEM_FLAG_PROPAGATED_ONLY)

    @property
    def qualifiers_propagated(self) -> tuple[Qualifier, ...]:
        return self.__get_qualifiers_nothrow(WBEM_FLAG_PROPAGATED_ONLY).value
