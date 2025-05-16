"""
COMカテゴリオブジェクト。
COMカテゴリ情報の取得 :class:`CategoryInformation` や登録・解除 :class:`CategoryRegister` 機能を提供します。
"""

from ctypes import (
    POINTER,
    Structure,
    byref,
    c_int32,
    c_uint32,
    c_void_p,
    c_wchar,
    c_wchar_p,
)
from typing import TYPE_CHECKING, Any, Iterator, Sequence

from comtypes import GUID, STDMETHOD, CoCreateInstance, IUnknown

from .core import ComResult, check_hresult, cotaskmem, cr, query_interface


class IEnumGUID(IUnknown):
    """IEnumGUIDインターフェイス"""

    __slots__ = ()
    _iid_ = GUID("{0002E000-0000-0000-C000-000000000046}")


IEnumGUID._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(GUID), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(IEnumGUID),)),
]


class GuidEnumerator:
    """IEnumGUIDインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IEnumGUID)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumGUID)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __iter__(self) -> Iterator[GUID]:
        check_hresult(self.__o.Reset())
        hr = 0
        x = GUID()
        while (hr := self.__o.Next(1, byref(x), None)) == 0:
            yield x
        check_hresult(hr)

    def clone_nothrow(self) -> "ComResult[GuidEnumerator]":
        x = POINTER(IEnumGUID)()
        return cr(self.__o.Clone(byref(x)), GuidEnumerator(x))

    def clone(self) -> "GuidEnumerator":
        return self.clone_nothrow().value


class CategoryInfo(Structure):
    __slots__ = ()
    _fields_ = (("catid", GUID), ("lcid", c_uint32), ("description", c_wchar * 128))
    if TYPE_CHECKING:
        catid: GUID
        """カテゴリのGUID。"""
        lcid: int
        """カテゴリのロケールID。"""
        description: str
        """カテゴリの概要。"""


class IEnumCATEGORYINFO(IUnknown):
    """IEnumCATEGORYINFOインターフェイス"""

    __slots__ = ()
    _iid_ = GUID("{0002E011-0000-0000-C000-000000000046}")


IEnumCATEGORYINFO._methods_ = [
    STDMETHOD(c_int32, "Next", (c_uint32, POINTER(CategoryInfo), POINTER(c_uint32))),
    STDMETHOD(c_int32, "Skip", (c_uint32,)),
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Clone", (POINTER(IEnumCATEGORYINFO),)),
]


class CategoryInfoEnumerator:
    """IEnumCATEGORYINFOインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IEnumCATEGORYINFO)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IEnumCATEGORYINFO)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __iter__(self) -> Iterator[CategoryInfo]:
        check_hresult(self.__o.Reset())
        hr = 0
        x = CategoryInfo()
        while (hr := self.__o.Next(1, byref(x), None)) == 0:
            yield x
        check_hresult(hr)

    def clone_nothrow(self) -> "ComResult[CategoryInfoEnumerator]":
        x = POINTER(IEnumCATEGORYINFO)()
        return cr(self.__o.Clone(byref(x)), CategoryInfoEnumerator(x))

    def clone(self) -> "CategoryInfoEnumerator":
        return self.clone_nothrow().value


class ICatRegister(IUnknown):
    """ICatRegisterインターフェイス"""

    _iid_ = GUID("{0002E012-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "RegisterCategories", (c_uint32, POINTER(CategoryInfo))),
        STDMETHOD(c_int32, "UnRegisterCategories", (c_uint32, POINTER(GUID))),
        STDMETHOD(c_int32, "RegisterClassImplCategories", (POINTER(GUID), c_uint32, POINTER(GUID))),
        STDMETHOD(c_int32, "UnRegisterClassImplCategories", (POINTER(GUID), c_uint32, POINTER(GUID))),
        STDMETHOD(c_int32, "RegisterClassReqCategories", (POINTER(GUID), c_uint32, POINTER(GUID))),
        STDMETHOD(c_int32, "UnRegisterClassReqCategories", (POINTER(GUID), c_uint32, POINTER(GUID))),
    ]


class CategoryRegister:
    """ICatRegisterインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(ICatRegister)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, ICatRegister)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "CategoryRegister":
        CLSID_StdComponentCategoriesMgr = GUID("{0002E005-0000-0000-C000-000000000046}")
        return CategoryRegister(CoCreateInstance(CLSID_StdComponentCategoriesMgr, ICatRegister))

    def register_categories_nothrow(self, catinfos: Sequence[CategoryInfo]) -> ComResult[None]:
        x = (CategoryInfo * len(catinfos))(*catinfos)
        return cr(self.__o.RegisterCategories(len(catinfos), x), None)

    def register_categories(self, catinfos: Sequence[CategoryInfo]) -> None:
        return self.register_categories_nothrow(catinfos).value

    def unregister_categories_nothrow(self, catinfos: Sequence[CategoryInfo]) -> ComResult[None]:
        x = (CategoryInfo * len(catinfos))(*catinfos)
        return cr(self.__o.UnRegisterCategories(len(catinfos), x), None)

    def unregister_categories(self, catinfos: Sequence[CategoryInfo]) -> None:
        return self.unregister_categories_nothrow(catinfos).value

    def register_classimplcategories_nothrow(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> ComResult[None]:
        x = (CategoryInfo * len(catinfos))(*catinfos)
        return cr(self.__o.RegisterClassImplCategories(clsid, len(catinfos), x), None)

    def register_classimplcategories(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> None:
        return self.register_classimplcategories_nothrow(clsid, catinfos).value

    def unregister_classimplcategories_nothrow(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> ComResult[None]:
        x = (CategoryInfo * len(catinfos))(*catinfos)
        return cr(self.__o.UnRegisterClassImplCategories(clsid, len(catinfos), x), None)

    def unregister_classimplcategories(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> None:
        return self.unregister_classimplcategories_nothrow(clsid, catinfos).value

    def register_classreqcategories_nothrow(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> ComResult[None]:
        x = (CategoryInfo * len(catinfos))(catinfos)
        return cr(self.__o.RegisterClassReqCategories(clsid, len(catinfos), x), None)

    def register_classreqcategories(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> None:
        return self.register_classimplcategories_nothrow(clsid, catinfos).value

    def unregister_classreqcategories_nothrow(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> ComResult[None]:
        x = (CategoryInfo * len(catinfos))(catinfos)
        return cr(self.__o.UnRegisterClassReqCategories(clsid, len(catinfos), x), None)

    def unregister_classreqcategories(self, clsid: GUID, catinfos: Sequence[CategoryInfo]) -> None:
        return self.unregister_classimplcategories_nothrow(clsid, catinfos).value


class ICatInformation(IUnknown):
    """ICatInformationインターフェイス"""

    __slots__ = ()
    _iid_ = GUID("{0002E013-0000-0000-C000-000000000046}")
    _methods_ = [
        STDMETHOD(c_int32, "EnumCategories", (c_uint32, POINTER(POINTER(IEnumCATEGORYINFO)))),
        STDMETHOD(c_int32, "GetCategoryDesc", (POINTER(GUID), c_uint32, POINTER(c_wchar_p))),
        STDMETHOD(
            c_int32,
            "EnumClassesOfCategories",
            (c_uint32, POINTER(GUID), c_uint32, POINTER(GUID), POINTER(POINTER(IEnumGUID))),
        ),
        STDMETHOD(c_int32, "IsClassOfCategories", (POINTER(GUID), c_uint32, POINTER(GUID), c_uint32, POINTER(GUID))),
        STDMETHOD(c_int32, "EnumImplCategoriesOfClass", (POINTER(GUID), POINTER(POINTER(IEnumGUID)))),
        STDMETHOD(c_int32, "EnumReqCategoriesOfClass", (POINTER(GUID), POINTER(POINTER(IEnumGUID)))),
    ]


class CategoryInformation:
    """ICatInformationインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(ICatInformation)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, ICatInformation)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "CategoryInformation":
        CLSID_StdComponentCategoriesMgr = GUID("{0002E005-0000-0000-C000-000000000046}")
        return CategoryInformation(CoCreateInstance(CLSID_StdComponentCategoriesMgr, ICatInformation))

    def get_enumcategories_nothrow(self, lcid: int = 0) -> ComResult[CategoryInfoEnumerator]:
        x = POINTER(IEnumCATEGORYINFO)()
        return cr(self.__o.EnumCategories(lcid, byref(x)), CategoryInfoEnumerator(x))

    def get_enumcategories(self, lcid: int = 0) -> CategoryInfoEnumerator:
        return self.get_enumcategories_nothrow(lcid).value

    def get_categorydesc_nothrow(self, catid: GUID, lcid: int = 0) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetCategoryDesc(catid, lcid, byref(p)), p.value or "")

    def get_categorydesc(self, catid: GUID, lcid: int = 0) -> str:
        return self.get_categorydesc_nothrow(catid, lcid).value

    def get_enumclassesofcategories_nothrow(
        self, impl_categoryids: Sequence[GUID] | None, req_caterogyids: Sequence[GUID] | None
    ) -> ComResult[GuidEnumerator]:
        x = POINTER(IEnumGUID)()
        impl_catids = (GUID * len(impl_categoryids))(impl_categoryids) if impl_categoryids else None
        req_catids = (GUID * len(req_caterogyids))(req_caterogyids) if req_caterogyids else None
        return cr(
            self.__o.EnumClassesOfCategories(
                len(impl_catids) if impl_catids else -1,
                impl_catids,
                len(req_catids) if req_catids else 0,
                req_catids,
                byref(x),
            ),
            GuidEnumerator(x),
        )

    def get_enumclassesofcategories(
        self, impl_categoryids: Sequence[GUID] | None, req_caterogyids: Sequence[GUID] | None
    ) -> GuidEnumerator:
        return self.get_enumclassesofcategories_nothrow(impl_categoryids, req_caterogyids).value

    def is_classofcategories_nothrow(
        self, clsid: GUID, impl_categoryids: Sequence[GUID] | None, req_caterogyids: Sequence[GUID] | None
    ) -> ComResult[bool]:
        impl_catids = (GUID * len(impl_categoryids))(impl_categoryids) if impl_categoryids else None
        req_catids = (GUID * len(req_caterogyids))(req_caterogyids) if req_caterogyids else None
        hr = self.__o.IsClassOfCategories(
            clsid,
            len(impl_catids) if impl_catids else -1,
            impl_catids,
            len(req_catids) if req_catids else 0,
            req_catids,
        )
        return cr(hr, hr == 0)

    def is_classofcategories(
        self, clsid: GUID, impl_categoryids: Sequence[GUID] | None, req_caterogyids: Sequence[GUID] | None
    ) -> bool:
        return self.is_classofcategories_nothrow(clsid, impl_categoryids, req_caterogyids).value

    def enum_implcategoriesofclass(self, clsid: GUID) -> ComResult[GuidEnumerator]:
        x = POINTER(IEnumGUID)()
        return cr(self.__o.EnumImplCategoriesOfClass(clsid, byref(x)), GuidEnumerator(x))

    def enum_reqcategoriesofclass(self, clsid: GUID) -> ComResult[GuidEnumerator]:
        x = POINTER(IEnumGUID)()
        return cr(self.__o.EnumReqCategoriesOfClass(clsid, byref(x)), GuidEnumerator(x))
