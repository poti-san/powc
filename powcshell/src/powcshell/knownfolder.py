"""デスクトップやマイドキュメントのような既知のフォルダ。

主なクラスは :class:`KnownFolderManager` と :class:`KnownFolder` です。
"""

from ctypes import (
    POINTER,
    Structure,
    byref,
    c_int32,
    c_uint32,
    c_void_p,
    c_wchar_p,
    cast,
)
from dataclasses import dataclass
from enum import IntEnum, IntFlag
from typing import Any, Sequence

from comtypes import GUID, STDMETHOD, CoCreateInstance, IUnknown

from powc.core import ComResult, cotaskmem, cotaskmem_free, cr, query_interface

from .itemidlist import ItemIDList
from .shellitem import IShellItem, ShellItem
from .shellitem2 import IShellItem2, ShellItem2


class FffpMode(IntEnum):
    """FFFP_MODE"""

    EXACT_MATCH = 0
    NEAREST_PARENT_MATCH = 1


class KnownFolderCategory(IntEnum):
    """KF_CATEGORY"""

    VIRTUAL = 1
    FIXED = 2
    COMMON = 3
    PERUSER = 4


class KnownFolderDefinitionFlag(IntFlag):
    """KF_DEFINITION_FLAGS"""

    DEFAULT = 0
    LOCAL_REDIRECT_ONLY = 0x2
    ROAMABLE = 0x4
    PRECREATE = 0x8
    STREAM = 0x10
    PUBLISHEXPANDEDPATH = 0x20
    NO_REDIRECT_UI = 0x40


class KnownFolderRedirectFlag(IntFlag):
    """KF_REDIRECT_FLAGS"""

    USER_EXCLUSIVE = 0x1
    COPY_SOURCE_DACL = 0x2
    OWNER_USER = 0x4
    SET_OWNER_EXPLICIT = 0x8
    CHECK_ONLY = 0x10
    WITH_UI = 0x20
    UNPIN = 0x40
    PIN = 0x80
    COPY_CONTENTS = 0x200
    DEL_SOURCE_CONTENTS = 0x400
    EXCLUDE_ALL_KNOWN_SUBFOLDERS = 0x80


class KnownFolderRedirectionCapability(IntFlag):
    """KF_REDIRECTION_CAPABILITIES"""

    ALLOW_ALL = 0xFF
    REDIRECTABLE = 0x1
    DENY_ALL = 0xFFF00
    DENY_POLICY_REDIRECTED = 0x100
    DENY_POLICY = 0x200
    DENY_PERMISSIONS = 0x400


class KnownFolderFlag(IntFlag):
    """KNOWN_FOLDER_FLAG"""

    DEFAULT = 0x00000000
    FORCE_APP_DATA_REDIRECTION = 0x00080000
    RETURN_FILTER_REDIRECTION_TARGET = 0x00040000
    FORCE_PACKAGE_REDIRECTION = 0x00020000
    NO_PACKAGE_REDIRECTION = 0x00010000
    FORCE_APPCONTAINER_REDIRECTION = 0x00020000
    NO_APPCONTAINER_REDIRECTION = 0x00010000
    CREATE = 0x00008000
    DONT_VERIFY = 0x00004000
    DONT_UNEXPAND = 0x00002000
    NO_ALIAS = 0x00001000
    INIT = 0x00000800
    DEFAULT_PATH = 0x00000400
    NOT_PARENT_RELATIVE = 0x00000200
    SIMPLE_IDLIST = 0x00000100
    ALIAS_ONLY = 0x80000000


class _KNOWNFOLDER_DEFINITION(Structure):
    __slots__ = ()
    _fields_ = (
        ("category", c_int32),
        ("pszName", c_void_p),
        ("pszDescription", c_void_p),
        ("fidParent", GUID),
        ("pszRelativePath", c_void_p),
        ("pszParsingName", c_void_p),
        ("pszTooltip", c_void_p),
        ("pszLocalizedName", c_void_p),
        ("pszIcon", c_void_p),
        ("pszSecurity", c_void_p),
        ("dwAttributes", c_uint32),
        ("kfdFlags", c_int32),
        ("ftidType", GUID),
    )


@dataclass(frozen=True)
class KnownFolderDefinition:
    """フォルダ定義情報。KNOWNFOLDER_DEFINITION構造体のラッパーです。"""

    category: KnownFolderCategory
    name: str | None
    description: str | None
    parent_folderid: GUID
    relative_path: str | None
    parsing_name: str | None
    tooltip: str | None
    localized_name: str | None
    icon: str | None
    security: str | None
    attributes: int
    flags: KnownFolderDefinitionFlag
    foldertypeid: GUID
    """FolderTypeIDクラスのGUID定数。"""

    @staticmethod
    def empty() -> "KnownFolderDefinition":
        return KnownFolderDefinition(
            KnownFolderCategory(0),
            "",
            "",
            GUID(),
            "",
            "",
            "",
            "",
            "",
            "",
            0,
            KnownFolderDefinitionFlag(0),
            GUID(),
        )


class IKnownFolder(IUnknown):
    _iid_ = GUID("{3AA7AF7E-9B36-420c-A8E3-F77D4674A488}")
    _methods_ = [
        STDMETHOD(c_int32, "GetId", (POINTER(GUID),)),
        STDMETHOD(c_int32, "GetCategory", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetShellItem", (c_uint32, POINTER(GUID), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "GetPath", (c_uint32, POINTER(c_wchar_p))),
        STDMETHOD(c_int32, "SetPath", (c_uint32, c_wchar_p)),
        STDMETHOD(c_int32, "GetIDList", (c_uint32, POINTER(c_void_p))),
        STDMETHOD(c_int32, "GetFolderType", (POINTER(GUID),)),
        STDMETHOD(c_int32, "GetRedirectionCapabilities", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetFolderDefinition", (POINTER(_KNOWNFOLDER_DEFINITION),)),
    ]
    __slots__ = ()


class KnownFolder:
    """既知フォルダ。IKnownFolderインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IKnownFolder)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IKnownFolder)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def id_nothrow(self) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.GetId(byref(x)), x)

    @property
    def id(self) -> GUID:
        return self.id_nothrow.value

    @property
    def category_nothrow(self) -> ComResult[KnownFolderCategory]:
        x = c_int32()
        return cr(self.__o.GetCategory(byref(x)), KnownFolderCategory(x.value))

    @property
    def category(self) -> KnownFolderCategory:
        return self.category_nothrow.value

    def get_shellitem_nothrow(self, flags: KnownFolderFlag) -> ComResult[ShellItem2]:
        x = POINTER(IShellItem2)()
        return cr(self.__o.GetShellItem(int(flags), IShellItem2._iid_, byref(x)), ShellItem2(x))

    def get_shellitem(self, flags: KnownFolderFlag) -> ShellItem2:
        return self.get_shellitem_nothrow(flags).value

    def get_shellitem_v2_nothrow(self, flags: KnownFolderFlag) -> ComResult[ShellItem]:
        x = POINTER(IShellItem)()
        return cr(self.__o.GetShellItem(int(flags), IShellItem._iid_, byref(x)), ShellItem(x))

    def get_shellitem_v2(self, flags: KnownFolderFlag) -> ShellItem:
        return self.get_shellitem_v2_nothrow(flags).value

    def get_path_nothrow(self, flags: KnownFolderFlag) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetPath(int(flags), byref(p)), p.value or "")

    def get_path(self, flags: KnownFolderFlag) -> str:
        return self.get_path_nothrow(flags).value

    def set_path_nothrow(self, flags: KnownFolderFlag, path: str) -> ComResult[None]:
        return cr(self.__o.SetPath(int(flags), path), None)

    def set_path(self, flags: KnownFolderFlag, path: str) -> None:
        return self.set_path_nothrow(flags, path).value

    def get_idlist_nothrow(self, flags: KnownFolderFlag) -> ComResult[ItemIDList]:
        x = ItemIDList()
        return cr(self.__o.GetIDList(int(flags), byref(x)), x)

    def get_idlist(self, flags: KnownFolderFlag) -> ItemIDList:
        return self.get_idlist_nothrow(flags).value

    @property
    def foldertype_nothrow(self) -> ComResult[GUID]:
        """フォルダの種類（FolderTypeID定数）を取得します。"""
        x = GUID()
        return cr(self.__o.GetFolderType(byref(x)), x)

    @property
    def foldertype(self) -> GUID:
        """フォルダの種類（FolderTypeID定数）を取得します。"""
        return self.foldertype_nothrow.value

    @property
    def redirection_capabilitiy_nothrow(self) -> ComResult[KnownFolderRedirectionCapability]:
        x = c_int32()
        return cr(self.__o.GetRedirectionCapabilities(byref(x)), KnownFolderRedirectionCapability(x.value))

    @property
    def redirection_capability(self) -> KnownFolderRedirectionCapability:
        return self.redirection_capabilitiy_nothrow.value

    @property
    def folder_definition_nothrow(self) -> ComResult[KnownFolderDefinition]:
        x = _KNOWNFOLDER_DEFINITION()
        hr = self.__o.GetFolderDefinition(byref(x))
        try:
            if hr < 0:
                return cr(
                    hr,
                    KnownFolderDefinition.empty(),
                )
            return cr(
                hr,
                KnownFolderDefinition(
                    KnownFolderCategory(x.category),
                    cast(x.pszName, c_wchar_p).value,
                    cast(x.pszDescription, c_wchar_p).value,
                    x.fidParent,
                    cast(x.pszRelativePath, c_wchar_p).value,
                    cast(x.pszParsingName, c_wchar_p).value,
                    cast(x.pszTooltip, c_wchar_p).value,
                    cast(x.pszLocalizedName, c_wchar_p).value,
                    cast(x.pszIcon, c_wchar_p).value,
                    cast(x.pszSecurity, c_wchar_p).value,
                    x.dwAttributes,
                    KnownFolderDefinitionFlag(x.kfdFlags),
                    x.ftidType,
                ),
            )
        finally:
            cotaskmem_free(x.pszName)
            cotaskmem_free(x.pszDescription)
            cotaskmem_free(x.pszRelativePath)
            cotaskmem_free(x.pszParsingName)
            cotaskmem_free(x.pszTooltip)
            cotaskmem_free(x.pszLocalizedName)
            cotaskmem_free(x.pszIcon)
            cotaskmem_free(x.pszSecurity)

    @property
    def folder_definition(self) -> KnownFolderDefinition:
        return self.folder_definition_nothrow.value


class IKnownFolderManager(IUnknown):
    """"""

    _iid_ = GUID("{8BE2D872-86AA-4d47-B776-32CCA40C7018}")
    _methods_ = [
        STDMETHOD(c_int32, "FolderIdFromCsidl", (c_int32, POINTER(GUID))),
        STDMETHOD(c_int32, "FolderIdToCsidl", (POINTER(GUID), POINTER(c_int32))),
        STDMETHOD(c_int32, "GetFolderIds", (POINTER(c_void_p), POINTER(c_uint32))),
        STDMETHOD(c_int32, "GetFolder", (POINTER(GUID), POINTER(POINTER(IKnownFolder)))),
        STDMETHOD(c_int32, "GetFolderByName", (c_wchar_p, POINTER(POINTER(IKnownFolder)))),
        STDMETHOD(c_int32, "RegisterFolder", (POINTER(GUID), POINTER(_KNOWNFOLDER_DEFINITION))),
        STDMETHOD(c_int32, "UnregisterFolder", (POINTER(GUID),)),
        STDMETHOD(c_int32, "FindFolderFromPath", (c_wchar_p, c_int32, POINTER(POINTER(IKnownFolder)))),
        STDMETHOD(c_int32, "FindFolderFromIDList", (c_void_p, POINTER(POINTER(IKnownFolder)))),
        STDMETHOD(
            c_int32,
            "Redirect",
            (POINTER(GUID), c_void_p, c_int32, c_wchar_p, c_uint32, POINTER(GUID), POINTER(c_wchar_p)),
        ),
    ]
    __slots__ = ()


class KnownFolderManager:
    """既知フォルダマネージャ。IKnownFolderManagerインターフェイスのラッパーです。

    Examples:
        >>> kfmanager = KnownFolderManager.create()
    """

    __o: Any  # POINTER(IKnownFolderManager)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IKnownFolderManager)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "KnownFolderManager":
        CLSID_KnownFolderManager = GUID("{4df0c730-df9d-4ae3-9153-aa6b82e9795a}")
        return KnownFolderManager(CoCreateInstance(CLSID_KnownFolderManager, IKnownFolderManager))

    def folderid_from_csidl_nothrow(self, csidl: int) -> ComResult[GUID]:
        x = GUID()
        return cr(self.__o.FolderIdFromCsidl(csidl, byref(x)), x)

    def folderid_from_csidl(self, csidl: int) -> GUID:
        return self.folderid_from_csidl_nothrow(csidl).value

    def folderid_to_csidl_nothrow(self, folderid: GUID) -> ComResult[int]:
        x = c_int32()
        return cr(self.__o.FolderIdToCsidl(folderid, byref(x)), x.value)

    def folderid_to_csidl(self, folderid: GUID) -> int:
        return self.folderid_to_csidl_nothrow(folderid).value

    @property
    def folderids_nothrow(self) -> ComResult[tuple[GUID]]:
        with cotaskmem(c_void_p()) as p:
            c = c_uint32()
            hr = self.__o.GetFolderIds(byref(p), byref(c))
            if hr < 0:
                return cr(hr, tuple())
            return cr(hr, tuple((GUID * c.value).from_address(p.value or 0)))

    @property
    def folderids(self) -> tuple[GUID]:
        return self.folderids_nothrow.value

    def get_folder_nothrow(self, folderid: GUID) -> ComResult[KnownFolder]:
        p = POINTER(IKnownFolder)()
        return cr(self.__o.GetFolder(byref(folderid), byref(p)), KnownFolder(p))

    def get_folder(self, folderid: GUID) -> KnownFolder:
        return self.get_folder_nothrow(folderid).value

    def get_folder_byname_nothrow(self, name: str) -> ComResult[KnownFolder]:
        p = POINTER(IKnownFolder)()
        return cr(self.__o.GetFolderByName(name, byref(p)), KnownFolder(p))

    def get_folder_byname(self, name: str) -> KnownFolder:
        return self.get_folder_byname_nothrow(name).value

    def register_folder_nothrow(self, folderid: GUID, definition: KnownFolderDefinition) -> ComResult[None]:
        kf = _KNOWNFOLDER_DEFINITION()
        kf.category = definition.category
        kf.pszName = definition.name
        kf.pszDescription = definition.description
        kf.fidParent = definition.parent_folderid
        kf.pszRelativePath = definition.relative_path
        kf.pszParsingName = definition.parsing_name
        kf.pszTooltip = definition.tooltip
        kf.pszLocalizedName = definition.localized_name
        kf.pszIcon = definition.icon
        kf.pszSecurity = definition.security
        kf.dwAttributes = int(definition.attributes)
        kf.kfdFlags = int(definition.flags)
        kf.ftidType = definition.foldertypeid.value
        return cr(self.__o.RegisterFolder(byref(folderid), byref(kf)), None)

    def register_folder(self, folderid: GUID, definition: KnownFolderDefinition) -> None:
        return self.register_folder_nothrow(folderid, definition).value

    def unregister_folder_nothrow(self, folderid: GUID) -> ComResult[None]:
        return cr(self.__o.UnregisterFolder(byref(folderid)), None)

    def unregister_folder(self, folderid: GUID) -> None:
        return self.unregister_folder_nothrow(folderid).value

    def find_folder_from_path_nothrow(self, path: str, mode: FffpMode) -> ComResult[KnownFolder]:
        p = POINTER(IKnownFolder)()
        return cr(self.__o.FindFolderFromPath(path, int(mode), byref(p)), KnownFolder(p))

    def find_folder_from_path(self, path: str, mode: FffpMode) -> KnownFolder:
        return self.find_folder_from_path_nothrow(path, mode).value

    def find_folder_from_idlist_nothrow(self, pidl: c_void_p) -> ComResult[KnownFolder]:
        p = POINTER(IKnownFolder)()
        return cr(self.__o.FindFolderFromIDList(pidl, byref(p)), KnownFolder(p))

    def find_folder_from_idlist(self, pidl: c_void_p) -> KnownFolder:
        return self.find_folder_from_idlist_nothrow(pidl).value

    def redirect_nothrow(
        self,
        folderid: GUID,
        flags: KnownFolderRedirectFlag,
        target_path: str | None = None,
        exclusions: Sequence[GUID] | None = None,
        window_handle: int = 0,
    ) -> ComResult[str]:

        if exclusions:
            pexclusions = (GUID * len(exclusions))()
            for i in range(len(exclusions)):
                pexclusions[i] = exclusions[i]
        else:
            pexclusions = None

        with cotaskmem(c_wchar_p()) as sp:
            return cr(
                self.__o.Redirect(
                    folderid,
                    window_handle,
                    int(flags),
                    target_path,
                    len(exclusions) if exclusions else 0,
                    pexclusions if exclusions else None,
                    byref(sp),
                ),
                sp.value or "",
            )

    def redirect(
        self,
        folderid: GUID,
        flags: KnownFolderRedirectFlag,
        target_path: str | None = None,
        exclusions: Sequence[GUID] | None = None,
        window_handle: int = 0,
    ) -> str:

        return self.redirect_nothrow(folderid, flags, target_path, exclusions, window_handle).value
