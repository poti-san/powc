"""
シェルアイテムのフォント情報を取得するユーティリティーを提供します。
"""

from typing import Iterator

from powc.core import ComResult
from powcpropsys.propkey import PropertyKey
from powcpropsys.propvariant import PropVariant

from ..knownfolderid import KnownFolderID
from ..shellitem2 import ShellItem2


class ShellFontPropertyKey:
    """フォントのシステムプロパティキー"""

    ACTIVESTATUS = PropertyKey.from_canonicalname("System.Fonts.ActiveStatus")
    CATEGORY = PropertyKey.from_canonicalname("System.Fonts.Category")
    COLLECTIONNAME = PropertyKey.from_canonicalname("System.Fonts.CollectionName")
    DESCRIPTION = PropertyKey.from_canonicalname("System.Fonts.Description")
    DESIGNEDFOR = PropertyKey.from_canonicalname("System.Fonts.DesignedFor")
    FAMILYNAME = PropertyKey.from_canonicalname("System.Fonts.FamilyName")
    FILENAMES = PropertyKey.from_canonicalname("System.Fonts.FileNames")
    FONT_EMBEDDABILITY = PropertyKey.from_canonicalname("System.Fonts.FontEmbeddability")
    STYLES = PropertyKey.from_canonicalname("System.Fonts.Styles")
    TYPE = PropertyKey.from_canonicalname("System.Fonts.Type")
    VENDORS = PropertyKey.from_canonicalname("System.Fonts.Vendors")
    VERSION = PropertyKey.from_canonicalname("System.Fonts.Version")


class ShellFontItemInfo:
    """フォントを表すShellItem2のフォントシステムプロパティ情報を取得します。"""

    __item: ShellItem2

    __slots__ = ("__item",)

    def __init__(self, item: ShellItem2):
        self.__item = item

    @property
    def item(self) -> ShellItem2:
        return self.__item

    @staticmethod
    def iter() -> "Iterator[ShellFontItemInfo]":
        font_folder = ShellItem2.create_knownfolder(KnownFolderID.FONTS)
        return (ShellFontItemInfo(font) for font in font_folder.iter_items())

    @property
    def activestatus_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.ACTIVESTATUS)

    @property
    def activestatus_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.ACTIVESTATUS)

    @property
    def category_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.CATEGORY)

    @property
    def category_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.CATEGORY)

    @property
    def collectionnname_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.COLLECTIONNAME)

    @property
    def collectionnname_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.COLLECTIONNAME)

    @property
    def description_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.DESCRIPTION)

    @property
    def description_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.DESCRIPTION)

    @property
    def designedfor_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.DESIGNEDFOR)

    @property
    def designedfor_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.DESIGNEDFOR)

    @property
    def familyname_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.FAMILYNAME)

    @property
    def familyname_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.FAMILYNAME)

    @property
    def filenames_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.FILENAMES)

    @property
    def filenames_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.FILENAMES)

    @property
    def embeddability_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.FONT_EMBEDDABILITY)

    @property
    def embeddability_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.FONT_EMBEDDABILITY)

    @property
    def styles_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.STYLES)

    @property
    def styles_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.STYLES)

    @property
    def type_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.TYPE)

    @property
    def type_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.TYPE)

    @property
    def vendors_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.VENDORS)

    @property
    def vendors_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.VENDORS)

    @property
    def version_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellFontPropertyKey.VERSION)

    @property
    def version_raw(self) -> PropVariant:
        return self.__item.get_prop(ShellFontPropertyKey.VERSION)

    # Values

    @property
    def activestatus(self) -> str | None:
        x = self.activestatus_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def category(self) -> str | None:
        x = self.category_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def collectionname(self) -> str | None:
        x = self.collectionnname_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def description(self) -> str | None:
        x = self.description_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def designedfor(self) -> str | None:
        x = self.designedfor_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def familyname(self) -> str | None:
        x = self.familyname_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def filenames(self) -> tuple[str, ...] | None:
        x = self.filenames_raw_nothrow
        return x.value_unchecked.to_strings() if x and not x.value_unchecked.is_empty else None

    @property
    def embeddability(self) -> tuple[str, ...] | None:
        x = self.embeddability_raw_nothrow
        return x.value_unchecked.to_strings() if x and not x.value_unchecked.is_empty else None

    @property
    def styles(self) -> tuple[str, ...] | None:
        x = self.styles_raw_nothrow
        return x.value_unchecked.to_strings() if x and not x.value_unchecked.is_empty else None

    @property
    def type(self) -> str | None:
        x = self.type_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def vendors(self) -> tuple[str, ...] | None:
        x = self.vendors_raw_nothrow
        return x.value_unchecked.to_strings() if x and not x.value_unchecked.is_empty else None

    @property
    def version(self) -> tuple[str, ...] | None:
        x = self.version_raw_nothrow
        return x.value_unchecked.to_strings() if x and not x.value_unchecked.is_empty else None
        return x.value_unchecked.to_strings() if x and not x.value_unchecked.is_empty else None
