"""
ごみ箱内のシェルアイテム情報を取得するユーティリティーを提供します。
"""

from datetime import datetime

from powc.core import ComResult
from powcpropsys.propkey import PropertyKey
from powcpropsys.propvariant import PropVariant

from ..knownfolderid import KnownFolderID
from ..shellitem2 import ShellItem2


class ShellRecycleBinItemPropertyKey:
    """ごみ箱内項目のシステムプロパティキー"""

    DATE_DELETED = PropertyKey.from_canonicalname("System.Recycle.DateDeleted")
    DELETEDFROM = PropertyKey.from_canonicalname("System.Recycle.DeletedFrom")


class ShellRecycleBinItemInfo:
    """ごみ箱内項目を表すShellItem2の特殊システムプロパティ情報を取得します。"""

    __item: ShellItem2

    __slots__ = ("__item",)

    def __init__(self, item: ShellItem2):
        self.__item = item

    @property
    def item(self) -> ShellItem2:
        return self.__item

    @staticmethod
    def create() -> "ShellRecycleBinItemInfo":
        return ShellRecycleBinItemInfo(ShellItem2.create_knownfolder(KnownFolderID.RECYCLE_BIN_FOLDER))

    @property
    def date_deleted_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellRecycleBinItemPropertyKey.DATE_DELETED)

    @property
    def date_deleted_raw(self) -> PropVariant:
        return self.date_deleted_raw_nothrow.value

    @property
    def date_deleted(self) -> datetime | None:
        x = self.date_deleted_raw_nothrow
        return x.value_unchecked.get_filetime() if x and not x.value_unchecked.is_empty else None

    @property
    def deletedfrom_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellRecycleBinItemPropertyKey.DELETEDFROM)

    @property
    def deletedfrom_raw(self) -> PropVariant:
        return self.deletedfrom_raw_nothrow.value

    @property
    def deletedfrom(self) -> str | None:
        x = self.deletedfrom_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None
