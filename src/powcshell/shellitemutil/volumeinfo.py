"""
ボリュームのシェルアイテム情報を取得するユーティリティーを提供します。
"""

from typing import Iterator

from powc.core import ComResult
from powcpropsys.propkey import PropertyKey
from powcpropsys.propvariant import PropVariant

from ..knownfolderid import KnownFolderID
from ..shellitem2 import ShellItem2


class ShellVolumePropertyKey:
    """ボリュームのシステムプロパティキー"""

    BITLOCKER_CANCHANGE_PASSPHRAGE_BYPROXY = PropertyKey.from_canonicalname(
        "System.Volume.BitLockerCanChangePassphraseByProxy"
    )
    BITLOCKER_CANCHANGE_PIN = PropertyKey.from_canonicalname("System.Volume.BitLockerCanChangePin")
    BITLOCKER_REQUIRES_ADMIN = PropertyKey.from_canonicalname("System.Volume.BitLockerRequiresAdmin")
    FILESYSTEM = PropertyKey.from_canonicalname("System.Volume.FileSystem")
    IS_MAPPED_DRIVE = PropertyKey.from_canonicalname("System.Volume.IsMappedDrive")
    IS_ROOT = PropertyKey.from_canonicalname("System.Volume.IsRoot")
    CAPACITY = PropertyKey.from_canonicalname("System.Capacity")


class ShellVolumeInfo:
    """ボリュームを表すShellItem2の特殊システムプロパティ情報を取得します。"""

    __item: ShellItem2

    __slots__ = ("__item",)

    def __init__(self, item: ShellItem2):
        self.__item = item

    @property
    def item(self) -> ShellItem2:
        return self.__item

    @staticmethod
    def iter() -> "Iterator[ShellVolumeInfo]":
        return (ShellVolumeInfo(item) for item in ShellItem2.create_knownfolder(KnownFolderID.COMPUTER_FOLDER).items)

    @property
    def bitlocker_canchange_passphrage_byproxy_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.BITLOCKER_CANCHANGE_PASSPHRAGE_BYPROXY)

    @property
    def bitlocker_canchange_passphrage_byproxy_raw(self) -> PropVariant:
        return self.bitlocker_canchange_passphrage_byproxy_raw_nothrow.value

    @property
    def bitlocker_canchange_passphrage_byproxy(self) -> bool | None:
        x = self.bitlocker_canchange_passphrage_byproxy_raw_nothrow
        return x.value_unchecked.get_bool() if x and not x.value_unchecked.is_empty else None

    @property
    def bitlocker_canchange_pin_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.BITLOCKER_CANCHANGE_PIN)

    @property
    def bitlocker_canchange_pin_raw(self) -> PropVariant:
        return self.bitlocker_canchange_pin_raw_nothrow.value

    @property
    def bitlocker_canchange_pin(self) -> bool | None:
        x = self.bitlocker_canchange_pin_raw_nothrow
        return x.value_unchecked.get_bool() if x and not x.value_unchecked.is_empty else None

    @property
    def bitlocker_requires_admin_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.BITLOCKER_REQUIRES_ADMIN)

    @property
    def bitlocker_requires_admin_raw(self) -> PropVariant:
        return self.bitlocker_requires_admin_raw_nothrow.value

    @property
    def bitlocker_requires_admin(self) -> bool | None:
        x = self.bitlocker_requires_admin_raw_nothrow
        return x.value_unchecked.get_bool() if x and not x.value_unchecked.is_empty else None

    @property
    def filesystem_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.FILESYSTEM)

    @property
    def filesystem_raw(self) -> PropVariant:
        return self.filesystem_raw_nothrow.value

    @property
    def filesystem(self) -> str | None:
        x = self.filesystem_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def is_mapped_drive_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.IS_MAPPED_DRIVE)

    @property
    def is_mapped_drive_raw(self) -> PropVariant:
        return self.is_mapped_drive_raw_nothrow.value

    @property
    def is_mapped_drive(self) -> bool | None:
        x = self.is_mapped_drive_raw_nothrow
        return x.value_unchecked.get_bool() if x and not x.value_unchecked.is_empty else None

    @property
    def is_root_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.IS_ROOT)

    @property
    def is_root_raw(self) -> PropVariant:
        return self.is_root_raw_nothrow.value

    @property
    def is_root(self) -> bool | None:
        x = self.is_root_raw_nothrow
        return x.value_unchecked.get_bool() if x and not x.value_unchecked.is_empty else None

    @property
    def capacity_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellVolumePropertyKey.CAPACITY)

    @property
    def capacity_raw(self) -> PropVariant:
        return self.capacity_raw_nothrow.value

    @property
    def capacity(self) -> int | None:
        x = self.capacity_raw_nothrow
        return x.value_unchecked.get_uint64() if x and not x.value_unchecked.is_empty else None
