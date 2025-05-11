"""
コンピューターのシェルアイテム情報を取得するユーティリティーを提供します。
"""

from powc.core import ComResult
from powcpropsys.propkey import PropertyKey
from powcpropsys.propvariant import PropVariant

from ..knownfolderid import KnownFolderID
from ..shellitem2 import ShellItem2


class ShellComputerFolderPropertyKey:
    """コンピューターのシステムプロパティキー"""

    MEMORY = PropertyKey.from_canonicalname("System.Computer.Memory")
    PROCESSOR = PropertyKey.from_canonicalname("System.Computer.Processor")
    SIMPLENAME = PropertyKey.from_canonicalname("System.Computer.SimpleName")
    WORKGROUP = PropertyKey.from_canonicalname("System.Computer.Workgroup")


class ShellComputerFolderInfo:
    """コンピューターを表すShellItem2の特殊システムプロパティ情報を取得します。"""

    __item: ShellItem2

    __slots__ = ("__item",)

    def __init__(self, item: ShellItem2):
        self.__item = item

    @property
    def item(self) -> ShellItem2:
        return self.__item

    @staticmethod
    def create() -> "ShellComputerFolderInfo":
        return ShellComputerFolderInfo(ShellItem2.create_knownfolder(KnownFolderID.COMPUTER_FOLDER))

    @property
    def memory_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellComputerFolderPropertyKey.MEMORY)

    @property
    def memory_raw(self) -> PropVariant:
        return self.memory_raw_nothrow.value

    @property
    def memory(self) -> str | None:
        x = self.memory_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def processor_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellComputerFolderPropertyKey.PROCESSOR)

    @property
    def processor_raw(self) -> PropVariant:
        return self.processor_raw_nothrow.value

    @property
    def processor(self) -> str | None:
        x = self.processor_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def simplename_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellComputerFolderPropertyKey.SIMPLENAME)

    @property
    def simplename_raw(self) -> PropVariant:
        return self.simplename_raw_nothrow.value

    @property
    def simplename(self) -> str | None:
        x = self.simplename_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None

    @property
    def workgroup_raw_nothrow(self) -> ComResult[PropVariant]:
        return self.__item.get_prop_nothrow(ShellComputerFolderPropertyKey.WORKGROUP)

    @property
    def workgroup_raw(self) -> PropVariant:
        return self.workgroup_raw_nothrow.value

    @property
    def workgroup(self) -> str | None:
        x = self.workgroup_raw_nothrow
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None
        return x.value_unchecked.get_wstr() if x and not x.value_unchecked.is_empty else None
