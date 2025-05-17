"""シェル項目ファイル操作の便利機能。

主なクラスは :class:`ShellFileOperationProgressSinkForCall` です。"""

from inspect import signature as inspect_sig
from pprint import pformat
from typing import Any, Callable, OrderedDict, TextIO, override

from comtypes import hresult

from .shellfileop import ShellFileOperationProgressSinkBase
from .shellitem2 import ShellItem2


class ShellFileOperationProgressSinkForCall(ShellFileOperationProgressSinkBase):
    """ShellFileOperationのコールクラス。

    メソッドの呼び出しを引数として構築時に与えられた関数を呼び出します。
    """

    __f: Callable[[str, OrderedDict[str, Any]], None]

    __slots__ = ("__f",)

    def __init__(self, f: Callable[[str, OrderedDict[str, Any]], None]) -> None:
        self.__f = f
        super().__init__()

    @override
    def start_operations(self) -> int:
        self.__f(
            "StartOperations",
            inspect_sig(ShellFileOperationProgressSinkForCall.start_operations).bind(self).arguments,
        )
        return hresult.S_OK

    @override
    def finish_operations(self, hr_result: int) -> int:
        self.__f(
            "FinishOperations",
            inspect_sig(ShellFileOperationProgressSinkForCall.finish_operations).bind(self, hr_result).arguments,
        )
        return hresult.S_OK

    @override
    def pre_renameitem(self, flags: int, item: ShellItem2, newname: str) -> int:
        self.__f(
            "PreRenameItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.pre_renameitem)
            .bind(self, flags, item, newname)
            .arguments,
        )
        return hresult.S_OK

    @override
    def post_renameitem(
        self,
        flags: int,
        item: ShellItem2,
        newname: str,
        hr_rename: int,
        newly_created_item: ShellItem2,
    ) -> int:
        self.__f(
            "PostRenameItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.post_renameitem)
            .bind(self, flags, item, newname, hr_rename, newly_created_item)
            .arguments,
        )
        return hresult.S_OK

    @override
    def pre_moveitem(self, flags: int, item: ShellItem2, dest_folder: ShellItem2, newname: str) -> int:
        self.__f(
            "PreMoveItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.pre_moveitem)
            .bind(self, flags, item, dest_folder, newname)
            .arguments,
        )
        return hresult.S_OK

    @override
    def post_moveitem(
        self,
        flags: int,
        item: ShellItem2,
        dest_folder: ShellItem2,
        newname: str,
        hr_move: int,
        newly_created_item: ShellItem2,
    ) -> int:
        self.__f(
            "PostMoveItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.post_moveitem)
            .bind(self, flags, item, dest_folder, newname, hr_move, newly_created_item)
            .arguments,
        )
        return hresult.S_OK

    @override
    def pre_copyitem(self, flags: int, item: ShellItem2, dest_folder: ShellItem2, newname: str) -> int:
        self.__f(
            "PreCopyItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.pre_copyitem)
            .bind(self, flags, item, dest_folder, newname)
            .arguments,
        )
        return hresult.S_OK

    @override
    def post_copyitem(
        self,
        flags: int,
        item: ShellItem2,
        dest_folder: ShellItem2,
        newname: str,
        hr_copy: int,
        newly_created_item: ShellItem2,
    ) -> int:
        self.__f(
            "PostCopyItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.post_copyitem)
            .bind(self, flags, item, dest_folder, newname, hr_copy, newly_created_item)
            .arguments,
        )
        return hresult.S_OK

    @override
    def pre_deleteitem(self, flags: int, item: ShellItem2) -> int:
        self.__f(
            "PreDeleteItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.pre_deleteitem).bind(self, flags, item).arguments,
        )
        return hresult.S_OK

    @override
    def post_deleteitem(
        self,
        flags: int,
        item: ShellItem2,
        hr_delete: int,
        newly_created_item: ShellItem2,
    ) -> int:
        self.__f(
            "PostDeleteItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.post_deleteitem)
            .bind(self, flags, item, hr_delete, newly_created_item)
            .arguments,
        )
        return hresult.S_OK

    @override
    def pre_newitem(self, flags: int, item: ShellItem2, newname: str) -> int:
        self.__f(
            "PreNewItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.pre_newitem).bind(self, flags, item, newname).arguments,
        )
        return hresult.S_OK

    @override
    def post_newitem(
        self,
        flags: int,
        dest_folder: ShellItem2,
        newname: str,
        templatename: str,
        file_attrs: int,
        hr_new: int,
        newly_created_item: ShellItem2,
    ) -> int:
        self.__f(
            "PostNewItem",
            inspect_sig(ShellFileOperationProgressSinkForCall.post_newitem)
            .bind(self, flags, dest_folder, newname, templatename, file_attrs, hr_new, newly_created_item)
            .arguments,
        )
        return hresult.S_OK

    @override
    def update_progress(self, work_total: int, work_so_far: int) -> int:
        self.__f(
            "UpdateProgress",
            inspect_sig(ShellFileOperationProgressSinkForCall.update_progress)
            .bind(self, work_total, work_so_far)
            .arguments,
        )
        return hresult.S_OK

    @staticmethod
    def for_print(f: TextIO | None = None) -> "ShellFileOperationProgressSinkForCall":
        """メソッドの呼び出しをすべてprintで報告するインスタンスを作成します。"""
        return ShellFileOperationProgressSinkForCall(lambda s, dict: print(f"{s}({dict})", file=f))

    @staticmethod
    def for_pprint(f: TextIO | None = None) -> "ShellFileOperationProgressSinkForCall":
        """メソッドの呼び出しをすべてprintとpformatの組み合わせで報告するインスタンスを作成します。"""
        return ShellFileOperationProgressSinkForCall(
            lambda s, dict: print(f"{s}:\n{pformat(dict, sort_dicts=False)}", file=f)
        )
