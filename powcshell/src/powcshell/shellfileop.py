"""シェル項目のファイル操作。

主なクラスは :class:`ShellFileOperation` です。
"""

import typing
from ctypes import POINTER, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from enum import IntFlag
from typing import Any

from comtypes import GUID, STDMETHOD, CoCreateInstance, COMObject, IUnknown, hresult

from powc.core import ComResult, cr, query_interface
from powcpropsys.propchange import PropertyChangeArray
from powcshell.shellitemarray import ShellItemArray, ShellItemAttributeFlag
from powcshell.shellitemenum import EnumShellItems

from .shellitem import IShellItem, ShellItem
from .shellitem2 import ShellItem2


class TransferSourceFlag(IntFlag):
    """TRANSFER_SOURCE_FLAGS"""

    NORMAL = 0
    FAIL_EXIST = 0
    RENAME_EXIST = 0x1
    OVERWRITE_EXIST = 0x2
    ALLOW_DECRYPTION = 0x4
    NO_SECURITY = 0x8
    COPY_CREATION_TIME = 0x10
    COPY_WRITE_TIME = 0x20
    USE_FULL_ACCESS = 0x40
    DELETE_RECYCLE_IF_POSSIBLE = 0x80
    COPY_HARD_LINK = 0x100
    COPY_LOCALIZED_NAME = 0x200
    MOVE_AS_COPY_DELETE = 0x400
    SUSPEND_SHELLEVENTS = 0x800


class ShellFileOperationFlag(IntFlag):
    """FOF_*定数およびFOFX_*定数"""

    MULTI_DEST_FILES = 0x0001
    CONFIRM_MOUSE = 0x0002
    SILENT = 0x0004
    RENAME_ON_COLLISION = 0x0008
    NO_CONFIRMATION = 0x0010
    WANT_MAPPING_HANDLE = 0x0020
    ALLOW_UNDO = 0x0040
    FILES_ONLY = 0x0080
    SIMPLE_PROGRESS = 0x0100
    NO_CONFIRM_MKDIR = 0x0200
    NO_ERROR_UI = 0x0400
    NO_COPY_SECURITY_ATTRIBS = 0x0800
    NO_RECURSION = 0x1000
    NO_CONNECTED_ELEMENTS = 0x2000
    WANT_NUKE_WARNING = 0x4000
    NO_RECURSE_REPARSE = 0x8000
    NO_UI = SILENT | NO_CONFIRMATION | NO_ERROR_UI | NO_CONFIRM_MKDIR
    NO_SKIP_JUNCTIONS = 0x00010000
    PREFER_HARD_LINK = 0x00020000
    SHOW_ELEVATION_PROMPT = 0x00040000
    RECYCLE_ON_DELETE = 0x00080000
    EARLY_FAILURE = 0x00100000
    PRESERVE_FILE_EXTENSIONS = 0x00200000
    KEEP_NEWER_FILE = 0x00400000
    NO_COPY_HOOKS = 0x00800000
    NO_MINIMIZE_BOX = 0x01000000
    MOVE_ACLS_ACROSS_VOLUMES = 0x02000000
    DONT_DISPLAYS_OUR_CE_PATH = 0x04000000
    DONT_DISPLAY_DEST_PATH = 0x08000000
    REQUIRE_ELEVATION = 0x10000000
    ADD_UNDO_RECORD = 0x20000000
    COPY_AS_DOWNLOAD = 0x40000000
    DONT_DISPLAY_LOCATIONS = 0x80000000


class IFileOperationProgressSink(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{04b0f1a7-9490-44bc-96e1-4296a31252e2}")
    _methods_ = [
        STDMETHOD(c_int32, "StartOperations", ()),
        STDMETHOD(c_int32, "FinishOperations", (c_int32,)),
        STDMETHOD(c_int32, "PreRenameItem", (c_uint32, POINTER(IShellItem), c_wchar_p)),
        STDMETHOD(c_int32, "PostRenameItem", (c_uint32, POINTER(IShellItem), c_wchar_p, c_int32, POINTER(IShellItem))),
        STDMETHOD(c_int32, "PreMoveItem", (c_uint32, POINTER(IShellItem), POINTER(IShellItem), c_wchar_p)),
        STDMETHOD(
            c_int32,
            "PostMoveItem",
            (c_uint32, POINTER(IShellItem), POINTER(IShellItem), c_wchar_p, c_int32, POINTER(IShellItem)),
        ),
        STDMETHOD(c_int32, "PreCopyItem", (c_uint32, POINTER(IShellItem), POINTER(IShellItem), c_wchar_p)),
        STDMETHOD(
            c_int32,
            "PostCopyItem",
            (c_uint32, POINTER(IShellItem), POINTER(IShellItem), c_wchar_p, c_int32, POINTER(IShellItem)),
        ),
        STDMETHOD(c_int32, "PreDeleteItem", (c_uint32, POINTER(IShellItem))),
        STDMETHOD(
            c_int32,
            "PostDeleteItem",
            (c_uint32, POINTER(IShellItem), c_int32, POINTER(IShellItem)),
        ),
        STDMETHOD(c_int32, "PreNewItem", (c_uint32, POINTER(IShellItem), c_wchar_p)),
        STDMETHOD(
            c_int32,
            "PostNewItem",
            (c_uint32, POINTER(IShellItem), c_wchar_p, c_wchar_p, c_uint32, c_int32, POINTER(IShellItem)),
        ),
        STDMETHOD(c_int32, "UpdateProgress", (c_uint32, c_uint32)),
        STDMETHOD(c_int32, "ResetTimer", ()),
        STDMETHOD(c_int32, "PauseTimer", ()),
        STDMETHOD(c_int32, "ResumeTimer", ()),
    ]


class IFileOperation(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{947aab5f-0a5c-4c13-b4d6-4bf7836fc9f8}")
    _methods_ = [
        STDMETHOD(c_int32, "Advise", (POINTER(IFileOperationProgressSink), POINTER(c_uint32))),
        STDMETHOD(c_int32, "Unadvise", (c_uint32,)),
        STDMETHOD(c_int32, "SetOperationFlags", (c_uint32,)),
        STDMETHOD(c_int32, "SetProgressMessage", (c_wchar_p,)),
        STDMETHOD(c_int32, "SetProgressDialog", (POINTER(IUnknown),)),  # TODO IOperationsProgressDialog
        STDMETHOD(c_int32, "SetProperties", (POINTER(IUnknown),)),  # TODO IPropertyChangeArray
        STDMETHOD(c_int32, "SetOwnerWindow", (c_void_p,)),
        STDMETHOD(c_int32, "ApplyPropertiesToItem", (POINTER(IShellItem),)),
        STDMETHOD(c_int32, "ApplyPropertiesToItems", (POINTER(IUnknown),)),
        STDMETHOD(c_int32, "RenameItem", (POINTER(IShellItem), c_wchar_p, POINTER(IFileOperationProgressSink))),
        STDMETHOD(c_int32, "RenameItems", (POINTER(IUnknown), c_wchar_p)),
        STDMETHOD(
            c_int32,
            "MoveItem",
            (POINTER(IShellItem), POINTER(IShellItem), c_wchar_p, POINTER(IFileOperationProgressSink)),
        ),
        STDMETHOD(c_int32, "MoveItems", (POINTER(IUnknown), POINTER(IShellItem))),
        STDMETHOD(
            c_int32,
            "CopyItem",
            (POINTER(IShellItem), POINTER(IShellItem), c_wchar_p, POINTER(IFileOperationProgressSink)),
        ),
        STDMETHOD(c_int32, "CopyItems", (POINTER(IUnknown), POINTER(IShellItem))),
        STDMETHOD(
            c_int32,
            "DeleteItem",
            (POINTER(IShellItem), POINTER(IFileOperationProgressSink)),
        ),
        STDMETHOD(c_int32, "DeleteItems", (POINTER(IUnknown),)),
        STDMETHOD(
            c_int32,
            "NewItem",
            (POINTER(IShellItem), c_uint32, c_wchar_p, c_wchar_p, POINTER(IFileOperationProgressSink)),
        ),
        STDMETHOD(c_int32, "PerformOperations", ()),
        STDMETHOD(c_int32, "GetAnyOperationsAborted", (POINTER(c_int32),)),
    ]


class ShellFileOperationProgressSinkBase:
    """ShellFileOperationのコールバックベースクラス。
    コールバックを作成する場合は継承クラスを作成して目的のメソッドを実装してください。
    実装したメソッドでは成功時はcomtypes.hresult.S_OK (0)、失敗時はcomtypes.hresult.E_FAIL (0x80004005)等を返してください。

    ユーティリティクラスとしてshellfileoputil以下に
    ShellFileOperationProgressSinkForCallやShellFileOperationProgressSinkForPrintがあります。

    Examples:
        >>> class ShellFileOperationProgressSinkDerivered(ShellFileOperationProgressSinkBase):
        >>>     ...
    """

    __slots__ = ()

    def start_operations(self) -> int:
        return hresult.S_OK

    def finish_operations(self, hr_result: int) -> int:
        return hresult.S_OK

    def pre_rename_item(self, flags: TransferSourceFlag, item: ShellItem2, newname: str) -> int:
        return hresult.S_OK

    def post_rename_item(
        self,
        flags: TransferSourceFlag,
        item: ShellItem2,
        newname: str,
        hr_rename: int,
        newly_created_item: ShellItem2,
    ) -> int:
        return hresult.S_OK

    def pre_move_item(self, flags: TransferSourceFlag, item: ShellItem2, dest_folder: ShellItem2, newname: str) -> int:
        return hresult.S_OK

    def post_move_item(
        self,
        flags: TransferSourceFlag,
        item: ShellItem2,
        dest_folder: ShellItem2,
        newname: str,
        hr_move: int,
        newly_created_item: ShellItem2,
    ) -> int:
        return hresult.S_OK

    def pre_copy_item(self, flags: TransferSourceFlag, item: ShellItem2, dest_folder: ShellItem2, newname: str) -> int:
        return hresult.S_OK

    def post_copy_item(
        self,
        flags: TransferSourceFlag,
        item: ShellItem2,
        dest_folder: ShellItem2,
        newname: str,
        hr_copy: int,
        newly_created_item: ShellItem2,
    ) -> int:
        return hresult.S_OK

    def pre_delete_item(self, flags: TransferSourceFlag, item: ShellItem2) -> int:
        return hresult.S_OK

    def post_delete_item(
        self,
        flags: TransferSourceFlag,
        item: ShellItem2,
        hr_delete: int,
        newly_created_item: ShellItem2,
    ) -> int:
        return hresult.S_OK

    def pre_new_item(self, flags: TransferSourceFlag, item: ShellItem2, newname: str) -> int:
        return hresult.S_OK

    def post_new_item(
        self,
        flags: TransferSourceFlag,
        dest_folder: ShellItem2,
        newname: str,
        templatename: str,
        file_attrs: int,
        hr_new: int,
        newly_created_item: ShellItem2,
    ) -> int:
        return hresult.S_OK

    def update_progress(self, work_total: int, work_so_far: int) -> int:
        return hresult.S_OK


class _ShellFileOperationProgressSink(COMObject):
    _com_interfaces_ = [IFileOperationProgressSink]

    IShellItemPointer = Any

    __sink: ShellFileOperationProgressSinkBase

    __slots__ = ("__sink",)

    def __init__(self, sink: ShellFileOperationProgressSinkBase) -> None:
        self.__sink = sink
        super().__init__()

    @staticmethod
    def wrap(sink: ShellFileOperationProgressSinkBase | None) -> "_ShellFileOperationProgressSink | None":
        if sink:
            return typing.cast(_ShellFileOperationProgressSink, _ShellFileOperationProgressSink(sink))
        else:
            return None

    def IFileOperationProgressSink_StartOperations(self, this) -> int:
        return self.__sink.start_operations()

    def IFileOperationProgressSink_FinishOperations(self, this, hr_result: int) -> int:
        return self.__sink.finish_operations(hr_result)

    def IFileOperationProgressSink_PreRenameItem(self, this, flags: int, item: IShellItemPointer, newname: str) -> int:
        return self.__sink.pre_rename_item(TransferSourceFlag(flags), ShellItem2(item), newname)

    def IFileOperationProgressSink_PostRenameItem(
        self,
        this,
        flags: int,
        item: IShellItemPointer,
        newname: str,
        hr_rename: int,
        newly_created_item: IShellItemPointer,
    ) -> int:
        return self.__sink.post_rename_item(
            TransferSourceFlag(flags), ShellItem2(item), newname, hr_rename, ShellItem2(newly_created_item)
        )

    def IFileOperationProgressSink_PreMoveItem(
        self, this, flags: int, item: IShellItemPointer, dest_folder: IShellItemPointer, newname: str
    ) -> int:
        return self.__sink.pre_move_item(TransferSourceFlag(flags), ShellItem2(item), ShellItem2(dest_folder), newname)

    def IFileOperationProgressSink_PostMoveItem(
        self,
        this,
        flags: int,
        item: IShellItemPointer,
        dest_folder: IShellItemPointer,
        newname: str,
        hr_move: int,
        newly_created_item: IShellItemPointer,
    ) -> int:
        return self.__sink.post_move_item(
            TransferSourceFlag(flags), ShellItem2(item), ShellItem2(dest_folder), newname, hr_move, newly_created_item
        )

    def IFileOperationProgressSink_PreCopyItem(
        self, this, flags: int, item: IShellItemPointer, dest_folder: IShellItemPointer, newname: str
    ) -> int:
        return self.__sink.pre_copy_item(TransferSourceFlag(flags), ShellItem2(item), ShellItem2(dest_folder), newname)

    def IFileOperationProgressSink_PostCopyItem(
        self,
        this,
        flags: int,
        item: IShellItemPointer,
        dest_folder: IShellItemPointer,
        newname: str,
        hr_copy: int,
        newly_created_item: IShellItemPointer,
    ) -> int:
        return self.__sink.post_copy_item(
            TransferSourceFlag(flags),
            ShellItem2(item),
            ShellItem2(dest_folder),
            newname,
            hr_copy,
            ShellItem2(newly_created_item),
        )

    def IFileOperationProgressSink_PreDeleteItem(self, this, flags: int, item: IShellItemPointer) -> int:
        return self.__sink.pre_delete_item(TransferSourceFlag(flags), ShellItem2(item))

    def IFileOperationProgressSink_PostDeleteItem(
        self,
        this,
        flags: int,
        item: IShellItemPointer,
        hr_delete: int,
        newly_created_item: IShellItemPointer,
    ) -> int:
        return self.__sink.post_delete_item(
            TransferSourceFlag(flags), ShellItem2(item), hr_delete, ShellItem2(newly_created_item)
        )

    def IFileOperationProgressSink_PreNewItem(self, this, flags: int, item: IShellItemPointer, newname: str) -> int:
        return self.__sink.pre_new_item(TransferSourceFlag(flags), ShellItem2(item), newname)

    def IFileOperationProgressSink_PostNewItem(
        self,
        this,
        flags: int,
        dest_folder: IShellItemPointer,
        newname: str,
        templatename: str,
        file_attrs: int,
        hr_new: int,
        newly_created_item: IShellItemPointer,
    ) -> int:
        return self.__sink.post_new_item(
            TransferSourceFlag(flags),
            ShellItem2(dest_folder),
            newname,
            templatename,
            file_attrs,
            hr_new,
            ShellItem2(newly_created_item),
        )

    def IFileOperationProgressSink_UpdateProgress(self, this, work_total: int, work_so_far: int) -> int:
        return self.__sink.update_progress(work_total, work_so_far)

    def IFileOperationProgressSink_ResetTimer(self, this) -> int:
        # サポートされていません。
        # return self.__sink.ResetTimer()
        return hresult.S_OK

    def IFileOperationProgressSink_PauseTimer(self, this) -> int:
        # サポートされていません。
        # return self.__sink.PauseTimer()
        return hresult.S_OK

    def IFileOperationProgressSink_ResumeTimer(self, this) -> int:
        # サポートされていません。
        # return self.__sink.ResumeTimer()
        return hresult.S_OK


class ShellFileOperation:
    """シェルのファイル操作。IShellItemインターフェイスのラッパーです。"""

    __slots__ = ("__o",)
    __o: Any  # POINTER(IFileOperation)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IFileOperation)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "ShellFileOperation":
        CLSID_FileOperation = GUID("{3ad05575-8857-4850-9277-11b85bdb8e09}")
        return ShellFileOperation(CoCreateInstance(CLSID_FileOperation, IFileOperation))

    def advise_nothrow(self, sink: ShellFileOperationProgressSinkBase) -> ComResult[int]:
        cookie = c_uint32()
        return cr(self.__o.Advise(_ShellFileOperationProgressSink.wrap(sink), byref(cookie)), cookie.value)

    def advise(self, sink: ShellFileOperationProgressSinkBase) -> int:
        return self.advise_nothrow(sink).value

    def unadvise_nothrow(self, cookie: int) -> ComResult[None]:
        return cr(self.__o.Unadvise(cookie), None)

    def unadvise(self, cookie: int) -> None:
        return self.unadvise_nothrow(cookie).value

    def set_opflags_nothrow(self, flags: ShellFileOperationFlag) -> ComResult[None]:
        return cr(self.__o.SetOperationFlags(int(flags)), None)

    def set_opflags(self, flags: ShellFileOperationFlag) -> None:
        return self.set_opflags_nothrow(flags).value

    def set_progress_msg_nothrow(self, msg: str) -> ComResult[None]:
        return cr(self.__o.SetProgressMessage(msg), None)

    def set_progress_msg(self, msg: str) -> None:
        return self.set_progress_msg_nothrow(msg).value

    # TODO
    # STDMETHOD(c_int32, "SetProgressDialog", (POINTER(IUnknown),)),  # TODO IOperationsProgressDialog

    def set_props_nothrow(self, array: PropertyChangeArray) -> ComResult[None]:
        return cr(self.__o.SetProperties(array.wrapped_obj), None)

    def set_props(self, array: PropertyChangeArray) -> None:
        return self.set_props_nothrow(array).value

    def set_owner_window_nothrow(self, window_handle: int) -> ComResult[None]:
        return cr(self.__o.SetOwnerWindow(window_handle), None)

    def set_owner_window(self, window_handle: int) -> None:
        return self.set_owner_window_nothrow(window_handle).value

    def apply_props_item_nothrow(self, item: ShellItem) -> ComResult[None]:
        return cr(self.__o.ApplyPropertiesToItem(item.wrapped_obj), None)

    def apply_props_item(self, item: ShellItem) -> None:
        return self.apply_props_item_nothrow(item).value

    def apply_props_items_nothrow(self, items: ShellItemArray | EnumShellItems) -> ComResult[None]:
        return cr(self.__o.ApplyPropertiesToItems(items.wrapped_obj), None)

    def apply_props_items(self, items: ShellItemArray | EnumShellItems) -> None:
        return self.apply_props_items_nothrow(items).value

    def rename_item_nothrow(
        self, item: ShellItem, new_name: str, sink: ShellFileOperationProgressSinkBase | None = None
    ) -> ComResult[None]:
        return cr(self.__o.RenameItem(item.wrapped_obj, new_name, _ShellFileOperationProgressSink.wrap(sink)), None)

    def rename_item(
        self, item: ShellItem, new_name: str, sink: ShellFileOperationProgressSinkBase | None = None
    ) -> None:
        return self.rename_item_nothrow(item, new_name, sink).value

    def rename_items_nothrow(self, items: ShellItemArray | EnumShellItems, new_name: str) -> ComResult[None]:
        return cr(self.__o.RenameItems(items.wrapped_obj, new_name), None)

    def rename_items(self, items: ShellItemArray | EnumShellItems, new_name: str) -> None:
        return self.rename_items_nothrow(items, new_name).value

    def move_item_nothrow(
        self,
        item: ShellItem,
        dest_folder: ShellItem,
        new_name: str,
        sink: ShellFileOperationProgressSinkBase | None = None,
    ) -> ComResult[None]:
        return cr(
            self.__o.MoveItem(
                item.wrapped_obj, dest_folder.wrapped_obj, new_name, _ShellFileOperationProgressSink.wrap(sink)
            ),
            None,
        )

    def move_item(
        self,
        item: ShellItem,
        dest_folder: ShellItem,
        new_name: str,
        sink: ShellFileOperationProgressSinkBase | None = None,
    ) -> None:
        return self.move_item_nothrow(item, dest_folder, new_name, sink).value

    def move_items_nothrow(self, items: ShellItemArray | EnumShellItems, dest_folder: ShellItem) -> ComResult[None]:
        return cr(self.__o.MoveItems(items.wrapped_obj, dest_folder.wrapped_obj), None)

    def move_items(self, items: ShellItemArray | EnumShellItems, dest_folder: ShellItem) -> None:
        return self.move_items_nothrow(items, dest_folder).value

    def copy_item_nothrow(
        self,
        item: ShellItem,
        dest_folder: ShellItem,
        copy_name: str | None = None,
        sink: ShellFileOperationProgressSinkBase | None = None,
    ) -> ComResult[None]:
        return cr(
            self.__o.CopyItem(
                item.wrapped_obj, dest_folder.wrapped_obj, copy_name, _ShellFileOperationProgressSink.wrap(sink)
            ),
            None,
        )

    def copy_item(
        self,
        item: ShellItem,
        dest_folder: ShellItem,
        copy_name: str | None = None,
        sink: ShellFileOperationProgressSinkBase | None = None,
    ) -> None:
        return self.copy_item_nothrow(item, dest_folder, copy_name, sink).value

    def copy_items_nothrow(self, items: ShellItemArray | EnumShellItems, dest_folder: ShellItem) -> ComResult[None]:
        return cr(self.__o.CopyItems(items.wrapped_obj, dest_folder.wrapped_obj), None)

    def copy_items(self, items: ShellItemArray | EnumShellItems, dest_folder: ShellItem) -> None:
        return self.copy_items_nothrow(items, dest_folder).value

    def delete_item_nothrow(
        self, item: ShellItem, sink: ShellFileOperationProgressSinkBase | None = None
    ) -> ComResult[None]:
        return cr(
            self.__o.DeleteItem(item.wrapped_obj, _ShellFileOperationProgressSink.wrap(sink)),
            None,
        )

    def delete_item(self, item: ShellItem, sink: ShellFileOperationProgressSinkBase | None = None) -> None:
        return self.delete_item_nothrow(item, sink).value

    def delete_items_nothrow(self, items: ShellItemArray | EnumShellItems) -> ComResult[None]:
        # TODO: IDataObject、IPersistIDList オブジェクト対応
        return cr(self.__o.DeleteItems(items.wrapped_obj), None)

    def delete_items(self, items: ShellItemArray | EnumShellItems) -> None:
        return self.delete_items_nothrow(items).value

    def new_item_nothrow(
        self,
        dest_folder: ShellItem,
        attrs: ShellItemAttributeFlag | int,
        name: str,
        template_name: str | None = None,
        sink: ShellFileOperationProgressSinkBase | None = None,
    ) -> ComResult[None]:
        return cr(
            self.__o.NewItem(
                dest_folder.wrapped_obj, int(attrs), name, template_name, _ShellFileOperationProgressSink.wrap(sink)
            ),
            None,
        )

    def new_item(
        self,
        dest_folder: ShellItem,
        attrs: ShellItemAttributeFlag | int,
        name: str,
        template_name: str | None = None,
        sink: ShellFileOperationProgressSinkBase | None = None,
    ) -> None:
        return self.new_item_nothrow(dest_folder, attrs, name, template_name, sink).value

    def perform_operations_nothrow(self) -> ComResult[None]:
        return cr(self.__o.PerformOperations(), None)

    def perform_operations(self) -> None:
        return self.perform_operations_nothrow().value

    @property
    def any_operations_aborted_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.GetAnyOperationsAborted(byref(x)), x.value != 0)

    @property
    def any_operations_aborted(self) -> bool:
        return self.any_operations_aborted_nothrow.value


# MIDL_INTERFACE("cd8f23c1-8f61-4916-909d-55bdd0918753")
# IFileOperation2 : public IFileOperation
# {
# public:
#     virtual HRESULT STDMETHODCALLTYPE SetOperationFlags2(
#         /* [in] */ FILE_OPERATION_FLAGS2 operationFlags2) = 0;

# };
