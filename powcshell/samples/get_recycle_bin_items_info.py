# ごみ箱内のシェル項目情報の表示

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2
from powcshell.shellitemutil.recyclebiniteminfo import ShellRecycleBinItemInfo

recyclebin = ShellItem2.create_knownfolder(KnownFolderID.RECYCLE_BIN_FOLDER)
for item in recyclebin.items:
    info = ShellRecycleBinItemInfo(item)
    print(f"{item.name_desktopabsparsing}, {info.date_deleted}, {info.deletedfrom}")
