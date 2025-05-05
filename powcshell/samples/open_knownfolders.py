# 特殊フォルダを関連付けで開く

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

# フォント
ShellItem2.create_knownfolder(KnownFolderID.FONTS).execute_fs("open")

# プログラムのアンインストールまたは変更
ShellItem2.create_knownfolder(KnownFolderID.CHANGE_REMOVE_PROGRAMS).execute_fs("open")
