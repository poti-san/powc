"""デスクトップのシェル項目の表示名と解析名列挙"""

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

item = ShellItem2.create_knownfolder(KnownFolderID.DESKTOP)

for item in item.items:
    print(f"{item.name_normaldisplay}: {item.name_desktopabsediting}")
