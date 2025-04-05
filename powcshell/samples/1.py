"""デスクトップの表示名と属性の取得"""

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

item = ShellItem2.create_knownfolder(KnownFolderID.DESKTOP)

print(f"{item.name_normaldisplay} {repr(item.attributes)}")
