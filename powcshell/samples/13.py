"""一つのシェル項目からなるシェル項目配列の作成"""

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2
from powcshell.shellitemarray import ShellItemArray

item = ShellItem2.create_knownfolder(KnownFolderID.DESKTOP)
array = ShellItemArray.create_from_item(item)

print(len(array))
