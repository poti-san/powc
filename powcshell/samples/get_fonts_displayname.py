# フォントの表示名列挙

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

font_folder = ShellItem2.create_knownfolder(KnownFolderID.FONTS)

for item in font_folder.items:
    print(f"{item.name_normaldisplay}")
