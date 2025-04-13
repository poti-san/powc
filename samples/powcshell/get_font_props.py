# 登録されたフォントのプロパティストアを取得する。

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

# フォントフォルダでは「MS UI Gothic 標準」ですが、解析名は「MS UI Gothic Regular」です。
font_item = ShellItem2.create_knownfolder_item(KnownFolderID.FONTS, "MS UI Gothic Regular")
propstore = font_item.propstore

for key, value in propstore.iter_items():
    print(f"{key}: {value}")
