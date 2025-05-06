# powcshell

Pythonからシェル（特にシェルアイテム）を使いやすくするパッケージです。powcpropsysパッケージに依存します。 

次のようなコードを記述できます。

**デスクトップのシェル項目の表示名と解析名列挙**

```python
from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

item = ShellItem2.create_knownfolder(KnownFolderID.DESKTOP)
for item in item.items:
    print(f"{item.name_normaldisplay}: {item.name_desktopabsediting}")
```

**特殊フォルダを関連付けで開く**

```python
from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

# フォント
ShellItem2.create_knownfolder(KnownFolderID.FONTS).execute_fs("open")

# プログラムのアンインストールまたは変更
ShellItem2.create_knownfolder(KnownFolderID.CHANGE_REMOVE_PROGRAMS).execute_fs("open")
```
