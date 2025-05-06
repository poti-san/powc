# powcshell

Pythonからシェル（特にシェルアイテム）を使いやすくするパッケージです。powcpropsysパッケージに依存します。

次のようなコードを記述できます。

## デスクトップのシェル項目の表示名と解析名列挙

```python
from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

item = ShellItem2.create_knownfolder(KnownFolderID.DESKTOP)
for item in item.items:
    print(f"{item.name_normaldisplay}: {item.name_desktopabsediting}")
```

## 特殊フォルダを関連付けで開く

```python
from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

# フォント
ShellItem2.create_knownfolder(KnownFolderID.FONTS).execute_fs("open")

# プログラムのアンインストールまたは変更
ShellItem2.create_knownfolder(KnownFolderID.CHANGE_REMOVE_PROGRAMS).execute_fs("open")
```

## シェルリンクの作成と読み込み

```python
from pathlib import Path

from powcshell.shelllink import ShellLink

scriptdir_path = Path(__file__).parent
testtext_path = scriptdir_path / "test.txt"

if not testtext_path.is_file():
    with testtext_path.open("w", encoding="utf-16le") as f:
        f.write("012 abc あいう 🍊🍎🍑")
        f.flush()

link1 = ShellLink.create()
link1.path = str(testtext_path)
link1.persist_file.save(str(scriptdir_path / "test.txt.lnk"), True)

link2 = ShellLink.create_from_file(str(scriptdir_path / r"test.txt.lnk"))
print(link2.persist_file.curfile)
```

## 一つのシェル項目からなるシェル項目配列の作成

```python
from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2
from powcshell.shellitemarray import ShellItemArray

item = ShellItem2.create_knownfolder(KnownFolderID.DESKTOP)
array = ShellItemArray.create_from_item(item)

print(len(array))
```

## ファイルをシェル操作で削除

```python
from pathlib import Path
from pprint import pformat

from powcshell.shellfileop import ShellFileOperation
from powcshell.shellfileoputil import ShellFileOperationProgressSinkForCall
from powcshell.shellitem2 import ShellItem2

scriptdir_path = Path(__file__).parent
testfile_path = scriptdir_path / "test.txt"

# テストファイルの作成
with testfile_path.open("w") as f:
    print("TEST", file=f)
    f.flush()

testfile_item = ShellItem2.create_parsingname(str(testfile_path))

sink = ShellFileOperationProgressSinkForCall(
    lambda funcname, args: print(f"{funcname}:\n{pformat(args, sort_dicts=False)}")
)
op = ShellFileOperation.create()
op.advise(sink)
op.delete_item(testfile_item)
op.perform_operations()
```
