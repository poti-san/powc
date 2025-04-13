# powc readme

PythonからWindowsのCOM機能を使いやすくするためのパッケージ集です。
基本的にcomtypesパッケージのみに依存する設計です。

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

**システムプロパティの列挙**

```python
from powcpropsys.propsys import PropertySystem

ps = PropertySystem.create()

for desc in ps.propdescs_all:
    print(f"{str(desc)} {desc.propkey}")
```

## インストール

ローカルパッケージとしての利用を前提としています。使用時はダウンロードして`pip install e.ps1`を実行してください。

## ファイル構成

|ファイル名|説明|
|:--|:--|
|readme.md|このファイル。|
|パッケージコーディング方針.md|パッケージのコーディング方針。|
|pip install e.ps1|各パッケージを開発モードで登録します。|
|pip uninstall.ps1|各パッケージを登録解除します。|
|sphnix apidoc.ps1|sphnix-apidocで各パッケージのrstファイルを作成します。|
|sphnix make html.ps1|sphnixで各パッケージのドキュメントを作成します。|
|doc.ps1|パッケージ集のドキュメントを開きます。|
|src/|ソースコードディレクトリ。|
|samples/|サンプルコードディレクトリ。|
|sphnix/|Sphnix用のデータディレクトリ。|
|docs|Sphnixで作成したhtmlデータの複製。Windows環境でctypesを使用しているため、作成してアップロードします。|


## 謝辞

このパッケージ集はcomtypesパッケージに依存しています。comtypesパッケージの作成者や貢献者の方々の多大な努力に感謝します。
