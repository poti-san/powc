# powcパッケージ集

```{toctree}
---
hidden:
---
package_summary/index
apiref
genindex
modindex
```

PythonからWindowsのCOMを使いやすくするためのパッケージ集です。基本的にcomtypesパッケージまたはパッケージ集内のパッケージのみに依存する設計です。

## パッケージ

|パッケージ|概要|依存関係|
|:--|:--|:--|
|powc|COM基本機能|comtypes|
|powcpropsys|プロパティシステム|powc|
|powcshell|シェル（特にシェルアイテム）|powcpropsys|
|powcwmi|WMI|powc|
|powccoreaudio|CoreAudio API|powcpropsys|
|powcd2d|Direct2D関係|powcpropsys|

```{mermaid}
flowchart TB
	comtypes(comtypes) --> powc
	powc --> powcpropsys
	powcpropsys --> powcshell
	powc --> powcwmi
	powcpropsys --> powccoreaudio
	powcpropsys --> powcd2d
```

## インストール

ローカルパッケージとしての利用を想定しています。ダウンロードして次のいずれかの方法でインストールできます。

1. `powc/pip install e.ps1`の実行→全てのパッケージをまとめてインストールします。
2. `powc/powc<XXX>/pip install e.ps1`の実行→それぞれのパッケージをインストールします。

## ファイル構成

|ファイル名|説明|
|:--|:--|
|readme.md|このファイル。|
|パッケージコーディング方針.md|パッケージのコーディング方針。|
|pip install e.ps1|全てのパッケージを開発モードで登録します。|
|sphnix apidoc.ps1|sphnix-apidocで各パッケージのrstファイルを作成します。|
|sphnix make html.ps1|sphnixで各パッケージのドキュメントを作成します。|
|doc.ps1|パッケージ集のドキュメントを開きます。|
|powc<XXX>/|各パッケージのディレクトリ。|
|powc<XXX>/samples/|パッケージのサンプルコードディレクトリ。|
|powc<XXX>/src/|パッケージのソースコードディレクトリ。|
|sphnix/|Sphnix用のデータディレクトリ。|
|docs/|Sphnixで作成したhtmlデータの複製。|

## 謝辞

このパッケージ集はcomtypesパッケージに依存しています。comtypesパッケージの作成者や貢献者の方々の多大な努力に感謝します。
