# powc readme

PythonからWindowsのCOM機能を使いやすくするためのパッケージ集です。
基本的にcomtypesパッケージのみに依存する設計です。

## GitHubリポジトリに関する注意
GitHubに不慣れなので、更新時はおそらくすべて削除してアップロードしなおします。

## ファイル構成

|ファイル名|説明|
|:--|:--|
|readme.md|このファイル。|
|パッケージコーディング方針.md|パッケージのコーディング方針。|
|pip install e.ps1|各パッケージを開発モードで登録します。|
|pip uninstall.ps1|各パッケージを登録解除します。|
|sphnix apidoc.ps1|sphnix-apidocで各パッケージのrstファイルを作成します。|
|sphnix make html.ps1|sphnixで各パッケージのドキュメントを作成します。|
|doc powc.ps1<br>doc powcpropsys.ps1<br>doc powcshell.ps1|各パッケージのドキュメントを開きます。先に`sphnix make html.ps1`を実行してください。|
|_packagenames.ps1|このディレクトリ内のps1ファイルで参照するパッケージ名変数を定義します。|
|powc/|powcパッケージ。|
|powcpropsys/|powcpropsysパッケージ。powcに依存します。|
|powcshell/|powcshellパッケージ。powcとpowcpropsysに依存します。|

## 謝辞

このパッケージ集はcomtypesパッケージに依存しています。comtypesパッケージの作成者や貢献者の方々の多大な努力に感謝します。
