"""ファイルデータの読み込み

スクリプトの実行ディレクトリに存在するtest.txtの内容を読み込みます。
test.txtが存在しない場合は適当な内容で作成します。
"""

from pathlib import Path

from powcshell.shellitem2 import ShellItem2

scriptdir_path = Path(__file__).parent
testtext_path = scriptdir_path / "test.txt"

if not testtext_path.is_file():
    with testtext_path.open("w", encoding="utf-16le") as f:
        f.write("012 abc あいう 🍊🍎🍑")
        f.flush()

item = ShellItem2.create_parsingname(str(testtext_path))
stream = item.open_stream()
print(stream.read_bytes_all().decode("utf-16le"))
