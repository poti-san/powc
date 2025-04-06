"""シェルリンクの作成と読み込み"""

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
