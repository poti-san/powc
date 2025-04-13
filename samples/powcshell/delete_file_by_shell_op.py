# ファイルをシェル操作で削除

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
