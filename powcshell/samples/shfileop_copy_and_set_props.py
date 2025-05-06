"""ファイルを作成して属性の変更"""

from pathlib import Path

from powcpropsys.propchange import (
    PropertyChange,
    PropertyChangeAction,
    PropertyChangeArray,
)
from powcpropsys.propkey import PropertyKey
from powcpropsys.propvariant import PropVariant

from powcshell.shellfileop import ShellFileOperation
from powcshell.shellfileoputil import ShellFileOperationProgressSinkForCall
from powcshell.shellitem2 import ShellItem2

scriptdir_path = Path(__file__).parent
testfile_path = scriptdir_path / "2x2.jpg"
testfile_item = ShellItem2.create_parsingname(str(testfile_path))


sink = ShellFileOperationProgressSinkForCall.for_pprint()
op = ShellFileOperation.create()
op.advise(sink)
op.copy_item(testfile_item, testfile_item.parent, "test.jpg")

# プロパティ変更配列の作成と適用
pcarray = PropertyChangeArray.create()
pkey = PropertyKey.from_canonicalname("System.Comment")
pcarray.append(PropertyChange.create(PropertyChangeAction.SET, pkey, PropVariant.init_wstr("TEST TEXT")))
op.set_props(pcarray)

try:
    op.perform_operations()
except Exception as ex:
    print(f"test.jpgが実行前に存在した可能性があります。\n\n{ex}")
    print(f"test.jpgが実行前に存在した可能性があります。\n\n{ex}")
