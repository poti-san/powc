"""プロパティストアでは列挙されないがプロパティシステムには登録されたファイルプロパティの列挙"""

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

for item in ShellItem2.create_knownfolder(KnownFolderID.PRINTERS_FOLDER).iter_items():
    props = item.propstore
    propinfos = tuple(props.iter_items_in_propsystem())
    pass

pass
