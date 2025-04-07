"""ファイルのプロパティストアではキーが列挙されないがシステムには登録されたプロパティの列挙"""

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2

if __name__ == "__main__":

    for item in ShellItem2.create_knownfolder(KnownFolderID.PRINTERS_FOLDER).iter_items():
        props = item.propstore
        propinfos = tuple(props.iter_validitems_in_propsystem())
        pass

    pass
