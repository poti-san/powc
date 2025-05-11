"""シェル既知フォルダのフォルダIDと名前の列挙"""

from powcshell.knownfolder import KnownFolderManager

manager = KnownFolderManager.create()
for folderid in manager.folderids:
    kfnt = manager.get_folder_nothrow(folderid)
    if not kfnt:
        continue
    kf = kfnt.value_unchecked
    print(f"{folderid}, {kf.folder_definition.name}")

pass
