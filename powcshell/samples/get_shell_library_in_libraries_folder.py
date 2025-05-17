from powc.stream import StorageMode
from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2
from powcshell.shelllibrary import ShellLibrary

libfolder = ShellItem2.create_knownfolder(KnownFolderID.LIBRARIES)
for libitem in libfolder.iter_items():
    lib = ShellLibrary.load_fromitem(libitem, StorageMode.READ)
    print(f"{libitem.name_normaldisplay}: {tuple(item.name_normaldisplay for item in lib.folders_all.iter_items())}")
    print(f"{libitem.name_normaldisplay}: {tuple(item.name_normaldisplay for item in lib.folders_all.iter_items())}")
