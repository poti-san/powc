"""ボリューム一覧のシステムプロパティの表示"""

from powcshell.shellitemutil.volumeinfo import ShellVolumeInfo

for info in ShellVolumeInfo.iter():
    print(f"{info.item.name_normaldisplay} {info.filesystem} {info.capacity}")
