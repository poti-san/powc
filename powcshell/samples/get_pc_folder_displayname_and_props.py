"""PCの表示名とプロパティの取得"""

from powcshell.shellitemutil.computerfolderinfo import ShellComputerFolderInfo

info = ShellComputerFolderInfo.create()
print(
    f"""
    {info.item.name_normaldisplay}:
      Simple Name: {info.simplename}
      Memory: {info.memory}
      Processor: {info.processor}
      Work Group: {info.workgroup}
    """
)
