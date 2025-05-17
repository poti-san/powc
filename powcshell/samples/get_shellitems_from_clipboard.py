"""クリップボードからシェル項目配列の取得"""

from powcshell.shellitemarray import ShellItemArray

items = ShellItemArray.getitems_fromclipboard_nothrow()
if items:
    print(tuple(item.name_normaldisplay for item in items.value_unchecked))
else:
    print("クリップボードからシェル項目配列を取得できません。")
