"""フォントのファミリ・カテゴリ・ベンダーの列挙"""

from powcshell.shellitemutil.fontiteminfo import ShellFontItemInfo

for info in ShellFontItemInfo.iter():
    print(f'Family:"{info.familyname}" Category:"{info.category or ""}" Vendors:"{info.vendors or ()}"')
