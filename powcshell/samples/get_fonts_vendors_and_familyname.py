# フォントのベンダー一覧とファミリー名の表示

from collections import defaultdict
from operator import itemgetter

from powcshell.knownfolderid import KnownFolderID
from powcshell.shellitem2 import ShellItem2
from powcshell.shellitemutil.fontiteminfo import ShellFontItemInfo

font_folder = ShellItem2.create_knownfolder(KnownFolderID.FONTS)
result = defaultdict(list)
for font_file in font_folder.iter_items():
    info = ShellFontItemInfo(font_file)
    result[info.vendors or ()].append(info.familyname or "")

print(tuple((vendors, len(familynames)) for vendors, familynames in sorted(result.items(), key=itemgetter(0))))

# ファミリー名を個数にせず表示する場合は以下のコードを使います。
# pp(sorted(result.items(), key=itemgetter(0)))
