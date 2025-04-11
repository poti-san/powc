# クラスモニカーとファイルモニカーを作成して表示名を取得します。

from comtypes import GUID

from powc.comobj import Moniker

clsmoniker = Moniker.create_classmoniker(GUID.from_progid("Shell.Application"))
filemoniker = Moniker.create_file(r"C:\\")

print((clsmoniker.displayname, filemoniker.displayname))
