# COMカテゴリのカテゴリIDと説明を取得します。

from powc.comcat import CategoryInformation

catinfo = CategoryInformation.create()
for cat in catinfo.get_enumcategories():
    print((cat.catid, cat.description))
