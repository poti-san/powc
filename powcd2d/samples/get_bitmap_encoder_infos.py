# イメージエンコーダーのMIMEタイプと対応拡張子を取得する。

from powcd2d.wic import WICImagingFactory

factory = WICImagingFactory.create()

for info in factory.iter_infos_encoder():
    print((info.friendlyname, info.mimetypes, info.fileextensions))
