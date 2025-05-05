# 全てのWICコンポーネント情報を取得する。

from powcd2d.wic import WICImagingFactory

factory = WICImagingFactory.create()

for info in factory.iter_infos_all():
    print((info.friendlyname, info.friendlyname, info.componenttype))
