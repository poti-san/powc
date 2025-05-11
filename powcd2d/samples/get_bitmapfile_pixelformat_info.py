# ビットマップファイルのピクセルフォーマット情報を取得する。

from pathlib import Path

from powc.stream import ComStream

from powcd2d.wic import WICImagingFactory

testfile_path = Path(__file__).parent / "2x2.bmp"
stream = ComStream.openread_on_file(str(testfile_path))

factory = WICImagingFactory.create()
decoder = factory.create_decoder_from_stream(stream, 0)
framedecode = decoder.get_frame(0)
info = factory.create_componentinfo(framedecode.pixelformat)
print(
    (
        info.friendlyname,
        repr(info.componenttype),
        info.author,
        info.version_nothrow.value_unchecked if (ver := info.version_nothrow) else None,
    )
)
