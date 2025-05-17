from powcshell.desktopwallpaper import DesktopWallpaper

wallpaper = DesktopWallpaper.create()

print(
    f"""
    壁紙:{wallpaper.get_wallpaper(0)}
    表示位置:{wallpaper.position}
    背景色:{wallpaper.bgcolor}
    モニタデバイスパス:{tuple(wallpaper.monitordevicepaths)}
    """
)
