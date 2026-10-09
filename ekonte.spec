# PyInstaller の設定（Mac・Windows共通）。ビルドは build_mac.sh / build_windows.bat から。
import os, sys
sys.path.insert(0, SPECPATH)  # noqa: F821  (PyInstaller が定義)
from ekonte import __version__
from ekonte.samples import bundled as sample_files

APP = "Ekonte"                       # ファイル名は英字（Windowsでの文字化け防止）
DISPLAY = "絵コンテ作成ソフト"         # Macで表示される名前

a = Analysis(
    ["app.py"],
    datas=[("resources", "resources")] + sample_files(),  # サンプルは ekonte/samples.py の一覧
    excludes=["tkinter", "PySide6.QtWebEngineCore", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.Qt3DCore"],
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=APP, console=False, upx=False)
coll = COLLECT(exe, a.binaries, a.datas, name=APP, upx=False)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=f"{APP}.app",
        bundle_identifier="jp.qbrick.ekonte",
        info_plist={
            "CFBundleDisplayName": DISPLAY,
            "CFBundleName": DISPLAY,
            "CFBundleShortVersionString": __version__,
            "NSHighResolutionCapable": True,
        },
    )
