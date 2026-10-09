"""アプリに同梱するサンプル・テンプレート（画面の「サンプル・テンプレート…」から保存する）。

一覧はここ1か所で管理し、ekonte.spec もこれを見て同梱するファイルを決める。
"""
import os, shutil, sys, tempfile, zipfile

REPO = os.path.join(os.path.dirname(__file__), "..")

# (分類, 保存するときの名前, リポジトリ内の場所（"作る:記入例"/"作る:空" はその場で香盤表・カット表を作る）, 説明)
FILES = [
    ("記入例", "カット絵_記入例.pdf", "サンプル/カット絵_記入例.pdf", "カット絵用紙に絵を描き、番号を塗った例"),
    ("記入例", "カット絵_記入例_スキャン.pdf", "サンプル/カット絵_記入例_スキャン.pdf", "記入例を印刷してスキャンした例。① でそのまま読み取れます"),
    ("記入例", "香盤表・カット表_記入例.xlsx", "作る:記入例", "記入例と同じ作品の香盤表とカット表（1つのファイル。書き方の説明つき）"),
    ("記入例", "絵コンテ_出力例.pdf", "サンプル/絵コンテ_出力例.pdf", "上の3つから書き出した絵コンテ"),
    ("記入例", "絵コンテ_出力例_香盤表.xlsx", "サンプル/絵コンテ_出力例_香盤表.xlsx", "絵コンテと一緒に書き出される香盤表（シーンごと・カットごと）"),
    ("縦型（9:16）の記入例", "カット絵_記入例_縦.pdf", "サンプル/カット絵_記入例_縦.pdf", "縦型のカット絵用紙に描いた例（香盤表・カット表は上と共通）"),
    ("縦型（9:16）の記入例", "カット絵_記入例_縦_スキャン.pdf", "サンプル/カット絵_記入例_縦_スキャン.pdf", "縦型の記入例をスキャンした例。① でそのまま読み取れます"),
    ("縦型（9:16）の記入例", "絵コンテ_出力例_縦.pdf", "サンプル/絵コンテ_出力例_縦.pdf", "縦型の記入例から書き出した絵コンテ（1ページ4カット）"),
    ("テンプレート", "香盤表・カット表_テンプレート.xlsx", "作る:空", "空の香盤表とカット表（1つのファイル。書き方の説明つき）"),
]

# 香盤表・カット表を作るときの土台（部の香盤表）。記入例か空か → リポジトリ内の場所
KOUBAN_BASE = {True: "サンプル/香盤表_記入例.xlsx", False: "香盤表_テンプレート.xlsx"}


ZIP_NAME = "絵コンテ作成ソフト_サンプル"
# Zip の中のフォルダ名（Windows で使えない「:」を避ける）
ZIP_DIRS = {"記入例": "1_記入例", "縦型（9:16）の記入例": "2_縦型の記入例", "テンプレート": "3_テンプレート"}


def bundled():
    """PyInstaller の datas 用：(リポジトリ内の場所, 同梱先フォルダ)"""
    srcs = [src for _, _, src, _ in FILES if not src.startswith("作る:")] + list(KOUBAN_BASE.values())
    return [(src, "samples") for src in dict.fromkeys(srcs)]


def path_of(src):
    """リポジトリ内の場所 → 実際のファイルの場所（開発時はリポジトリ、アプリ化後は同梱先）"""
    base = getattr(sys, "_MEIPASS", None)
    return os.path.join(base, "samples", os.path.basename(src)) if base else os.path.join(REPO, src)


def save(name, dest):
    """name のファイルを dest に保存する。"""
    src = next(s for _, n, s, _ in FILES if n == name)
    if src.startswith("作る:"):
        from .templates import build
        build(dest, example=(src == "作る:記入例"))
        return
    shutil.copyfile(path_of(src), dest)


def save_zip(dest):
    """全ファイルを1つの Zip にまとめて dest に保存する。中身は 絵コンテ作成ソフト_サンプル/分類/ファイル。"""
    with tempfile.TemporaryDirectory() as tmp, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for group, name, _, _ in FILES:
            path = os.path.join(tmp, name)
            save(name, path)
            z.write(path, f"{ZIP_NAME}/{ZIP_DIRS[group]}/{name}")
