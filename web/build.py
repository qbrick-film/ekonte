"""ブラウザ版を組み立てる。  python web/build.py [出力フォルダ]（既定: dist-web）

出力フォルダの中身は、そのまま GitHub Pages に置ける。
  index.html・style.css・app.mjs・worker.mjs・engine.mjs・icon.svg   画面と計算の受け渡し（web/ から。版の番号を埋め込む）
  app.zip            ekonte の Python（画面以外）と埋め込みフォント
  samples/           サンプル・テンプレート（ekonte/samples.py の一覧と同じ）と、その Zip・一覧（samples.json）
  vendor/pyodide/    ブラウザで動く Python（Pyodide）と部品
部品は web/vendor.json に書いた場所から取り寄せ、指紋（sha256）が一致したものだけを使う（web/.cache/ に保存して使い回す）。
部品の一覧（pyodide-lock.json）も使うものだけに絞り、ブラウザが読み込むときにもう一度指紋を照合させる。
"""
import hashlib, json, os, shutil, sys, tarfile, urllib.request, zipfile

WEB = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(WEB)
sys.path.insert(0, ROOT)
from ekonte import __version__, REPO, samples

CACHE = os.path.join(WEB, ".cache")
PAGES = ["index.html", "style.css", "app.mjs", "worker.mjs", "engine.mjs", "icon.svg"]
# ブラウザ版で使う Python（gui.py・sample.py などデスクトップ版だけのものは入れない）
PY_FILES = ["__init__.py", "web.py", "omr.py", "layout.py", "compose.py", "excel_import.py", "fonts.py", "sheet.py"]
FONT_FILES = ["BIZUDGothic-Regular.ttf", "BIZUDGothic-Bold.ttf", "OFL.txt"]
MARK = ".ekonte-web"  # このスクリプトが作ったフォルダの印（印のないフォルダは消さない）
ZIP_TIME = (2020, 1, 1, 0, 0, 0)  # Zip の中の日時を固定し、同じ中身なら同じ Zip にする


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def fetch(url, sha256):
    """url のファイルを取り寄せ、指紋が一致したものの場所を返す。"""
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, f"{sha256[:16]}_{os.path.basename(url)}")
    if os.path.exists(path) and sha256_of(path) == sha256:
        return path
    print(f"  取り寄せ: {url}")
    with urllib.request.urlopen(url, timeout=300) as res, open(path + ".part", "wb") as f:
        shutil.copyfileobj(res, f)
    got = sha256_of(path + ".part")
    if got != sha256:
        os.remove(path + ".part")
        raise SystemExit(f"指紋が一致しないので使いません: {url}\n  期待 {sha256}\n  実際 {got}")
    os.replace(path + ".part", path)
    return path


def prepare(out):
    if os.path.exists(out):
        if not os.path.exists(os.path.join(out, MARK)):
            raise SystemExit(f"{out} はこのスクリプトが作ったフォルダではないので、消さずに止めます")
        shutil.rmtree(out)
    os.makedirs(out)
    open(os.path.join(out, MARK), "w").close()


def copy_pages(out):
    for name in PAGES:
        with open(os.path.join(WEB, name), encoding="utf-8") as f:
            text = f.read().replace("__VERSION__", __version__).replace("__REPO__", REPO)
        with open(os.path.join(out, name), "w", encoding="utf-8") as f:
            f.write(text)


def app_zip(out):
    def add(z, src, name):
        info = zipfile.ZipInfo(name, ZIP_TIME)
        info.compress_type = zipfile.ZIP_DEFLATED
        with open(src, "rb") as f:
            z.writestr(info, f.read(), compresslevel=9)

    with zipfile.ZipFile(os.path.join(out, "app.zip"), "w") as z:
        for name in PY_FILES:
            add(z, os.path.join(ROOT, "ekonte", name), f"ekonte/{name}")
        for name in FONT_FILES:
            add(z, os.path.join(ROOT, "resources", "fonts", name), f"resources/fonts/{name}")


def vendor(out):
    with open(os.path.join(WEB, "vendor.json"), encoding="utf-8") as f:
        cfg = json.load(f)
    dest = os.path.join(out, "vendor", "pyodide")
    os.makedirs(dest)
    core = cfg["pyodide"]
    with tarfile.open(fetch(core["url"], core["sha256"])) as tar:
        for name in core["files"] + ["pyodide-lock.json"]:
            with tar.extractfile(f"pyodide/{name}") as src, open(os.path.join(dest, name), "wb") as dst:
                shutil.copyfileobj(src, dst)
    with open(os.path.join(dest, "pyodide-lock.json"), encoding="utf-8") as f:
        lock = json.load(f)
    packages = {}
    for p in cfg["packages"]:
        file_name = os.path.basename(p["url"])
        shutil.copyfile(fetch(p["url"], p["sha256"]), os.path.join(dest, file_name))
        entry = lock["packages"].get(p["name"])
        if entry:  # Pyodide が配っている部品は、Pyodide 自身の一覧とも食い違いがないか確かめる
            if (entry["file_name"], entry["sha256"]) != (file_name, p["sha256"]):
                raise SystemExit(f"{p['name']} が Pyodide {core['version']} の一覧と違います")
        else:
            entry = dict(name=p["name"], version=p["version"], file_name=file_name, install_dir="site",
                         sha256=p["sha256"], package_type="package", imports=p["imports"], depends=p["depends"],
                         unvendored_tests=False, tool={})
        packages[p["name"]] = entry
    lock["packages"] = packages
    with open(os.path.join(dest, "pyodide-lock.json"), "w", encoding="utf-8") as f:
        json.dump(lock, f, ensure_ascii=False, indent=1)


def sample_files(out):
    dest = os.path.join(out, "samples")
    os.makedirs(dest)
    listing = []
    for group, name, _, desc in samples.FILES:
        samples.save(name, os.path.join(dest, name))
        listing.append(dict(group=group, name=name, desc=desc))
    zip_name = samples.ZIP_NAME + ".zip"
    samples.save_zip(os.path.join(dest, zip_name))
    with open(os.path.join(dest, "samples.json"), "w", encoding="utf-8") as f:
        json.dump(dict(files=listing, zip=zip_name), f, ensure_ascii=False, indent=1)


def main(out):
    out = os.path.abspath(out)
    prepare(out)
    copy_pages(out)
    app_zip(out)
    vendor(out)
    sample_files(out)
    total = 0
    for dirpath, _, files in os.walk(out):
        total += sum(os.path.getsize(os.path.join(dirpath, f)) for f in files)
    print(f"ブラウザ版 v{__version__} → {out}（{total / 1048576:.1f} MB）")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "dist-web"))
