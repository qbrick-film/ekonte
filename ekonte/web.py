"""ブラウザ版（web/）の入口。ブラウザの中で動く Python（Pyodide）に読み込まれ、画面（web/app.mjs）から呼ばれる。

読み取り・Excelの読み込み・絵コンテの組み立ては、デスクトップ版（gui.py）と同じ関数をそのまま使う。
ここが受け持つのは、ブラウザから届いたファイルを渡すことと、できたファイルを返すことだけ。
ファイルはブラウザの中の作業場所（WORK）に置くだけで、どこにも送らない。
"""
import io, json, os, zipfile
from .omr import read_pdf, page_count
from .compose import compose, make_cut, check_cuts
from .excel_import import load_kouban, load_props, load_equipment, load_cut_table
from .sheet import build as build_sheet

WORK = "/tmp/ekonte"
PREVIEW = {"picture": 1000, "mark": 700}  # 画面に出す画像の長辺（px）


def _count_rows(sources):
    n_scene = sum(len(v) for by_scene, _ in sources.values() for v in by_scene.values())
    n_cut = sum(len(v) for _, by_cut in sources.values() for v in by_cut.values())
    return f"シーン指定 {n_scene}件・カット指定 {n_cut}件"


# Excelの種類: (読み込む関数, 画面に出す要約)。要約は gui.py の FilePicker と同じ
EXCEL = {
    "kouban": (load_kouban, lambda d: f"{len(d)}シーン"),
    "cuts": (load_cut_table, lambda d: f"{len(d[0])}カット"),
    "props": (load_props, _count_rows),
    "equipment": (load_equipment, _count_rows),
}

results = []  # 最後に読み取ったカット絵（omr.PageResult）
excel = {}    # 選ばれたExcel: 種類 → 作業場所のファイル
outputs = {}  # 最後に作った絵コンテ: "pdf" / "images" → bytes


def _bytes(data):
    """ブラウザから届いたデータ（JavaScript の Uint8Array）を bytes にする。"""
    return data.to_bytes() if hasattr(data, "to_bytes") else bytes(data)


def _write(name, data):
    os.makedirs(WORK, exist_ok=True)
    path = os.path.join(WORK, name)
    with open(path, "wb") as f:
        f.write(_bytes(data))
    return path


def _read(path):
    with open(path, "rb") as f:
        return f.read()


def _png(img, size):
    img = img.copy()
    img.thumbnail((size, size))
    buf = io.BytesIO()
    img.save(buf, "PNG", compress_level=1)
    return buf.getvalue()


def read(data, on_page):
    """カット絵PDFを読み取る。1ページ読むごとに on_page(ページの情報のJSON, カット絵のPNG, マーク欄のPNG) を呼ぶ。"""
    results.clear()
    path = _write("カット絵.pdf", data)
    total = page_count(path)
    for r in read_pdf(path):
        results.append(r)
        info = dict(page=r.page, total=total, name=r.name, status=r.status, warnings=r.warnings,
                    orientation=r.orientation, format=r.format)
        mark = _png(r.mark_area, PREVIEW["mark"]) if r.mark_area is not None else b""
        on_page(json.dumps(info, ensure_ascii=False), _png(r.picture, PREVIEW["picture"]), mark)
    return total


def set_excel(key, data, ext):
    """Excelを読み込んで中身を確かめ、要約（「12シーン」など）を返す。読めなければ例外を投げ、前に選んだものを残す。"""
    loader, describe = EXCEL[key]
    ext = ext if ext in (".xlsx", ".xlsm") else ".xlsx"
    new = _write(f"{key}_確認中{ext}", data)
    try:
        summary = describe(loader(new))
    except Exception:
        os.remove(new)
        raise
    path = os.path.join(WORK, f"{key}{ext}")
    os.replace(new, path)
    excel[key] = path
    return summary


def clear_excel(key):
    excel.pop(key, None)


def check(names):
    """カット表とカット絵の突き合わせ。names: カット絵の番号の一覧（JSON）。戻り値: {"missing": [...], "extra": [...]} のJSON"""
    missing, extra = check_cuts(json.loads(names), excel["cuts"])
    return json.dumps(dict(missing=missing, extra=extra), ensure_ascii=False)


def export(rows, fmt, with_images):
    """絵コンテPDFを作る。rows: [[読み取り結果の何番目か（0始まり）, カットの番号], ...] のJSON。

    戻り値は compose の結果（ページ数・カット数・警告）のJSON。できたファイルは output() で受け取る。
    with_images: カット絵の画像（2-3.png など）も Zip にまとめる（デスクトップ版の「フォルダに書き出す」にあたる）。
    """
    cuts = [make_cut(name, results[i].picture) for i, name in json.loads(rows)]
    path = os.path.join(WORK, "絵コンテ.pdf")
    os.makedirs(WORK, exist_ok=True)
    r = compose(cuts, path, excel.get("kouban"), excel.get("props"), excel.get("equipment"), excel.get("cuts"), fmt)
    outputs["pdf"] = _read(path)
    outputs["images"] = _images_zip(cuts) if with_images else b""
    return json.dumps(r, ensure_ascii=False)


def _images_zip(cuts):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for c in cuts:
            if c["image"] is not None:
                png = io.BytesIO()
                c["image"].save(png, "PNG")
                z.writestr(f"{c['s']}-{c['c']}.png", png.getvalue())
    return buf.getvalue()


def output(kind):
    return outputs.get(kind, b"")


def sheet(pages, fmt):
    """カット絵用紙のPDF。"""
    path = os.path.join(WORK, "カット絵用紙.pdf")
    os.makedirs(WORK, exist_ok=True)
    build_sheet(path, int(pages), fmt)
    return _read(path)


def info():
    """動作確認用：使っている部品の版と、悪意のあるExcel（XML爆弾）への対策が効いているか。"""
    import sys
    from importlib.metadata import version
    from openpyxl.xml import DEFUSEDXML
    parts = {name: version(name) for name in ("numpy", "pillow", "pypdfium2", "reportlab", "openpyxl", "defusedxml")}
    return json.dumps(dict(python=sys.version.split()[0], **parts, excel_guard=DEFUSEDXML))
