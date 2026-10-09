"""ブラウザ版（web/）の入口。ブラウザの中で動く Python（Pyodide）に読み込まれ、画面（web/app.mjs）から呼ばれる。

読み取り・Excelの読み込み・絵コンテの組み立ては、デスクトップ版（gui.py）と同じ関数をそのまま使う。
ここが受け持つのは、ブラウザから届いたファイルを渡すことと、できたファイルを返すことだけ。
ファイルはブラウザの中の作業場所（WORK）に置くだけで、どこにも送らない。
"""
import io, json, os, zipfile
from .omr import read_pdf, page_count
from .compose import compose, make_cut, check_cuts
from .excel_import import load_book, describe_book
from .sheet import build as build_sheet

WORK = "/tmp/ekonte"
PREVIEW = {"picture": 1000, "mark": 700}  # 画面に出す画像の長辺（px）
JPEG = b"\xff\xd8"  # JPEG のファイルの始まり


# Excelの種類: (読み込む関数, 画面に出す要約)。香盤表とカット表を1つにまとめたファイルだけ（gui.py と同じ）
EXCEL = {"book": (load_book, describe_book)}

results = []  # 最後に読み取ったカット絵（omr.PageResult。picture は圧縮して持つ。マーク欄は画面に渡したら手放す）
excel = {}    # 選ばれたExcel: 種類 → 作業場所のファイル（選んだときの名前のまま置く）
outputs = {}  # 最後に作ったもの: "pdf"（絵コンテ）/ "kouban"（香盤表）/ "images" → bytes


def _bytes(data):
    """ブラウザから届いたデータ（JavaScript の Uint8Array）を bytes にする。"""
    return data.to_bytes() if hasattr(data, "to_bytes") else bytes(data)


def _write(name, data):
    path = os.path.join(WORK, name)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(_bytes(data))
    return path


def _read(path):
    with open(path, "rb") as f:
        return f.read()


def _preview(img, size):
    """画面に出す縮小画像。カラーは JPEG、白黒は PNG。"""
    img = img.copy()
    img.thumbnail((size, size))
    buf = io.BytesIO()
    if img.mode == "RGB":
        img.save(buf, "JPEG", quality=85)
    else:
        img.save(buf, "PNG", compress_level=1)
    return buf.getvalue()


def _pack(img):
    """書き出しまで持っておくカット絵を圧縮する（1ページ約2.5MB → 0.2MB）。白黒は形を変えない PNG、
    カラーはスキャンの細かいむらで PNG が大きくなりすぎるので、画質の高い JPEG。"""
    buf = io.BytesIO()
    if img.mode == "RGB":
        img.save(buf, "JPEG", quality=92, subsampling=0)
    else:
        img.save(buf, "PNG")
    return buf.getvalue()


def read(data, on_page):
    """カット絵PDFを読み取る。1ページ読むごとに on_page(ページの情報のJSON, カット絵の画像, マーク欄のPNG) を呼ぶ。
    カット絵の画像は、カラー（情報の color）なら JPEG、白黒なら PNG。"""
    results.clear()
    path = _write("カット絵.pdf", data)
    total = page_count(path)
    for r in read_pdf(path):
        info = dict(page=r.page, total=total, name=r.name, status=r.status, warnings=r.warnings,
                    orientation=r.orientation, format=r.format, color=r.color)
        mark = _preview(r.mark_area, PREVIEW["mark"]) if r.mark_area is not None else b""
        on_page(json.dumps(info, ensure_ascii=False), _preview(r.picture, PREVIEW["picture"]), mark)
        # 何百ページ読んでもメモリが足りるよう、書き出しに使うカット絵は圧縮して持ち、マーク欄は手放す
        r.picture, r.mark_area = _pack(r.picture), None
        results.append(r)
    return total


def set_excel(key, data, name):
    """Excelを読み込んで中身を確かめ、要約（「香盤表 3シーン・カット表 8カット」など）を返す。
    読めなければ例外を投げ、前に選んだものを残す。name: 選んだファイルの名前（書き出す香盤表に書き添える）"""
    loader, describe = EXCEL[key]
    name = os.path.basename(str(name).replace("\\", "/")) or "香盤表・カット表.xlsx"
    if not name.lower().endswith((".xlsx", ".xlsm")):
        name += ".xlsx"
    new = _write(os.path.join("確認中", name), data)
    try:
        summary = describe(loader(new))
    except Exception:
        os.remove(new)
        raise
    path = os.path.join(WORK, key, name)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if excel.get(key) not in (None, path):
        os.remove(excel[key])
    os.replace(new, path)
    excel[key] = path
    return summary


def clear_excel(key):
    excel.pop(key, None)


def check(names):
    """カット表とカット絵の突き合わせ。names: カット絵の番号の一覧（JSON）。戻り値: {"missing": [...], "extra": [...]} のJSON"""
    missing, extra = check_cuts(json.loads(names), excel["book"]) if "book" in excel else ([], [])
    return json.dumps(dict(missing=missing, extra=extra), ensure_ascii=False)


def export(rows, fmt, with_images, on_progress=None):
    """絵コンテPDFと香盤表Excelを作る。rows: [[読み取り結果の何番目か（0始まり）, カットの番号], ...] のJSON。

    戻り値は compose の結果（ページ数・カット数・警告）のJSON。できたファイルは output() で受け取る。
    with_images: カット絵の画像（2-3.png など）も Zip にまとめる（デスクトップ版の「フォルダに書き出す」にあたる）。
    on_progress(済んだカット数, 全カット数): 進み具合を知らせる
    """
    cuts = [make_cut(name, io.BytesIO(results[i].picture)) for i, name in json.loads(rows)]
    path = os.path.join(WORK, "絵コンテ.pdf")
    kouban = os.path.join(WORK, "絵コンテ_香盤表.xlsx")
    os.makedirs(WORK, exist_ok=True)
    r = compose(cuts, path, excel.get("book"), fmt, kouban_out=kouban, progress=on_progress)
    outputs["pdf"] = _read(path)
    outputs["kouban"] = _read(kouban)
    outputs["images"] = _images_zip(cuts) if with_images else b""
    return json.dumps(r, ensure_ascii=False)


def _images_zip(cuts):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for c in cuts:
            if c["image"] is not None:
                data = c["image"].getvalue()  # 読み取り時に圧縮してある（カラーは JPEG、白黒は PNG）
                ext = "jpg" if data.startswith(JPEG) else "png"
                z.writestr(f"{c['s']}-{c['c']}.{ext}", data)
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
