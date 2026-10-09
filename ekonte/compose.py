"""カット絵の画像とExcelから絵コンテPDFを組み立てる（香盤表のExcelも一緒に書き出せる）。

python -m ekonte.compose カット絵フォルダ 出力.pdf [--excel 香盤表・カット表.xlsx] [--kouban 香盤表.xlsx] [--format 縦]
  --excel  書いた香盤表・カット表（1つのファイル）
  --kouban 一緒に書き出す香盤表（kouban.py）

レイアウトは2種類。横 = A4縦の表に1行1カット（6カット/ページ）、縦 = A4横に縦型9:16のカットを4つ並べる。
"""
import argparse, html, os, re
from dataclasses import dataclass
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import Paragraph
from reportlab.lib.utils import ImageReader
from .excel_import import Book, load_book
from .fonts import JP
NAME_RE = re.compile(r"^(\d+)-(\d+)([a-j]?)\.png$")

CUT_RE = re.compile(r"^(\d+)-(\d+)([a-j]?)$")


def make_cut(name, image=None):
    """番号（"2-3a"）と画像（ファイルパス・PIL画像・None）からカットを作る。"""
    m = CUT_RE.match(name)
    if not m:
        raise ValueError(f"番号の形式が正しくありません: {name}")
    s, c, sub = m.groups()
    s, c = str(int(s)), str(int(c))
    return dict(s=s, c=c + sub, key=(int(s), int(c), sub), image=image)


def collect_cuts(folder):
    cuts = []
    for f in os.listdir(folder):
        if NAME_RE.match(f):
            cuts.append(make_cut(f[:-4], os.path.join(folder, f)))
    return cuts


# NOTE欄の項目: (キー, 表示名, カット指定があればシーン指定を置き換えるか)
#   カメラ（＋レンズ）はカット指定があればそのカットの設定に置き換え、小道具・機材・NOTEは足し合わせる
#   どれもカット表（香盤表に小道具・機材の列があればそれも）から入る
NOTE_ITEMS = [
    ("camera", "カメラ", True),
    ("gear", "機材", False),
    ("props", "小道具", False),
    ("note", None, False),   # カット表の NOTE 列（自由記入。表示名を付けない）
]


def _items(sources, key, s, c, override):
    if key not in sources:
        return []
    by_scene, by_cut = sources[key]
    cut_items = by_cut.get((s, c), [])
    if override and cut_items:
        return cut_items
    return by_scene.get(s, []) + cut_items


def to_markup(text):
    """Excelの文字をPDFの段落用に変換する。

    & < > はそのままだと書式の記号と解釈されて崩れるので置き換え、
    セル内改行（Alt+Enter / option+return）は改行として残す。
    """
    t = str(text).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    return html.escape(t, quote=False).replace("\n", "<br/>")


LIST_SEP = re.compile(r"[、,，\n]")


def without(text, shown):
    """香盤表の備考から、小道具・機材などとしてすでに載せたものを除く（「拳銃、雨降らし」→「雨降らし」）。

    同じものを2回載せないため。除くものがなければ、書いたとおりのまま返す。
    """
    parts = [p.strip() for p in LIST_SEP.split(text)]
    if not any(p in shown for p in parts):
        return text
    return "、".join(p for p in parts if p and p not in shown)


def note_lines(cut, first_in_scene, kouban, sources):
    """NOTE欄の行 [(文, 太字か)]。シーンの情報（シーンの最初のカットだけ）→ カメラ → 機材 → 小道具 → NOTE → 備考 の順。"""
    lines = []
    info = kouban.get(cut["s"], {})
    if first_in_scene and info:
        head = " / ".join(v for v in (info["bamen"], info["ls"], info["dn"]) if v)
        if head:
            lines.append((head, True))
        if info["loca"]:
            lines.append((f"ロケ地: {info['loca']}", False))
    shown = set()
    for key, label, override in NOTE_ITEMS:
        items = list(dict.fromkeys(_items(sources, key, cut["s"], cut["c"], override)))
        shown.update(items)
        if key == "camera":  # カメラは1台1行
            lines += [(f"{label}: {v}", False) for v in items]
        elif label is None:
            lines += [(v, False) for v in items]
        elif items:
            lines.append((f"{label}: " + "、".join(items), False))
    biko = without(info.get("biko", ""), shown)
    if biko:
        lines.append((f"備考: {biko}", False))
    return lines


def note_markup(lines):
    return "<br/>".join(f"<b>{to_markup(t)}</b>" if bold else to_markup(t) for t, bold in lines)


@dataclass
class Row:
    """絵コンテの1カット分の中身。PDF（compose）と香盤表Excel（kouban.py）の両方がこれを使う。"""
    cut: dict          # make_cut() の結果（画像も含む）
    first: bool        # シーンの最初のカットか
    action: list       # ACTION/SE 欄の行
    scenario: list     # SCENARIO 欄の行
    time: str          # TIME 欄
    note: list         # NOTE 欄の行 [(文, 太字か)]
    items: dict        # このカットの小道具・カメラ・機材など {キー: [...]}（香盤表でシーンごとにまとめる）


def plan(cuts, book):
    """絵コンテに載せるカットを番号順に並べ、各欄の中身を決める。

    cuts: make_cut() の結果のリスト（カット絵）。カット表にあってカット絵がないカットも、PICTURE を空けて載せる。
    """
    have = {(x["s"], x["c"]) for x in cuts}
    cuts = list(cuts) + [make_cut(f"{s}-{cc}") for s, cc in book.cuts if (s, cc) not in have]
    rows, prev = [], None
    for cut in sorted(cuts, key=lambda x: x["key"]):
        first = cut["s"] != prev
        t = book.cuts.get((cut["s"], cut["c"]), {})
        items = {k: list(dict.fromkeys(_items(book.sources, k, cut["s"], cut["c"], o))) for k, _, o in NOTE_ITEMS}
        rows.append(Row(cut, first, t.get("action", []), t.get("scenario", []), t["time"][0] if t.get("time") else "",
                        note_lines(cut, first, book.kouban, book.sources), items))
        prev = cut["s"]
    return rows


class TableLayout:
    """横：A4縦の表に1行1カット。部で使っていた絵コンテのテンプレート（PDF）と同じ罫線位置（pt、上端基準）。"""
    pagesize = A4
    per_page = 6
    COL_X = [29.1, 49.8, 71.1, 234.1, 335.6, 437.2, 470.2, 568.4]
    HEADERS = ["S", "C", "PICTURE", "ACTION/SE", "SCENARIO", "TIME", "NOTE"]
    HEAD_TOP, HEAD_BOTTOM = 81.5, 108.6
    ROW_H = 107.3

    def slot(self, i):
        """i 段目の各欄の矩形 (x, 上端, 幅, 高さ)。"""
        top, X = self.HEAD_BOTTOM + self.ROW_H * i, self.COL_X
        cell = lambda k: (X[k], top, X[k + 1] - X[k], self.ROW_H)
        return dict(s=cell(0), c=cell(1), picture=cell(2), action=cell(3), scenario=cell(4), time=cell(5), note=cell(6))

    def draw_grid(self, c, page_no):
        X, n = self.COL_X, self.per_page
        bottom = self.HEAD_BOTTOM + self.ROW_H * n
        c.setLineWidth(1)
        c.rect(X[0], y(c, bottom), X[-1] - X[0], bottom - self.HEAD_TOP)
        for x in X[1:-1]:
            c.line(x, y(c, self.HEAD_TOP), x, y(c, bottom))
        for i in range(n):
            c.line(X[0], y(c, self.HEAD_BOTTOM + self.ROW_H * i), X[-1], y(c, self.HEAD_BOTTOM + self.ROW_H * i))
        c.setFont(JP, 8)
        for i, h in enumerate(self.HEADERS):
            c.drawCentredString((X[i] + X[i + 1]) / 2, y(c, (self.HEAD_TOP + self.HEAD_BOTTOM) / 2) - 3, h)
        c.setFont(JP, 9)
        c.drawCentredString(self.pagesize[0] / 2, y(c, 818), f"No. {page_no}")


class ColumnLayout:
    """縦：A4横に縦型（9:16）のカットを4つ横に並べる。左端の列に項目名。"""
    pagesize = landscape(A4)
    per_page = 4
    LEFT, TOP, BOTTOM = 24, 30, 567
    LABEL_W = 54
    ROWS = [("head", 22), ("picture", 246), ("action", 84), ("scenario", 84), ("note", None)]  # None = 残り
    LABELS = {"head": "CUT", "picture": "PICTURE", "action": "ACTION/SE", "scenario": "SCENARIO", "note": "NOTE"}
    HEAD_SPLIT = (0.3, 0.3, 0.4)  # 1段目を S・C・TIME に分ける割合

    def __init__(self):
        self.col_w = (self.pagesize[0] - self.LEFT * 2 - self.LABEL_W) / self.per_page
        self.rows, top = {}, self.TOP
        for key, h in self.ROWS:
            h = h or self.BOTTOM - top
            self.rows[key] = (top, h)
            top += h

    def slot(self, i):
        x = self.LEFT + self.LABEL_W + self.col_w * i
        r = {k: (x, top, self.col_w, h) for k, (top, h) in self.rows.items()}
        top, h = self.rows["head"]
        cx = x
        for key, f in zip(("s", "c", "time"), self.HEAD_SPLIT):
            r[key] = (cx, top, self.col_w * f, h)
            cx += self.col_w * f
        top, h = self.rows["picture"]
        ph = h - 8  # 9:16 の枠を欄の中央に置く
        pw = ph * 9 / 16
        r["picture"] = (x + (self.col_w - pw) / 2, top + 4, pw, ph)
        return r

    def draw_grid(self, c, page_no):
        right = self.pagesize[0] - self.LEFT
        c.setLineWidth(1)
        c.rect(self.LEFT, y(c, self.BOTTOM), right - self.LEFT, self.BOTTOM - self.TOP)
        for i in range(self.per_page + 1):
            x = self.LEFT + self.LABEL_W + self.col_w * i
            c.line(x, y(c, self.TOP), x, y(c, self.BOTTOM))
        c.setFont(JP, 7)
        for key, (top, h) in self.rows.items():
            if top > self.TOP:
                c.line(self.LEFT, y(c, top), right, y(c, top))
            c.drawCentredString(self.LEFT + self.LABEL_W / 2, y(c, top + h / 2) - 2.5, self.LABELS[key])
        for i in range(self.per_page):
            s = self.slot(i)
            c.setFont(JP, 6)
            c.setFillGray(0.45)
            for key, label in (("s", "S"), ("c", "C"), ("time", "TIME")):
                x, top, w, h = s[key]
                if key != "s":
                    c.line(x, y(c, top), x, y(c, top + h))
                c.drawString(x + 3, y(c, top + 8), label)
            c.setFillGray(0)
        c.setFont(JP, 9)
        c.drawCentredString(self.pagesize[0] / 2, y(c, 582), f"No. {page_no}")


LAYOUTS = {"横": TableLayout(), "縦": ColumnLayout()}


def draw_page(c, layout, page_no):
    """罫線と、全段の PICTURE の太枠（カットが入らない段にも描く）。"""
    layout.draw_grid(c, page_no)
    c.setLineWidth(2.5)
    for i in range(layout.per_page):
        x, top, w, h = layout.slot(i)["picture"]
        c.rect(x, y(c, top + h), w, h)
    c.setLineWidth(1)


def y(c, top):
    """上端基準の座標を reportlab の下端基準に変換する。"""
    return c._pagesize[1] - top


# 文字欄: (基本の文字サイズ, 中央揃えか)。収まらなければ MIN_FONT まで小さくする
TEXT_STYLE = {"action": (8, False), "scenario": (8, False), "time": (9, True), "note": (7, False)}
MIN_FONT = 5


def fit_text(c, markup, rect, size, center):
    """矩形に収まる大きさで文字を描く。最小サイズでも収まらなければ、はみ出す分を切って False を返す。

    改行ごとに別の段落にして積み重ねる（1つの段落に <br/> を入れると、行末がちょうど幅いっぱいのとき
    reportlab が空行を余分に入れてしまうため）。
    """
    x, top, rw, rh = rect
    pad = 3
    w, h = rw - pad * 2, rh - pad * 2
    lines = markup.split("<br/>")
    for fs in [size - 0.5 * i for i in range(int((size - MIN_FONT) * 2) + 1)]:
        style = ParagraphStyle("t", fontName=JP, fontSize=fs, leading=fs * 1.35, wordWrap="CJK",
                               alignment=TA_CENTER if center else 0)
        paras = [Paragraph(t or "&nbsp;", style) for t in lines]
        heights = [p.wrap(w, h)[1] for p in paras]
        if sum(heights) <= h:
            break
    fits = sum(heights) <= h
    c.saveState()
    if not fits:  # 隣の欄にはみ出さないよう、欄の範囲で切る
        path = c.beginPath()
        path.rect(x, y(c, top + rh), rw, rh)
        c.clipPath(path, stroke=0, fill=0)
    cur = top + pad
    for p, ph in zip(paras, heights):
        cur += ph
        p.drawOn(c, x + pad, y(c, cur))
    c.restoreState()
    return fits


def draw_cut(c, layout, i, cut, texts):
    """texts: {"action"|"scenario"|"time"|"note": 段落用の文字}。収まらなかった欄の名前を返す。"""
    s = layout.slot(i)
    c.setFont(JP, 9)
    for key in ("s", "c"):
        x, top, w, h = s[key]
        c.drawCentredString(x + w / 2, y(c, top + 14), cut[key])
    x, top, w, h = s["picture"]
    pad = 2
    if cut["image"] is not None:
        img = cut["image"] if isinstance(cut["image"], str) else ImageReader(cut["image"])
        c.drawImage(img, x + pad, y(c, top + h - pad), w - pad * 2, h - pad * 2, preserveAspectRatio=True, anchor="c")
    else:
        c.setFont(JP, 8)
        c.setFillGray(0.6)
        c.drawCentredString(x + w / 2, y(c, top + h / 2) - 3, "（カット絵なし）")
        c.setFillGray(0)
    overflow = []
    for key, (size, center) in TEXT_STYLE.items():
        if texts.get(key) and not fit_text(c, texts[key], s[key], size, center):
            overflow.append(key)
    return overflow


def check_cuts(picture_names, excel_path):
    """カット絵とカット表の突き合わせ → (カット絵がないカット, カット表にないカット)。カット表のシートがなければ ([], [])"""
    table = load_book(excel_path).cuts
    if not table:
        return [], []
    table_names = {f"{s}-{c}" for s, c in table}
    pics = set(picture_names)
    order = lambda n: make_cut(n)["key"]
    return sorted(table_names - pics, key=order), sorted(pics - table_names, key=order)


COL_LABEL = {"action": "ACTION/SE", "scenario": "SCENARIO", "time": "TIME", "note": "NOTE"}


def compose(cuts, out, excel_path=None, fmt="横", kouban_out=None):
    """絵コンテPDFを書き出す。

    cuts: make_cut() の結果のリスト（カット絵）。カット表にあってカット絵がないカットも、PICTURE を空けて載せる。
    excel_path: 香盤表・カット表のExcel（1つのファイル。省略可）
    fmt: "横"（1ページ6カット）/ "縦"（9:16、1ページ4カット）。
    kouban_out: 指定すると、同じ中身の香盤表（Excel）も書き出す
    番号順に並べ、シーンごとに改ページして配置する。
    戻り値: {"pages": ページ数, "cuts": カット数, "warnings": [...]}
    """
    book = load_book(excel_path) if excel_path else Book()
    rows = plan(cuts, book)
    if not rows:
        raise ValueError("カットがありません")
    warnings = []
    layout = LAYOUTS[fmt]
    c = canvas.Canvas(out, pagesize=layout.pagesize)
    page, slot = 0, layout.per_page
    for row in rows:
        # シーンが変わったら、段が余っていても次のページから始める
        if row.first or slot == layout.per_page:
            if page:
                c.showPage()
            page += 1
            slot = 0
            draw_page(c, layout, page)
        texts = {k: "<br/>".join(to_markup(v) for v in getattr(row, k)) for k in ("action", "scenario")}
        texts["time"] = to_markup(row.time) if row.time else ""
        texts["note"] = note_markup(row.note)
        cut = row.cut
        for key in draw_cut(c, layout, slot, cut, texts):
            warnings.append(f"{cut['s']}-{cut['c']} の {COL_LABEL[key]} が長すぎて欄に収まりません（No.{page}）")
        slot += 1
    c.save()
    if kouban_out:
        from .kouban import write
        write(kouban_out, rows, book, excel_path)
    return {"pages": page, "cuts": len(rows), "warnings": warnings}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("out")
    ap.add_argument("--excel", help="香盤表・カット表（1つのファイル）")
    ap.add_argument("--kouban", help="一緒に書き出す香盤表（.xlsx）")
    ap.add_argument("--format", choices=list(LAYOUTS), default="横", help="横（1ページ6カット）/ 縦（9:16、4カット）")
    a = ap.parse_args()
    cuts = collect_cuts(a.folder)
    if a.excel:
        missing, extra = check_cuts([f"{x['s']}-{x['c']}" for x in cuts], a.excel)
        if missing:
            print("カット絵がないカット:", ", ".join(missing))
        if extra:
            print("カット表にないカット:", ", ".join(extra))
    r = compose(cuts, a.out, a.excel, a.format, a.kouban)
    for w in r["warnings"]:
        print("⚠", w)
    print(f"{r['cuts']}カット / {r['pages']}ページ → {a.out}")
