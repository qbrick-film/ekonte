"""絵コンテと一緒に書き出す香盤表（Excel）。絵コンテと同じ中身（カット絵以外すべて）を2枚のシートにする。

  香盤表     1行1シーン。元の香盤表（場面・L/S・D/N・ロケ地・登場人物など・備考）に、
             カット表から集めたカット数・時間・小道具・カメラ・機材を足す。最後の行に合計
  カット一覧 1行1カット。絵コンテから PICTURE を除いたもの（S・C・ACTION/SE・SCENARIO・TIME・NOTE）

中身は compose.plan() の結果から作るので、絵コンテPDFと食い違わない。
直すときは元の Excel（香盤表・カット表）を直して書き出し直す（このファイルは読み込めないようにしてある）。
"""
import os, re
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from .compose import without
from .excel_import import OUTPUT_MARK

THIN = Side(style="thin", color="999999")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
SCENE_TOP = Border(top=Side(style="medium", color="555555"), bottom=THIN, left=THIN, right=THIN)  # シーンの区切り
HEAD_FILL = PatternFill("solid", fgColor="DDDDDD")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
LEFT = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(wrap_text=True, vertical="top", horizontal="center")
GATHER = [("props", "小道具", 24), ("camera", "カメラ", 18), ("gear", "機材", 18)]  # カット表からシーンごとにまとめるもの


def output_path(pdf_path, source=None):
    """絵コンテPDFの隣に置く香盤表の名前（絵コンテ.pdf → 絵コンテ_香盤表.xlsx）。元の Excel と同じ名前にはしない。"""
    base = os.path.splitext(pdf_path)[0] + "_香盤表"
    if source and os.path.abspath(base + ".xlsx") == os.path.abspath(source):
        base += "_書き出し"
    return base + ".xlsx"


def write(path, rows, book, source=None):
    """rows: compose.plan() の結果。book: 読み込んだ香盤表・カット表。source: 元の Excel の場所（ファイル名を書き添える）"""
    wb = Workbook()
    _scenes(wb.active, rows, book, source)
    _cuts(wb.create_sheet("カット一覧"), rows)
    wb.properties.subject = OUTPUT_MARK
    wb.properties.creator = "絵コンテ作成ソフト"
    wb.save(path)


def _number(v):
    """数字だけの文字は数として書く（Excel で合計などに使えるように）。"""
    if isinstance(v, str) and re.fullmatch(r"\d+(\.\d+)?", v):
        f = float(v)
        return int(f) if f.is_integer() else f
    return v


def _seconds(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _scene_order(s):
    m = re.match(r"(\d+)(.*)", s)
    return (int(m.group(1)), m.group(2)) if m else (float("inf"), s)


def _cell(ws, row, col, value, align=LEFT, border=BORDER, **font):
    c = ws.cell(row, col, value)
    c.alignment, c.border = align, border
    if font:
        c.font = Font(**font)
    return c


def _head(ws, row, columns):
    for i, (name, width) in enumerate(columns, 1):
        c = _cell(ws, row, i, name, CENTER, bold=True)
        c.fill = HEAD_FILL
        ws.column_dimensions[get_column_letter(i)].width = width


def _print_setup(ws, title_rows):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = title_rows


def _scenes(ws, rows, book, source):
    ws.title = "香盤表"
    gathered = {}  # S → {"cuts": カット数, "time": 秒の合計, "items": {キー: [...]}}
    for r in rows:
        g = gathered.setdefault(r.cut["s"], {"cuts": 0, "time": None, "items": {k: [] for k, _, _ in GATHER}})
        g["cuts"] += 1
        t = _seconds(r.time)
        if t is not None:
            g["time"] = (g["time"] or 0) + t
        for k, _, _ in GATHER:
            g["items"][k] += r.items.get(k, [])
    extra = [x for x in book.extra if x[2]]  # 香盤表のそのほかの列（登場人物など）。どのシーンにも書いていない列は省く
    scenes = set(gathered) | set(book.kouban) | {s for _, _, values in extra for s in values}

    ws["A1"] = "香盤表"
    ws["A1"].font = Font(bold=True, size=16)
    for row, (label, value) in enumerate((("タイトル", book.title), ("撮影日", book.date)), 2):
        ws.cell(row, 1, label).font = Font(bold=True)
        ws.cell(row, 3, value)
    note = "絵コンテ作成ソフトで作成" + (f"（元の Excel: {os.path.basename(source)}）" if source else "")
    ws.cell(4, 1, note + "。直すときは元の Excel を直して、もう一度書き出してください").font = Font(size=9, color="666666")

    head = 6  # 見出しの行（すぐ上の行に役者名など）
    columns = [("#S", 5), ("場面", 20), ("L/S", 6), ("D/N", 6), ("ロケ地", 18)]
    columns += [(name, 8) for name, _, _ in extra]
    columns += [("カット数", 7), ("時間（秒）", 8)] + [(label, width) for _, label, width in GATHER] + [("備考", 36)]
    _head(ws, head, columns)
    for j, (_, above, _) in enumerate(extra):
        if above:
            _cell(ws, head - 1, 6 + j, above, CENTER, border=Border(), size=9)
    centered = {1, 3, 4} | set(range(6, 6 + len(extra) + 2))

    r, total_cuts, total_time = head + 1, 0, None
    for s in sorted(scenes, key=_scene_order):
        info = book.kouban.get(s, {})
        g = gathered.get(s)
        if g:
            items = g["items"]
        else:  # カットのないシーンでも、シーン全体に書いた小道具などは載せる
            items = {k: book.sources.get(k, ({}, {}))[0].get(s, []) for k, _, _ in GATHER}
        items = {k: list(dict.fromkeys(v)) for k, v in items.items()}
        shown = {x for v in items.values() for x in v}
        values = [_number(s)] + [info.get(k, "") for k in ("bamen", "ls", "dn", "loca")]
        values += [v.get(s, "") for _, _, v in extra]
        values += [g["cuts"] if g else 0, g["time"] if g else None]
        values += ["、".join(items[k]) for k, _, _ in GATHER]
        values.append(without(info.get("biko", ""), shown))  # 小道具などの欄と同じものは備考に重ねない
        for i, v in enumerate(values, 1):
            _cell(ws, r, i, v, CENTER if i in centered else LEFT)
        total_cuts += g["cuts"] if g else 0
        if g and g["time"] is not None:
            total_time = (total_time or 0) + g["time"]
        r += 1
    cut_col = 6 + len(extra)
    for i in range(1, len(columns) + 1):
        _cell(ws, r, i, None, CENTER).fill = TOTAL_FILL
    ws.cell(r, 2, "合計").font = Font(bold=True)
    ws.cell(r, cut_col, total_cuts).font = Font(bold=True)
    ws.cell(r, cut_col + 1, total_time).font = Font(bold=True)
    ws.freeze_panes = ws.cell(head + 1, 2)
    _print_setup(ws, f"{head}:{head}")


def _note(lines):
    """NOTE欄の行を1つのセルに。シーンの情報の行は絵コンテと同じく太字にする。"""
    if not any(bold for _, bold in lines):
        return "\n".join(t for t, _ in lines)
    rich = CellRichText()
    for i, (text, bold) in enumerate(lines):
        text = ("\n" if i else "") + text
        rich.append(TextBlock(InlineFont(b=True), text) if bold else text)
    return rich


def _cuts(ws, rows):
    columns = [("S", 5), ("C", 6), ("ACTION/SE", 40), ("SCENARIO", 40), ("TIME", 7), ("NOTE", 48)]
    _head(ws, 1, columns)
    total = None
    for i, row in enumerate(rows, 2):
        cut = row.cut
        values = [_number(cut["s"]), _number(cut["c"]), "\n".join(row.action), "\n".join(row.scenario),
                  _number(row.time) if row.time else None, _note(row.note)]
        for j, v in enumerate(values, 1):
            _cell(ws, i, j, v, CENTER if j in (1, 2, 5) else LEFT, SCENE_TOP if row.first else BORDER)
        t = _seconds(row.time)
        if t is not None:
            total = (total or 0) + t
    r = len(rows) + 2
    for j in range(1, len(columns) + 1):
        _cell(ws, r, j, None, CENTER).fill = TOTAL_FILL
    ws.cell(r, 3, "合計").font = Font(bold=True)
    ws.cell(r, 5, total).font = Font(bold=True)
    ws.freeze_panes = "C2"
    _print_setup(ws, "1:1")
