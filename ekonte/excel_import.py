"""香盤表・カット表のExcelを読み込む。

1つのファイルに「香盤表」と「カット表」のシートを入れる（どちらか一方だけでもよい）。
シートは名前で探し、見つからなければ見出しで見分ける（C の列があればカット表、場面・D/N などがあれば香盤表）。
見出し行は列名で自動検出するので、列の並びや開始行が変わっても読める。
"""
import datetime, re, zipfile
from collections import defaultdict
from dataclasses import dataclass, field
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

SCENE_HEADERS = ("#S", "S#", "シーン", "S")
CUT_HEADERS = ("カット", "C", "C#")
KOUBAN_HEADERS = ("場面", "L/S", "L/LS", "D/N", "ロケ地")  # 香盤表のシートを見分けるのに使う列
OUTPUT_MARK = "絵コンテ作成ソフトで書き出した香盤表"   # 書き出した香盤表（kouban.py）の「件名」。読み込もうとしたら知らせる


@dataclass
class Book:
    """読み込んだ香盤表・カット表。"""
    kouban: dict = field(default_factory=dict)   # 香盤表 {S: {bamen, ls, dn, loca, biko}}
    cuts: dict = field(default_factory=dict)     # カット表 {(S, C): {"action": [...], "scenario": [...], "time": [...]}}
    sources: dict = field(default_factory=dict)  # NOTE の小道具・カメラ・機材など {キー: (シーン単位 {S: [..]}, カット単位 {(S, C): [..]})}
    title: str = ""                              # 香盤表のタイトル・撮影日（書き出す香盤表に引き継ぐ）
    date: str = ""
    extra: list = field(default_factory=list)    # 香盤表のそのほかの列（登場人物など）: [(見出し, すぐ上の段の文字（役者名など）, {S: 値})]
    has_kouban: bool = False
    has_cuts: bool = False


def _open(path):
    """Excelを開く。開けなければ、理由が分かる日本語のエラーにする。"""
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except (zipfile.BadZipFile, InvalidFileException) as e:
        raise ValueError("Excel（.xlsx）として開けません。古い形式（.xls）などのファイルなら、"
                         "Excel で .xlsx 形式で保存し直してから選んでください") from e


def _norm(v):
    return re.sub(r"\s", "", str(v)) if v is not None else ""


def _text(v):
    """セルの値を文字にする（3.0 → "3"、日付 → "2026/10/12"）。"""
    if v is None:
        return ""
    if isinstance(v, (datetime.date, datetime.datetime)):
        return f"{v.year}/{v.month}/{v.day}"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _num(v):
    """シーン/カット番号を文字列に正規化する（2, 2.0, '２' → '2'）。"""
    if v is None or _norm(v) == "":
        return None
    s = _norm(v).translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if s.isdigit():
        return str(int(s))
    m = re.fullmatch(r"0*(\d+)([A-Ja-j])", s)  # 追加カット "03A" → "3a"
    if m:
        return str(int(m.group(1))) + m.group(2).lower()
    return s


def _find_header(ws, required):
    """required のいずれかを含む行を見出し行とし、{見出し: 列番号} を返す。"""
    for row in ws.iter_rows(min_row=1, max_row=30):
        names = {_norm(c.value): c.column for c in row if c.value is not None}
        if any(h in names for h in required):
            return row[0].row, names
    raise ValueError(f"「{ws.title}」のシートに見出し行が見つかりません（{'/'.join(required)} の列が必要）")


def _col(names, *candidates, prefix=False):
    for c in candidates:
        for name, idx in names.items():
            if name == c or (prefix and name.startswith(c)):
                return idx
    return None


def _find_sheets(wb):
    """(香盤表のシート, カット表のシート)。名前で探し、なければ見出しで見分ける。"""
    kouban = next((ws for ws in wb.worksheets if "香盤" in ws.title), None)
    cuts = next((ws for ws in wb.worksheets if "カット表" in ws.title), None)
    for ws in wb.worksheets:
        if ws in (kouban, cuts) or "記入方法" in ws.title or "カット一覧" in ws.title:
            continue
        try:
            _, names = _find_header(ws, SCENE_HEADERS)
        except ValueError:
            continue
        if cuts is None and _col(names, *CUT_HEADERS):
            cuts = ws
        elif kouban is None and _col(names, *KOUBAN_HEADERS):
            kouban = ws
    return kouban, cuts


def load_book(path):
    """香盤表とカット表をまとめたExcelを読む。→ Book"""
    wb = _open(path)
    if wb.properties.subject == OUTPUT_MARK:
        raise ValueError("これはソフトが書き出した香盤表です。書いた元の Excel（香盤表・カット表）を選んでください")
    kouban, cuts = _find_sheets(wb)
    if kouban is None and cuts is None:
        raise ValueError("香盤表・カット表のシートが見つかりません。シート名を「香盤表」「カット表」にしてください")
    book = Book()
    if cuts is not None:
        book.cuts, book.sources = _read_cut_table(cuts)
        book.has_cuts = True
    if kouban is not None:
        _read_kouban(kouban, book)
    return book


def describe_book(book):
    """画面に出す要約（「香盤表 3シーン・カット表 8カット」など）"""
    parts = []
    if book.has_kouban:
        parts.append(f"香盤表 {len(book.kouban)}シーン")
    if book.has_cuts:
        parts.append(f"カット表 {len(book.cuts)}カット")
    return "・".join(parts)


def _read_kouban(ws, book):
    """香盤表のシート → book.kouban（NOTE に入れるシーンの情報）と、書き出す香盤表に引き継ぐもの"""
    header_row, names = _find_header(ws, SCENE_HEADERS)
    cols = dict(
        scene=_col(names, *SCENE_HEADERS),
        bamen=_col(names, "場面"),
        ls=_col(names, "L/S", "L/LS"),
        dn=_col(names, "D/N"),
        loca=_col(names, "ロケ地"),
        biko=_col(names, "備考", prefix=True),
    )
    # 香盤表に小道具・機材の列があれば、カット表のシーン全体の行と同じ扱いにして1つにまとめる（二重に載せない）
    shared = {k: _col(names, label, prefix=True) for k, label in (("props", "小道具"), ("gear", "機材"))}
    # そのほかの列（登場人物など）は、書き出す香盤表にそのまま引き継ぐ。書き出す香盤表が足す列（カット数・時間）は読まない
    known = set(cols.values()) | set(shared.values())
    above = header_row - 1
    extra = [(idx, _text(ws.cell(header_row, idx).value), _text(ws.cell(above, idx).value) if above else "")
             for name, idx in sorted(names.items(), key=lambda x: x[1])
             if idx not in known and name not in ("カット数", "時間（秒）")]
    values = {idx: {} for idx, _, _ in extra}
    for row in ws.iter_rows(min_row=header_row + 1):
        get = lambda idx: row[idx - 1].value if idx and idx <= len(row) else None
        s = _num(get(cols["scene"]))
        if s is None:
            continue
        info = {k: _text(get(cols[k])) for k in ("bamen", "ls", "dn", "loca", "biko")}
        if any(info.values()):
            book.kouban[s] = info
        for k, idx in shared.items():
            if idx and _text(get(idx)):
                by_scene, _ = book.sources.setdefault(k, (defaultdict(list), defaultdict(list)))
                by_scene[s].append(_text(get(idx)))
        for idx, _, _ in extra:
            if _text(get(idx)):
                values[idx][s] = _text(get(idx))
    book.extra = [(head, top, values[idx]) for idx, head, top in extra]
    # 見出しより上の「タイトル」「撮影日」の欄の右に書いたもの
    for row in ws.iter_rows(min_row=1, max_row=max(above, 1)):
        for i, c in enumerate(row):
            label = _norm(c.value)
            if label in ("タイトル", "撮影日"):
                value = next((_text(x.value) for x in row[i + 1:] if _text(x.value)), "")
                if label == "タイトル":
                    book.title = value
                else:
                    book.date = value
    book.has_kouban = True


CUT_FIELDS = {  # キー: 見出しの候補（前方一致）
    "action": ("ACTION/SE", "ACTION", "芝居"),
    "se": ("SE", "効果音"),  # 絵コンテでは ACTION/SE 欄に「SE：〜」として入れる
    "scenario": ("SCENARIO", "セリフ", "台詞"),
    "time": ("TIME", "秒"),
    "props": ("小道具",),
    "gear": ("機材", "特機"),
    "note": ("NOTE", "メモ"),
}


def _read_cut_table(ws):
    """カット表のシート → (カットの一覧, NOTE用のデータ)

    1行 = 1カット。書き方のルール:
      S と C を書いた行          … そのカット
      S だけ書いて C が空欄の行  … シーン全体（小道具・カメラ・機材・NOTE をそのシーンの全カットに入れる）
      S も C も空欄の行          … すぐ上の行の続き（各欄に1行ずつ追記する）
    同じ S・C を2回書いた場合も追記する。

    戻り値:
      cuts:    {(S, C): {"action": [...], "scenario": [...], "time": [...]}}  表に書いた順
      sources: compose の NOTE_ITEMS と同じ形 {キー: (シーン単位 {S: [..]}, カット単位 {(S, C): [..]})}
    """
    header_row, names = _find_header(ws, SCENE_HEADERS)
    c_scene = _col(names, *SCENE_HEADERS)
    c_cut = _col(names, *CUT_HEADERS)
    if c_cut is None:
        raise ValueError("C（カット）の列が見つかりません")
    cols = {k: _col(names, *cands, prefix=True) for k, cands in CUT_FIELDS.items()}
    # カメラは「カメラA」「カメラB」… の列。値はレンズ（例: 35mm）
    cams = [(name[len("カメラ"):], idx) for name, idx in names.items() if name.startswith("カメラ")]

    cuts = {}
    sources = {k: (defaultdict(list), defaultdict(list)) for k in ("props", "gear", "note", "camera")}
    target = None  # 直前の行の対象: ("scene", S) / ("cut", (S, C))
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        cell = lambda idx: row[idx - 1] if idx and idx <= len(row) else None
        raw_s, raw_c = cell(c_scene), cell(c_cut)
        s, c = _num(raw_s), _num(raw_c)
        if s is not None and not s.isdigit():
            continue  # 「→ S」などの説明行
        values = {k: str(cell(idx)).strip() for k, idx in cols.items() if cell(idx) not in (None, "")}
        values = {k: v for k, v in values.items() if v}
        cam_values = [(label, str(cell(idx)).strip()) for label, idx in cams if cell(idx) not in (None, "")]
        if s is None and c is None:
            if not values and not cam_values:
                continue  # 空行
            if target is None:
                continue  # 先頭の続き行は対象がないので無視
        elif s is not None and c is None:
            target = ("scene", s)
        elif s is not None:
            target = ("cut", (s, c))
            cuts.setdefault((s, c), {"action": [], "scenario": [], "time": []})
        else:
            continue  # C だけ書いた行は不正なので無視
        kind, key = target
        for k, v in values.items():  # CUT_FIELDS の順なので、同じ行なら ACTION → SE の順に入る
            if k == "se":
                if kind == "cut":
                    cuts[key]["action"].append("SE：" + re.sub(r"^SE\s*[：:]\s*", "", v, flags=re.I))
            elif k in ("action", "scenario", "time"):
                if kind == "cut":
                    cuts[key][k].append(_fmt_time(v) if k == "time" else v)
            else:
                (sources[k][1][key] if kind == "cut" else sources[k][0][key]).append(v)
        for label, lens in cam_values:
            item = f"{label} {lens}".strip()
            (sources["camera"][1][key] if kind == "cut" else sources["camera"][0][key]).append(item)
    return cuts, sources


def _fmt_time(v):
    """3.0 → "3"、1.5 → "1.5"。数字でなければそのまま。"""
    try:
        f = float(v)
    except ValueError:
        return v
    return str(int(f)) if f == int(f) else f"{f:g}"
