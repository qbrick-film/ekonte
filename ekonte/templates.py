"""書き込む Excel（香盤表・カット表・記入方法の3枚のシート）を作る。

python -m ekonte.templates 香盤表・カット表.xlsx [空]

香盤表は部の香盤表（香盤表_テンプレート.xlsx）をそのまま使い、カット表と記入方法のシートを足す。
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from . import samples

COLUMNS = [  # 見出し, 幅, 絵コンテのどこに入るか
    ("S", 5, "S"), ("C", 6, "C"),
    ("ACTION", 26, "ACTION/SE"), ("SE", 14, "ACTION/SE"), ("SCENARIO", 34, "SCENARIO"), ("TIME", 7, "TIME"),
    ("小道具", 16, "NOTE"), ("カメラA", 11, "NOTE"), ("カメラB", 11, "NOTE"), ("機材", 14, "NOTE"),
    ("NOTE", 24, "NOTE"),
]
CENTER_COLS = (1, 2, 6)
TIME_COL = "F"

# 記入例（S, C, ACTION, SE, SCENARIO, TIME, 小道具, カメラA, カメラB, 機材, NOTE）
EXAMPLE = [
    (1, 1, "公園の全景", None, None, 4, None, "24mm", None, None, None),
    (None, None, "2人が歩いてくる", None, None, None, None, None, None, None, None),
    (1, 2, "2人のミディアム", None, "花子「ねえ、聞いた？」", 5, None, "50mm", None, None, None),
    (None, None, "花子が振り向く", None, "太郎「何を？」", None, None, None, None, None, None),
    (2, None, None, None, None, None, "拳銃", "35mm", None, "三脚", None),
    (2, 1, "倉庫の奥で男が拳銃を構える", "雨音", None, 3, None, None, None, None, None),
    (2, 2, "拳銃のアップ", None, None, 2, None, None, None, None, None),
    (2, 3, "車が入ってくる（PAN →）", "エンジン音", "男「来たか」", 4, "車", "85mm", None, "ドリー", None),
    (2, "3a", "ヘッドライトのインサート", None, None, 1.5, None, None, None, None, "追加カット"),
    (3, None, None, None, None, None, None, "24mm", "50mm", "車載リグ", "2台撮り"),
    (3, 1, "車内。後部座席から2人の後頭部", None, "太郎「どこへ行く？」", 4, None, None, None, None, None),
    (3, 2, "運転席の顔アップ", None, "男「……」", 3, None, None, None, None, None),
]
# 記入例の香盤表の登場人物と、出るシーン
CAST = [("花子", ("1",)), ("太郎", ("1", "3")), ("男", ("2", "3"))]

HELP = [
    ("香盤表・カット表の書き方", "title"),
    ("", None),
    ("■ この Excel には「香盤表」と「カット表」の2枚のシートがあります", "head"),
    ("  香盤表 … 1行1シーン。場面・L/LS・D/N・ロケ地・登場人物・備考（部の香盤表と同じ形）", None),
    ("  カット表 … 1行1カット。ACTION・SE・SCENARIO・TIME・小道具・カメラ・機材・NOTE", None),
    ("  ソフトの ③ でこのファイルを選ぶと、両方を読み込みます。どちらか一方だけ書いてもかまいません。", None),
    ("  ソフトは絵コンテ（PDF）と一緒に、香盤表（Excel）も書き出します。シーンごとのカット数・時間・小道具・カメラ・機材を", None),
    ("  カット表から自動でまとめるので、同じことを2か所に書く必要はありません。", None),
    ("", None),
    ("■ 小道具・カメラ・機材はカット表に書きます", "head"),
    ("  香盤表の備考にも同じものを書いた場合は、絵コンテ・書き出す香盤表には1回だけ載ります（「、」で区切って書いたもの）。", None),
    ("", None),
    ("■ 香盤表の書き方", "head"),
    ("  #S にシーン番号を書き、場面・L/LS・D/N・ロケ地・備考を書きます。", None),
    ("  登場人物の列は、見出しに役の名前（例: 花子）、その上の段に役者の名前を書き、出るシーンに ○ を付けます。", None),
    ("  場面・L/LS・D/N・ロケ地はシーンの最初のカットの NOTE に、備考はシーンの全カットの NOTE に入ります。", None),
    ("  別のファイルの香盤表を使うときは、そのシートをこのファイルにコピーします。", None),
    ("    Excel: シートの見出しを右クリック →「移動またはコピー」→ 移動先にこのファイルを選び、「コピーを作成する」にチェック", None),
    ("    Googleスプレッドシート: シートのタブの ▼ →「別のワークブックにコピー」→ このファイル", None),
    ("    コピーしたシートの名前は「香盤表」にします（元の「香盤表」のシートは消します）。", None),
    ("", None),
    ("■ カット表：1行 = 1カット（白い行）", "head"),
    ("  S（シーン）と C（カット）を書きます。追加カットは C に 3a のように書きます。", None),
    ("  カット絵のマークシートの番号と同じにしてください。", None),
    ("", None),
    ("■ カット表：S も C も空欄の行 = すぐ上の行の続き（灰色の行）", "head"),
    ("  2行目以降のセリフや ACTION は、下の行に書き足していきます。各欄に1行ずつ追記されます。", None),
    ("  セル内で改行（Windows: Alt+Enter / Mac: option+return / Googleスプレッドシート: Ctrl+Enter）しても同じです。", None),
    ("", None),
    ("■ カット表：S だけ書いて C が空欄の行 = シーン全体（青い行）", "head"),
    ("  小道具・カメラ・機材・NOTE を、そのシーンの全カットに入れます。", None),
    ("  カットの行にも書いた場合: 小道具・機材・NOTE は追加、カメラは置き換え。", None),
    ("  追加カット（2-3a）は 2-3 の内容を引き継ぎません。", None),
    ("", None),
    ("■ カット表の各列の入れ先", "head"),
    ("  ACTION → ACTION/SE 欄       芝居・カメラワーク", None),
    ("  SE → ACTION/SE 欄           効果音。「SE：」は付けなくても自動で付きます（例: 雨音 → SE：雨音）", None),
    ("  SCENARIO → SCENARIO 欄      セリフ", None),
    ("  TIME → TIME 欄              秒数（数字）", None),
    ("  小道具・カメラA・カメラB・機材・NOTE → NOTE 欄", None),
    ("  カメラA/B にはレンズを書きます（例: 35mm）。1台なら カメラA だけ。列を増やすときは「カメラC」のように見出しを付けます。", None),
    ("  カット表の2行目（→ 〜）は説明なので消さないでください。3行目から書きます。", None),
    ("", None),
    ("■ Googleスプレッドシートで書く場合", "head"),
    ("  このファイルを Googleドライブに上げて「Googleスプレッドシートで開く」で使えます（シートは2枚とも入ります）。", None),
    ("  ソフトに読み込むときは「ファイル → ダウンロード → Microsoft Excel (.xlsx)」で保存したファイルを選びます。", None),
]


def _example_kouban(ws):
    """記入例の香盤表に、タイトルと登場人物（出るシーンに ○）を書き入れる。"""
    cells = [c for row in ws.iter_rows(max_row=30) for c in row if c.value is not None]
    head = next(c for c in cells if c.value == "#S")
    label = next(c for c in cells if c.value == "タイトル")
    merged = next((m for m in ws.merged_cells.ranges if label.coordinate in m), None)
    ws.cell(label.row, (merged.max_col if merged else label.column) + 1, "記入例")
    rows = {str(ws.cell(r, head.column).value): r for r in range(head.row + 1, ws.max_row + 1)}
    roles = [c for c in ws[head.row] if str(c.value or "").startswith("登場人物")]
    for c, (name, scenes) in zip(roles, CAST):
        c.value = name
        for s in scenes:
            ws.cell(rows[s], c.column, "○")


def build(path, example=True):
    """香盤表・カット表・記入方法の3枚のシートの Excel を作る。example: 記入例を入れる"""
    wb = openpyxl.load_workbook(samples.path_of(samples.KOUBAN_BASE[example]))
    kouban = wb.worksheets[0]
    kouban.title = "香盤表"
    if example:
        _example_kouban(kouban)
    ws = wb.create_sheet("カット表")
    thin = Side(style="thin", color="999999")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    fills = {
        "head": PatternFill("solid", fgColor="DDDDDD"),
        "sub": PatternFill("solid", fgColor="F2F2F2"),
        "scene": PatternFill("solid", fgColor="EAF1FB"),
        "cont": PatternFill("solid", fgColor="F5F5F5"),
    }
    for i, (h, w, dest) in enumerate(COLUMNS, 1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True)
        c.fill = fills["head"]
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        d = ws.cell(2, i, f"→ {dest}")
        d.font = Font(size=8, color="666666")
        d.fill = fills["sub"]
        d.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = w

    rows = EXAMPLE if example else []
    for r in range(3, 503):
        row = rows[r - 3] if r - 3 < len(rows) else None
        for i in range(1, len(COLUMNS) + 1):
            c = ws.cell(r, i, row[i - 1] if row else None)
            c.border = border
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if i in CENTER_COLS else "left")
            if row and row[0] is not None and row[1] is None:
                c.fill = fills["scene"]
            elif row and row[0] is None:
                c.fill = fills["cont"]
    ws.freeze_panes = "C3"

    dv_s = DataValidation(type="whole", operator="between", formula1="1", formula2="999", allow_blank=True,
                          showErrorMessage=True, errorTitle="S", error="シーン番号は 1〜999 の整数で入力してください")
    dv_t = DataValidation(type="decimal", operator="between", formula1="0", formula2="600", allow_blank=True,
                          showErrorMessage=True, errorTitle="TIME", error="秒数を数字で入力してください（例: 3 / 1.5）")
    ws.add_data_validation(dv_s)
    ws.add_data_validation(dv_t)
    dv_s.add("A3:A2000")
    dv_t.add(f"{TIME_COL}3:{TIME_COL}2000")

    hs = wb.create_sheet("記入方法")
    for i, (t, kind) in enumerate(HELP, 1):
        hs.cell(i, 1, t).font = Font(bold=kind is not None, size=13 if kind == "title" else 11)
    hs.column_dimensions["A"].width = 120
    wb.active = 0
    wb.save(path)


if __name__ == "__main__":
    build(sys.argv[1], example="空" not in sys.argv[2:])
