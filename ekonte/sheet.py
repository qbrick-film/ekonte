"""カット絵用紙PDFを生成する。  python -m ekonte.sheet out.pdf [ページ数] [横|縦]"""
import sys
from reportlab.pdfgen import canvas
from .layout import *
from .fonts import JP

LEFT = 30  # 見出し・注意書きの左端


def draw_page(c, fmt="横"):
    # 四隅の位置合わせマークと向き判定ドット（縦の用紙には縦型の印も）
    for cx, cy in CORNER_CENTERS.values():
        c.rect(cx - CORNER / 2, cy - CORNER / 2, CORNER, CORNER, stroke=0, fill=1)
    for dot in (ORIENT_DOT, FORMAT_DOT) if fmt == "縦" else (ORIENT_DOT,):
        c.circle(dot["x"], dot["y"], dot["r"], stroke=0, fill=1)

    title = "カット絵" if fmt == "横" else f"カット絵（{FORMATS[fmt]}）"
    c.setFont(JP, 14)
    c.drawString(LEFT, 545, title)
    c.setFont(JP, 9)
    c.drawString(LEFT + c.stringWidth(title, JP, 14) + 14, 547, "※ 右のマーク欄を塗りつぶして S（シーン）と C（カット）を指定してください")

    # 描画枠
    f = FRAMES[fmt]
    c.setLineWidth(2.5)
    c.rect(f["x"], f["y"], f["w"], f["h"])
    c.setLineWidth(1)

    # マーク欄の見出し
    s_cols, c_cols = COLUMNS[0:3], COLUMNS[3:5]
    c.setFont(JP, 11)
    c.drawCentredString((s_cols[0][2] + s_cols[-1][2]) / 2, 535, "S")
    c.drawCentredString((c_cols[0][2] + c_cols[-1][2]) / 2, 535, "C")
    c.drawCentredString(COLUMNS[5][2], 535, "枝")
    c.drawCentredString(710, 507, "-")

    for key, label, x, values in COLUMNS:
        # 手書き確認用のマス
        c.rect(x - 11, 498, 22, 26)
        c.setFont(JP, 7)
        c.setFillGray(0.4)
        c.drawCentredString(x, 488, label)
        c.setFillGray(0)
        for row, v in enumerate(values):
            bx, by = bubble_center(x, row)
            c.setStrokeGray(0.35)
            c.circle(bx, by, BUBBLE_R, stroke=1, fill=0)
            c.setFillGray(0.6)
            c.setFont(JP, 7)
            c.drawCentredString(bx, by - 2.5, v)
            c.setFillGray(0)
            c.setStrokeGray(0)

    c.setFont(JP, 7)
    c.setFillGray(0.4)
    c.drawString(615, 228, "例）2-3   → S一=2, C一=3")
    c.drawString(615, 216, "例）12-5a → S十=1, S一=2, C一=5, 枝=a")
    c.drawString(615, 204, "枝は追加カットのときのみ")
    c.drawString(LEFT, 44, "印刷するときは A4横・実際のサイズ（100%）で。「用紙に合わせる」にすると縮小されます")
    c.setFillGray(0)

def build(path, pages=1, fmt="横"):
    c = canvas.Canvas(path, pagesize=(PAGE_W, PAGE_H))
    # 印刷時に「用紙に合わせる」で縮小されないよう、実際のサイズでの印刷を指定する
    c.setViewerPreference("PrintScaling", "None")
    c.setTitle(f"カット絵用紙（{FORMATS[fmt]}・A4横・実際のサイズで印刷）")
    for _ in range(pages):
        draw_page(c, fmt)
        c.showPage()
    c.save()


if __name__ == "__main__":
    build(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 1, sys.argv[3] if len(sys.argv) > 3 else "横")
