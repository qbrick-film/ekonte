"""動作確認用：実際に描いて塗った想定のカット絵PDFと、それを印刷→スキャンした想定のPDFを作る。

python -m ekonte.sample 出力フォルダ
"""
import math, os, random, sys
from reportlab.pdfgen import canvas
import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFilter
from .layout import *
from .sheet import draw_page
from .fonts import JP

FW, FH = FRAMES["横"]["w"], FRAMES["横"]["h"]   # 横の描画枠
VW, VH = FRAMES["縦"]["w"], FRAMES["縦"]["h"]   # 縦（9:16）の描画枠


# ---------- 手描き風の線 ----------

class Pen:
    def __init__(self, c, seed, fmt="横"):
        self.c = c
        self.rnd = random.Random(seed)
        self.fx, self.fy = FRAMES[fmt]["x"], FRAMES[fmt]["y"]

    def _wobble(self, pts, amp=1.0):
        """折れ線を細かく分割し、なめらかな揺れを加える。"""
        out = []
        ph1, ph2 = self.rnd.uniform(0, 6.3), self.rnd.uniform(0, 6.3)
        t = 0
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            d = math.hypot(x2 - x1, y2 - y1) or 1
            nx, ny = -(y2 - y1) / d, (x2 - x1) / d
            steps = max(2, int(d / 5))
            for i in range(steps):
                f = i / steps
                t += d / steps
                off = amp * (math.sin(t / 23 + ph1) * 0.7 + math.sin(t / 7 + ph2) * 0.3)
                out.append((x1 + (x2 - x1) * f + nx * off, y1 + (y2 - y1) * f + ny * off))
        out.append(pts[-1])
        return out

    def stroke(self, pts, width=1.4, gray=0.12, amp=1.0, frame=True):
        """frame=True のとき描画枠からの相対座標で描く。"""
        if frame:
            pts = [(self.fx + x, self.fy + y) for x, y in pts]
        # 始点・終点を少しはみ出させる（手描きの勢い）
        (x1, y1), (x2, y2) = pts[0], pts[1]
        d = math.hypot(x2 - x1, y2 - y1) or 1
        k = self.rnd.uniform(0, 3) / d
        pts = [(x1 - (x2 - x1) * k, y1 - (y2 - y1) * k)] + pts[1:]
        w = self._wobble(pts, amp)
        c = self.c
        c.setStrokeGray(gray)
        c.setLineWidth(width * self.rnd.uniform(0.85, 1.15))
        c.setLineCap(1)
        c.setLineJoin(1)
        p = c.beginPath()
        p.moveTo(*w[0])
        for q in w[1:]:
            p.lineTo(*q)
        c.drawPath(p, stroke=1, fill=0)
        c.setStrokeGray(0)

    def line(self, x1, y1, x2, y2, **kw):
        self.stroke([(x1, y1), (x2, y2)], **kw)

    def poly(self, pts, close=False, **kw):
        self.stroke(pts + ([pts[0]] if close else []), **kw)

    def ellipse(self, cx, cy, rx, ry, start=0, sweep=360, **kw):
        a0 = math.radians(start + self.rnd.uniform(-10, 10))
        n = max(12, int((rx + ry) * sweep / 90))
        s = math.radians(sweep + self.rnd.uniform(0, 15))
        self.stroke([(cx + rx * math.cos(a0 + s * i / n), cy + ry * math.sin(a0 + s * i / n)) for i in range(n + 1)], **kw)

    def hatch(self, x, y, w, h, gap=7, angle=1.0, **kw):
        kw.setdefault("width", 0.7)
        kw.setdefault("gray", 0.35)
        for i in range(int((w + h) / gap)):
            x1 = x + i * gap
            pts = [(x1, y), (x1 - h * angle, y + h)]
            # 矩形内に収める
            (ax, ay), (bx, by) = pts
            if ax > x + w:
                ay += (ax - x - w) / angle; ax = x + w
            if bx < x:
                by -= (x - bx) / angle; bx = x
            if ay < by and ax >= x:
                self.line(ax, ay, bx, by, amp=0.4, **kw)

    def arrow(self, x1, y1, x2, y2, **kw):
        self.line(x1, y1, x2, y2, **kw)
        a = math.atan2(y2 - y1, x2 - x1)
        for s in (2.6, -2.6):
            self.line(x2, y2, x2 - 12 * math.cos(a + s / 6), y2 - 12 * math.sin(a + s / 6), **kw)

    def text(self, x, y, s, size=13):
        c = self.c
        c.saveState()
        c.translate(self.fx + x, self.fy + y)
        c.rotate(self.rnd.uniform(-4, 4))
        c.setFillGray(0.15)
        c.setFont(JP, size)
        c.drawString(0, 0, s)
        c.restoreState()

    def person(self, x, y, h, facing=0, arm=None):
        """足元(x,y)、身長h の人物。arm=(dx,dy) で片腕を伸ばす。"""
        hr = h * 0.09
        self.ellipse(x, y + h - hr, hr * 0.9, hr)
        neck = y + h - hr * 2
        hip = y + h * 0.45
        self.line(x, neck, x + facing * 2, hip, width=1.6)
        self.line(x + facing * 2, hip, x - h * 0.12, y, width=1.5)
        self.line(x + facing * 2, hip, x + h * 0.12, y, width=1.5)
        sh = neck - h * 0.06
        self.line(x, sh, x - h * 0.18, sh - h * 0.25)
        if arm:
            self.line(x, sh, x + arm[0], sh + arm[1])
        else:
            self.line(x, sh, x + h * 0.18, sh - h * 0.25)


# ---------- 各カットの絵 ----------

def cut_1_1(p):  # 公園・ロング
    p.line(0, 150, FW, 158, width=1.2)
    p.line(200, 0, 255, 150, width=1); p.line(380, 0, 300, 150, width=1)
    for tx, s in ((70, 1.0), (470, 1.2), (520, 0.8)):
        p.line(tx, 150, tx, 150 + 70 * s, width=2)
        p.ellipse(tx, 150 + 95 * s, 38 * s, 30 * s)
    p.ellipse(460, 320, 18, 18)
    p.person(265, 60, 70); p.person(300, 64, 66, arm=(-14, -10))
    p.text(12, 340, "FIX")


def cut_1_2(p):  # 2人のミディアム
    for x, f in ((190, 1), (380, -1)):
        p.ellipse(x, 240, 42, 50)
        p.line(x - 20, 192, x - 30, 160); p.line(x + 20, 192, x + 30, 160)
        p.poly([(x - 110, 0), (x - 95, 120), (x - 30, 160), (x + 30, 160), (x + 95, 120), (x + 110, 0)])
        p.line(x + 15 * f, 245, x + 25 * f, 245, width=1)
    p.hatch(0, 280, FW, 87, gap=10, gray=0.5)
    p.arrow(330, 330, 410, 330); p.text(332, 340, "振り向く", 12)


def cut_2_1(p):  # 倉庫・拳銃を構える
    vx, vy = 300, 200
    for x, y in ((0, 0), (FW, 0), (0, FH), (FW, FH)):
        p.line(x, y, vx + (x - vx) * 0.35, vy + (y - vy) * 0.35, width=1)
    p.poly([(vx - 98, vy - 70), (vx + 91, vy - 70), (vx + 91, vy + 58), (vx - 98, vy + 58)], close=True, width=1)
    for i in range(3):
        p.poly([(40, 60 + i * 60), (130, 80 + i * 55), (130, 120 + i * 55), (40, 110 + i * 60)], close=True, width=1)
    p.person(330, 30, 180, arm=(70, 8))
    p.poly([(400, 160), (425, 162), (425, 152), (410, 150), (408, 140)], width=2)
    p.hatch(0, 270, FW, 97, gap=8)


def cut_2_2(p):  # 拳銃アップ
    p.poly([(120, 200), (380, 210), (385, 250), (140, 245)], close=True, width=2.4)
    p.poly([(330, 205), (360, 120), (400, 125), (375, 208)], close=True, width=2.4)
    p.ellipse(345, 165, 22, 14, start=180, sweep=180)
    p.poly([(300, 120), (330, 60), (420, 40), (470, 90), (430, 140)], width=2)
    for i in range(4):
        p.line(390 + i * 9, 135 - i * 4, 405 + i * 11, 70 - i * 3, width=1)
    p.line(100, 228, 40, 230, width=1); p.line(100, 220, 30, 214, width=1)
    p.text(16, 330, "UP")


def cut_2_3(p):  # 車が入ってくる
    p.line(0, 80, FW, 85, width=1)
    p.poly([(150, 100), (160, 160), (230, 165), (270, 215), (390, 215), (430, 165), (480, 160), (485, 100)], close=True, width=2)
    p.poly([(280, 205), (335, 205), (335, 168), (250, 168)], close=True, width=1.2)
    p.poly([(345, 205), (385, 205), (415, 168), (345, 168)], close=True, width=1.2)
    p.ellipse(215, 100, 26, 26, width=2); p.ellipse(420, 100, 26, 26, width=2)
    for i in range(3):
        p.line(70, 120 + i * 25, 130, 122 + i * 25, width=1, gray=0.4)
    p.person(520, 82, 120)
    p.arrow(60, 300, 200, 300); p.text(70, 310, "PAN →", 12)


def cut_2_3a(p):  # インサート・ヘッドライト
    p.ellipse(280, 180, 120, 95, width=2.4)
    p.ellipse(280, 180, 70, 55, width=1.6)
    p.ellipse(280, 180, 25, 20, width=1.2)
    for a in range(0, 360, 30):
        r = math.radians(a)
        p.line(280 + 140 * math.cos(r), 180 + 112 * math.sin(r), 280 + 200 * math.cos(r), 180 + 160 * math.sin(r), width=0.9, gray=0.4)
    p.text(16, 330, "INSERT")


def cut_3_1(p):  # 車内・後ろから
    p.poly([(40, 330), (520, 330), (480, 170), (80, 170)], close=True, width=2)
    p.line(180, 170, 260, 330, width=0.8); p.line(380, 170, 300, 330, width=0.8)
    p.line(280, 180, 280, 210, width=0.8); p.line(280, 240, 280, 270, width=0.8)
    for x in (180, 390):
        p.ellipse(x, 120, 55, 62, width=1.8)
        p.hatch(x - 45, 120, 90, 50, gap=6, gray=0.25)
        p.line(x - 80, 0, x - 55, 70); p.line(x + 80, 0, x + 55, 70)
    p.ellipse(170, 165, 60, 18, start=0, sweep=180, width=1.6)


def cut_3_2(p):  # 運転席の顔アップ
    p.line(60, 0, 120, FH, width=1.2)
    p.ellipse(310, 190, 95, 125, width=2)
    for x in (270, 350):
        p.ellipse(x, 210, 16, 8, width=1.4); p.ellipse(x, 210, 4, 4, width=1)
        p.line(x - 20, 238, x + 18, 232, width=1.6)
    p.line(310, 200, 300, 160); p.line(300, 160, 315, 158)
    p.line(285, 125, 340, 128, width=1.4)
    p.hatch(215, 270, 190, 50, gap=5, gray=0.2)
    p.line(220, 70, 180, 0, width=1.5); p.line(400, 70, 440, 0, width=1.5)


# ---------- 縦（9:16）の絵。横と同じ話・同じ番号なので、記入例の香盤表・カット表をそのまま使える ----------

def tate_1_1(p):  # 公園・ロング
    p.line(0, 190, VW, 196, width=1.2)
    p.line(95, 0, 118, 190, width=1); p.line(170, 0, 140, 190, width=1)
    for tx, s in ((40, 1.1), (215, 0.9)):
        p.line(tx, 190, tx, 190 + 80 * s, width=2)
        p.ellipse(tx, 190 + 110 * s, 34 * s, 36 * s)
    p.ellipse(190, 400, 16, 16)
    p.person(118, 90, 64); p.person(146, 94, 60, arm=(-12, -9))
    p.text(10, 420, "FIX")


def tate_1_2(p):  # 2人のミディアム（縦に重ねる：手前に太郎、奥に花子）
    p.ellipse(150, 300, 30, 36)
    p.poly([(85, 210), (100, 250), (130, 265), (170, 265), (200, 250), (215, 210)])
    p.line(135, 266, 132, 248); p.line(165, 266, 168, 248)
    p.ellipse(105, 150, 44, 52, width=2)
    p.poly([(0, 0), (15, 70), (70, 98), (140, 98), (195, 70), (215, 0)], width=2)
    p.hatch(0, 370, VW, 78, gap=10, gray=0.5)
    p.arrow(170, 345, 225, 330); p.text(150, 352, "振り向く", 11)


def tate_2_1(p):  # 倉庫・拳銃を構える
    vx, vy = 126, 260
    for x, y in ((0, 0), (VW, 0), (0, VH), (VW, VH)):
        p.line(x, y, vx + (x - vx) * 0.4, vy + (y - vy) * 0.4, width=1)
    p.poly([(vx - 50, vy - 104), (vx + 50, vy - 104), (vx + 50, vy + 75), (vx - 50, vy + 75)], close=True, width=1)
    p.person(110, 20, 230, arm=(90, 10))
    p.poly([(200, 215), (222, 217), (222, 207), (208, 205), (206, 195)], width=2)
    p.hatch(0, 380, VW, 68, gap=8)


def tate_2_2(p):  # 拳銃アップ（銃口を上に向けて縦に）
    p.poly([(105, 120), (115, 400), (150, 400), (145, 140)], close=True, width=2.4)
    p.poly([(110, 150), (40, 170), (45, 210), (112, 190)], close=True, width=2.4)
    p.ellipse(80, 175, 14, 20, start=90, sweep=180)
    p.poly([(140, 140), (200, 110), (215, 30), (150, 10), (110, 60)], width=2)
    for i in range(4):
        p.line(150 + i * 10, 100 - i * 6, 165 + i * 10, 40 - i * 4, width=1)
    p.text(12, 420, "UP")


def tate_2_3(p):  # 車が入ってくる（正面から）
    p.line(0, 120, VW, 124, width=1)
    p.poly([(30, 130), (30, 230), (60, 240), (85, 300), (170, 300), (195, 240), (222, 230), (222, 130)], close=True, width=2)
    p.poly([(92, 290), (162, 290), (182, 245), (72, 245)], close=True, width=1.2)
    for x in (60, 192):
        p.ellipse(x, 200, 16, 12, width=1.6)
    p.poly([(85, 150), (167, 150), (167, 175), (85, 175)], close=True, width=1.2)
    for x in (50, 202):
        p.poly([(x - 14, 130), (x - 14, 100), (x + 14, 100), (x + 14, 130)], width=2)
    p.arrow(40, 395, 210, 395); p.text(70, 405, "PAN →", 12)


def tate_2_3a(p):  # インサート・ヘッドライト
    p.ellipse(126, 230, 95, 80, width=2.4)
    p.ellipse(126, 230, 55, 46, width=1.6)
    p.ellipse(126, 230, 20, 17, width=1.2)
    for a in range(0, 360, 30):
        r = math.radians(a)
        p.line(126 + 108 * math.cos(r), 230 + 96 * math.sin(r), 126 + 122 * math.cos(r), 230 + 150 * math.sin(r), width=0.9, gray=0.4)
    p.text(12, 420, "INSERT")


def tate_3_1(p):  # 車内・後ろから
    p.poly([(15, 400), (237, 400), (215, 250), (37, 250)], close=True, width=2)
    p.line(126, 260, 126, 290, width=0.8); p.line(126, 320, 126, 350, width=0.8)
    for x, y in ((70, 160), (182, 160)):
        p.ellipse(x, y, 42, 50, width=1.8)
        p.hatch(x - 34, y, 68, 40, gap=6, gray=0.25)
        p.line(x - 60, 0, x - 40, y - 45); p.line(x + 60, 0, x + 40, y - 45)


def tate_3_2(p):  # 運転席の顔アップ
    p.line(20, 0, 50, VH, width=1.2)
    p.ellipse(126, 250, 85, 115, width=2)
    for x in (95, 157):
        p.ellipse(x, 270, 14, 7, width=1.4); p.ellipse(x, 270, 4, 4, width=1)
        p.line(x - 18, 295, x + 16, 290, width=1.6)
    p.line(126, 262, 118, 225); p.line(118, 225, 131, 223)
    p.line(103, 190, 150, 193, width=1.4)
    p.hatch(45, 325, 165, 50, gap=5, gray=0.2)
    p.line(60, 150, 25, 0, width=1.5); p.line(192, 150, 227, 0, width=1.5)


FAINT_3_2 = {"faint": "c1"}  # C一を薄く塗った例（要確認になるはず）
CUTS = {  # 用紙の種類 → [(番号, 絵, 塗り方)]
    "横": [("1-1", cut_1_1, {}), ("1-2", cut_1_2, {}),
          ("2-1", cut_2_1, {}), ("2-2", cut_2_2, {}), ("2-3", cut_2_3, {}), ("2-3a", cut_2_3a, {}),
          ("3-1", cut_3_1, {}), ("3-2", cut_3_2, FAINT_3_2)],
    "縦": [("1-1", tate_1_1, {}), ("1-2", tate_1_2, {}),
          ("2-1", tate_2_1, {}), ("2-2", tate_2_2, {}), ("2-3", tate_2_3, {}), ("2-3a", tate_2_3a, {}),
          ("3-1", tate_3_1, {}), ("3-2", tate_3_2, FAINT_3_2)],
}


# ---------- マークを手で塗る ----------

def marks_for(number):
    s, c = number.split("-")
    sub = c[-1] if c[-1].isalpha() else None
    c = c.rstrip("abcdefghij")
    m = {}
    if len(s) >= 3: m["s100"] = s[-3]
    if len(s) >= 2: m["s10"] = s[-2]
    m["s1"] = s[-1]
    if len(c) >= 2: m["c10"] = c[-2]
    m["c1"] = c[-1]
    if sub: m["sub"] = sub
    return m


def fill_bubble(c, rnd, bx, by, faint=False):
    """塗りつぶし。中心が少しずれ、往復の線で塗る。"""
    cx, cy = bx + rnd.uniform(-1, 1), by + rnd.uniform(-1, 1)
    r = BUBBLE_R * rnd.uniform(0.85, 1.0)
    gap, width = (3.6, 0.9) if faint else (1.5, 2.4)
    c.setStrokeGray(0.45 if faint else 0.08)
    c.setLineWidth(width)
    c.setLineCap(1)
    p = c.beginPath()
    y, left = cy - r * 0.9, True
    p.moveTo(cx, y)
    while y <= cy + r * 0.9:
        half = math.sqrt(max(r * r - (y - cy) ** 2, 0)) * rnd.uniform(0.8, 1.0)
        p.lineTo(cx + (-half if left else half), y)
        left = not left
        y += gap
    c.drawPath(p, stroke=1, fill=0)
    c.setStrokeGray(0)


def write_digits(c, rnd, m):
    """確認用マスに手書き風の数字を書く。"""
    for key, _, x, _ in COLUMNS:
        if key in m:
            c.saveState()
            c.translate(x - 5 + rnd.uniform(-1.5, 1.5), 503 + rnd.uniform(-1, 1))
            c.rotate(rnd.uniform(-6, 6))
            c.setFillGray(0.1)
            c.setFont(JP, 17)
            c.drawString(0, 0, m[key])
            c.restoreState()


def build_drawn(path, fmt="横"):
    c = canvas.Canvas(path, pagesize=(PAGE_W, PAGE_H))
    for i, (number, draw, opt) in enumerate(CUTS[fmt]):
        draw_page(c, fmt)
        rnd = random.Random(100 + i)
        draw(Pen(c, seed=i, fmt=fmt))
        m = marks_for(number)
        write_digits(c, rnd, m)
        for key, _, x, vals in COLUMNS:
            if key in m:
                fill_bubble(c, rnd, *bubble_center(x, vals.index(m[key])), faint=opt.get("faint") == key)
        c.showPage()
    c.save()


# ---------- 印刷→スキャンを再現 ----------

def build_scanned(src, path, dpi=150):
    pdf = pdfium.PdfDocument(src)
    pages = []
    rnd = random.Random(7)
    for i, page in enumerate(pdf):
        img = page.render(scale=dpi / 72, fill_color=(255, 255, 255, 255)).to_pil().convert("L")
        w, h = img.size
        # 紙の色・わずかな縮小・傾き・ずれ
        angle = rnd.uniform(-2.5, 2.5) + (180 if i == 3 else 0)  # 4ページ目は上下逆に置いた想定
        s = rnd.uniform(0.95, 0.99)
        small = img.resize((int(w * s), int(h * s)), Image.LANCZOS)
        paper = Image.new("L", (w, h), 255)
        paper.paste(small, ((w - small.width) // 2 + rnd.randint(-15, 15), (h - small.height) // 2 + rnd.randint(-10, 10)))
        img = paper.rotate(angle, resample=Image.BICUBIC, fillcolor=255)
        # 汚れ・ノイズ・にじみ
        d = ImageDraw.Draw(img)
        for _ in range(250):
            x, y = rnd.randint(0, w), rnd.randint(0, h)
            d.point((x, y), fill=rnd.randint(60, 180))
        img = img.point(lambda v: int(v * 0.93 + 8)).filter(ImageFilter.GaussianBlur(0.7))
        pages.append(img.convert("RGB"))
    pages[0].save(path, save_all=True, append_images=pages[1:], resolution=dpi)


if __name__ == "__main__":
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    for fmt, suffix in (("横", ""), ("縦", "_縦")):
        drawn = os.path.join(out, f"カット絵_記入例{suffix}.pdf")
        build_drawn(drawn, fmt)
        build_scanned(drawn, os.path.join(out, f"カット絵_記入例{suffix}_スキャン.pdf"))
        print("作成:", drawn)
