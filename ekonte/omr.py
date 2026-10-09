"""カット絵PDFのマークを読み取り、描画枠を「S-C.png」として書き出す。

python -m ekonte.omr 入力.pdf 出力フォルダ

デジタルで直接描いたPDFも、印刷→スキャンしたPDFも同じ手順で読む。
  1. 四隅の位置合わせマークを探す
  2. 向き判定ドットで上下（90°回転も）を決める
  3. 射影変換で用紙座標→画像座標の対応を求め、傾き・ずれ・縮小を補正
  4. 各マークの中心部の黒の割合で「塗った・薄い・塗っていない」を判定
  5. 描画枠を座標で切り出す
判定に迷ったページは 出力フォルダ/要確認/ にマーク欄の画像を保存する。
"""
import json, math, os, sys
from dataclasses import dataclass, field
from collections import deque
import numpy as np
import pypdfium2 as pdfium
from PIL import Image
from .layout import *

KEY_LABEL = {"s100": "S百", "s10": "S十", "s1": "S一", "c10": "C十", "c1": "C一", "sub": "枝"}

RENDER_SCALE = 3      # PDFを画像化する倍率（216dpi相当）
MAX_RENDER_PX = 4000  # 画像の長辺の上限。極端に大きいページのPDFでメモリを使い果たさないようにする（A4は約2500px）
OUT_SCALE = 3         # 書き出す画像の解像度（px/pt）
BLACK = 140           # 位置合わせマーク探索で「黒」とみなす明るさ
FILLED = 0.40         # 白紙からの濃さの増分がこれ以上 → 塗った
FAINT = 0.08          # これ以上 FILLED 未満 → 薄い（要確認）
DOT_MIN = 0.45        # 向き判定ドットの濃さ
SAMPLE_R = 0.6        # 枠線を拾わないよう、半径の60%以内だけを見る


# ---------- 1. 位置合わせマーク ----------

def _blobs(mask):
    """2値画像中の黒い塊ごとに (重心x, 重心y, 画素数, 幅, 高さ) を返す。"""
    h, w = mask.shape
    seen = np.zeros_like(mask, dtype=bool)
    ys, xs = np.nonzero(mask)
    for y0, x0 in zip(ys, xs):
        if seen[y0, x0]:
            continue
        q = deque([(y0, x0)])
        seen[y0, x0] = True
        pts = []
        while q:
            y, x = q.popleft()
            pts.append((y, x))
            for ny, nx in ((y + 1, x), (y - 1, x), (y, x + 1), (y, x - 1)):
                if 0 <= ny < h and 0 <= nx < w and mask[ny, nx] and not seen[ny, nx]:
                    seen[ny, nx] = True
                    q.append((ny, nx))
        p = np.array(pts)
        yield p[:, 1].mean(), p[:, 0].mean(), len(p), np.ptp(p[:, 1]) + 1, np.ptp(p[:, 0]) + 1


def find_corners(gray):
    """位置合わせマークをページ全体から探し、画像上の左上・右上・右下・左下の順で返す。

    「中まで塗られた正方形」の塊を候補にし、いちばん外側の4つを四隅とする。
    画像の隅だけを探すと、縮小印刷（用紙に合わせる）で紙の中央に小さく載った場合に見つからないため。
    マーク欄の円や向き判定ドットが候補に混じっても、四隅より内側にあるので選ばれない。
    """
    h, w = gray.shape
    step = max(1, round(max(w, h) / 900))  # 長辺 ≈900px 相当に間引いて探す
    mask = gray[::step, ::step] < BLACK
    longest = max(mask.shape)
    cands = []
    for cx, cy, n, bw, bh in _blobs(mask):
        side = (bw + bh) / 2
        if not (0.005 * longest < side < 0.04 * longest):   # 用紙の 0.5〜4% の大きさ
            continue
        if not (0.75 < bw / bh < 1.33) or n < bw * bh * 0.75:  # 正方形で中まで塗られている（ぼやけを許容）
            continue
        cands.append((cx * step, cy * step, n))
    if len(cands) < 4:
        raise ValueError("位置合わせマーク（四隅の黒い四角）が見つかりません。用紙を設計図から生成しているか確認してください")
    c = np.array(cands, float)
    idx = [np.argmin(c[:, 0] + c[:, 1]), np.argmax(c[:, 0] - c[:, 1]),
           np.argmax(c[:, 0] + c[:, 1]), np.argmin(c[:, 0] - c[:, 1])]
    sizes = c[idx, 2]
    if sizes.max() > sizes.min() * 2:
        raise ValueError("位置合わせマークの大きさが揃っていません。用紙を設計図から生成しているか確認してください")
    corners = [c[i, :2] for i in idx]
    # 4点が用紙の縦横比（約1.45）の四角形になっているか確認する
    d = [np.hypot(*(corners[i] - corners[(i + 1) % 4])) for i in range(4)]
    ratio = max(d[0] + d[2], d[1] + d[3]) / max(1, min(d[0] + d[2], d[1] + d[3]))
    expect = (PAGE_W - 42) / (PAGE_H - 42)
    if not (expect * 0.85 < ratio < expect * 1.15):
        raise ValueError("位置合わせマークの配置が用紙と合いません。用紙を設計図から生成しているか確認してください")
    return [tuple(p) for p in corners]


# ---------- 2-3. 向き判定と射影変換 ----------

def to_top_left(x, y):
    """layout.py の座標（左下原点）を上端基準に直す。"""
    return x, PAGE_H - y


def homography(src, dst):
    """4点の対応から射影変換行列 H（src→dst）を求める。"""
    A, b = [], []
    for (x, y), (u, v) in zip(src, dst):
        A.append([x, y, 1, 0, 0, 0, -u * x, -u * y]); b.append(u)
        A.append([0, 0, 0, x, y, 1, -v * x, -v * y]); b.append(v)
    h = np.linalg.solve(np.array(A, float), np.array(b, float))
    return np.append(h, 1).reshape(3, 3)


def apply(H, pts):
    pts = np.asarray(pts, float)
    p = np.c_[pts, np.ones(len(pts))] @ H.T
    return p[:, :2] / p[:, 2:3]


def to_ink(gray):
    """画素の濃さを 0（紙の白）〜1（黒）に正規化する。紙の色はページごとに推定。"""
    paper = np.percentile(gray, 90)
    return np.clip((paper - gray.astype(float)) / (paper * 0.75), 0, 1)


def ink_ratio(ink, H, cx, cy, r):
    """用紙上の円（中心cx,cy 半径r、上端基準）の内側の平均の濃さ（0〜1）。

    白黒の2値ではなく濃さで測るので、スキャンで薄くなった塗りも拾える。"""
    g = np.linspace(-r, r, 15)
    xx, yy = np.meshgrid(g, g)
    inside = xx ** 2 + yy ** 2 <= r * r
    pts = apply(H, np.c_[xx[inside] + cx, yy[inside] + cy])
    h, w = ink.shape
    xs = np.clip(pts[:, 0].round().astype(int), 0, w - 1)
    ys = np.clip(pts[:, 1].round().astype(int), 0, h - 1)
    return float(ink[ys, xs].mean())


def solve_orientation(ink, img_corners):
    """4通りの回転を試し、向き判定ドットが黒くなる対応を採用する。"""
    sheet = [to_top_left(*CORNER_CENTERS[k]) for k in ("tl", "tr", "br", "bl")]
    dot = to_top_left(ORIENT_DOT["x"], ORIENT_DOT["y"])
    best = None
    for k in range(4):
        H = homography(sheet, img_corners[k:] + img_corners[:k])
        score = ink_ratio(ink, H, *dot, ORIENT_DOT["r"] * SAMPLE_R)
        if best is None or score > best[0]:
            best = (score, H, k)
    score, H, k = best
    if score < DOT_MIN:
        raise ValueError("向き判定ドットが見つかりません")
    return H, ("正位置", "90°回転", "上下逆", "270°回転")[k]


# ---------- 4. マーク判定 ----------

def read_marks(ink, H):
    detail, warnings, values, flagged = {}, [], {}, set()
    for key, label, x, vals in COLUMNS:
        raw = [ink_ratio(ink, H, *to_top_left(*bubble_center(x, row)), BUBBLE_R * SAMPLE_R) for row in range(len(vals))]
        # 円の中に印刷した数字の分だけ白紙でも少し濃いので、同じ桁の中央値（＝白紙の濃さ）を差し引く
        base = float(np.median(raw))
        ratios = [r - base for r in raw]
        filled = [vals[i] for i, r in enumerate(ratios) if r >= FILLED]
        faint = [vals[i] for i, r in enumerate(ratios) if FAINT <= r < FILLED]
        detail[key] = {v: round(r, 2) for v, r in zip(vals, ratios) if r >= 0.05}
        name = KEY_LABEL[key]
        if filled or faint:
            flagged.add(key)
        if len(filled) > 1:
            warnings.append(f"{name}: 複数マーク {'・'.join(filled)}")
            values[key] = None
        elif filled:
            values[key] = filled[0]
            if faint:
                warnings.append(f"{name}: {filled[0]} の他に薄いマーク {'・'.join(faint)}")
        elif faint:
            values[key] = faint[0] if len(faint) == 1 else None
            warnings.append(f"{name}: 薄いマーク {'・'.join(faint)}")
        else:
            values[key] = None
    for key in ("s1", "c1"):
        if key not in flagged:
            warnings.append(f"{KEY_LABEL[key]}: 未記入")
    if None in (values["s1"], values["c1"]):
        return None, warnings, detail
    s = int((values["s100"] or "0") + (values["s10"] or "0") + values["s1"])
    c = int((values["c10"] or "0") + values["c1"])
    return f"{s}-{c}{values['sub'] or ''}", warnings, detail


# ---------- 5. 座標による切り出し ----------

def crop(img, H, x, y, w, h):
    """用紙上の矩形（左下原点）を、傾きを補正して切り出す。"""
    left, top = x, PAGE_H - y - h
    # 出力画素 → 用紙座標 → 画像座標
    T = np.array([[1 / OUT_SCALE, 0, left], [0, 1 / OUT_SCALE, top], [0, 0, 1]])
    M = H @ T
    M = M / M[2, 2]
    size = (int(w * OUT_SCALE), int(h * OUT_SCALE))
    return img.transform(size, Image.PERSPECTIVE, M.flatten()[:8], Image.BICUBIC, fillcolor="white")


@dataclass
class PageResult:
    page: int                       # 1始まり
    name: str = None                # 読み取った番号（"2-3a"）。読めなければ None
    status: str = "OK"              # OK / 要確認 / エラー
    warnings: list = field(default_factory=list)
    orientation: str = ""
    format: str = "横"              # 用紙の種類（layout.FORMATS のキー）。縦型の印の有無で決める
    marks: dict = field(default_factory=dict)
    picture: Image.Image = None     # 描画枠の中（傾き補正済み）
    mark_area: Image.Image = None   # マーク欄（人が確認する用）


def read_page(page, index):
    """pdfium のページ1枚を読む。"""
    scale = min(RENDER_SCALE, MAX_RENDER_PX / max(*page.get_size(), 1))
    img = page.render(scale=scale, draw_annots=True, fill_color=(255, 255, 255, 255)).to_pil().convert("L")
    gray = np.asarray(img)
    ink = to_ink(gray)
    r = PageResult(page=index + 1)
    try:
        H, r.orientation = solve_orientation(ink, find_corners(gray))
    except ValueError as e:
        r.status, r.warnings = "エラー", [str(e)]
        r.picture = img
        return r
    r.name, r.warnings, r.marks = read_marks(ink, H)
    score = ink_ratio(ink, H, *to_top_left(FORMAT_DOT["x"], FORMAT_DOT["y"]), FORMAT_DOT["r"] * SAMPLE_R)
    r.format = "縦" if score >= DOT_MIN / 2 else "横"
    r.status = "要確認" if r.warnings else "OK"
    m, f, a = 3, FRAMES[r.format], MARK_AREA  # m: 描画枠の線を含めないよう内側を切り抜く
    r.picture = crop(img, H, f["x"] + m, f["y"] + m, f["w"] - m * 2, f["h"] - m * 2)
    r.mark_area = crop(img, H, a["x"], a["y"], a["w"], a["h"])
    return r


def page_count(pdf_path):
    return len(pdfium.PdfDocument(pdf_path))


def read_pdf(pdf_path):
    """PDFの全ページを順に読み、PageResult を1ページずつ返す。番号の重複も警告する。"""
    pdf = pdfium.PdfDocument(pdf_path)
    used, first = {}, None
    for i in range(len(pdf)):
        r = read_page(pdf[i], i)
        if r.status != "エラー":
            first = first or r
            if r.format != first.format:  # 絵コンテは1種類のレイアウトで作るので、混ざったら知らせる
                r.warnings.append(f"用紙が{FORMATS[r.format]}（p{first.page} は{FORMATS[first.format]}）")
                r.status = "要確認"
        if r.name in used:
            r.warnings.append(f"{r.name} は p{used[r.name]} と重複")
            r.status = "要確認"
        elif r.name:
            used[r.name] = r.page
        yield r


def main(pdf_path, outdir):
    os.makedirs(outdir, exist_ok=True)
    review_dir = os.path.join(outdir, "要確認")
    report, saved = [], set()
    for r in read_pdf(pdf_path):
        entry = dict(page=r.page, name=r.name, status=r.status, orientation=r.orientation, warnings=r.warnings, marks=r.marks)
        if r.status != "エラー":
            if r.name and r.name not in saved:
                saved.add(r.name)
                r.picture.save(os.path.join(outdir, f"{r.name}.png"))
                entry["file"] = f"{r.name}.png"
            if r.warnings:
                os.makedirs(review_dir, exist_ok=True)
                r.mark_area.save(os.path.join(review_dir, f"p{r.page}_マーク欄.png"))
                if "file" not in entry:
                    r.picture.save(os.path.join(review_dir, f"p{r.page}_カット絵.png"))
        report.append(entry)
        if r.status == "エラー":
            print(f"p{r.page}: ✗ {r.warnings[0]}")
        else:
            note = f"  [{r.orientation}]" if r.orientation != "正位置" else ""
            print(f"p{r.page}: {r.name or '???'}{note}" + (f"  ⚠ {'; '.join(r.warnings)}" if r.warnings else ""))
    with open(os.path.join(outdir, "読み取り結果.json"), "w", encoding="utf-8") as fp:
        json.dump(report, fp, ensure_ascii=False, indent=2)
    ok = sum(r["status"] == "OK" for r in report)
    print(f"--- {len(report)}ページ中 OK {ok} / 要確認・エラー {len(report) - ok}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
