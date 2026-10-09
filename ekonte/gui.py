"""絵コンテ作成ソフトの画面。

① カット絵PDFを読み取る → ② 要確認を直す → ③ Excel（香盤表・カット表）を選ぶ → ④ 絵コンテPDFと香盤表Excelを書き出す
"""
import json, math, os, re, sys, traceback
from PySide6.QtCore import Qt, QThread, Signal, QSettings, QUrl, QTimer, QEvent, QRect, QRectF, QPoint, QPointF, QObject
from PySide6.QtNetwork import QNetworkAccessManager, QNetworkRequest, QNetworkReply, QSslSocket
from PySide6.QtGui import (QImage, QPixmap, QColor, QDesktopServices, QShortcut, QKeySequence, QPainter, QPainterPath,
                           QPen, QFontMetrics, QPolygonF)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, QMessageBox, QProgressBar, QSplitter,
    QCheckBox, QGroupBox, QLineEdit, QSpinBox, QAbstractItemView, QStyledItemDelegate, QSizePolicy,
    QDialog, QDialogButtonBox, QComboBox,
)
from . import __version__, REPO
from .omr import read_pdf, page_count
from .compose import compose, make_cut, check_cuts, CUT_RE
from .excel_import import load_book, describe_book
from .kouban import output_path as kouban_path_for
from .templates import build as build_template
from .sheet import build as build_sheet
from . import samples
from .layout import FORMATS

APP_NAME = "絵コンテ作成ソフト"
RELEASES_API = f"https://api.github.com/repos/{REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{REPO}/releases"  # 更新の知らせで開くのはこの下のページだけ
COL_PAGE, COL_NAME, COL_STATUS, COL_MSG = range(4)
STATUS_COLOR = {"OK": None, "修正済み": "#dff3e4", "確認済み": "#dff3e4", "要確認": "#fff4c2", "エラー": "#fbd5d5", "除外": "#e8e8e8"}


def to_pixmap(img):
    """PIL画像 → QPixmap"""
    img = img.convert("RGB")
    data = img.tobytes()
    q = QImage(data, img.width, img.height, img.width * 3, QImage.Format_RGB888)
    return QPixmap.fromImage(q.copy())


class ReadWorker(QThread):
    progress = Signal(int, int)       # 読んだページ数, 全ページ数
    page_done = Signal(object)        # PageResult
    failed = Signal(str)

    def __init__(self, path):
        super().__init__()
        self.path = path

    def run(self):
        try:
            total = page_count(self.path)
            for r in read_pdf(self.path):
                self.page_done.emit(r)
                self.progress.emit(r.page, total)
        except Exception as e:
            traceback.print_exc()
            self.failed.emit(str(e))


class NameDelegate(QStyledItemDelegate):
    """番号セルの編集。入力中から形式をチェックする。"""
    def createEditor(self, parent, option, index):
        e = QLineEdit(parent)
        e.setPlaceholderText("例: 2-3 / 2-3a（空欄で除外）")
        return e


class ImageView(QLabel):
    """枠に合わせて縮小表示する画像欄。"""
    def __init__(self, placeholder):
        super().__init__(placeholder)
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumSize(200, 120)
        # 画像の大きさで欄の大きさが変わらないようにする（変わると縮小→再配置が無限に続く）
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.setStyleSheet("QLabel { background: #fafafa; border: 1px solid #ccc; color: #888; }")
        self._pix = None

    def set_image(self, img):
        self._pix = to_pixmap(img) if img is not None else None
        if self._pix is None:
            self.clear()
        self._refresh()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._refresh()

    def _refresh(self):
        if self._pix:
            self.setPixmap(self._pix.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))


class FilePicker(QWidget):
    """Excelファイルの選択欄（選んだ時点で読み込んで中身を確認する）。"""
    def __init__(self, key, label, loader, describe, settings, template=None):
        super().__init__()
        self.template = template
        self.key, self.loader, self.describe, self.settings = key, loader, describe, settings
        self.path = None
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        title = QLabel(label)
        title.setMinimumWidth(150)
        self.info = QLabel("未選択")
        self.info.setWordWrap(True)
        pick = QPushButton("選択…")
        clear = QPushButton("解除")
        pick.clicked.connect(self.pick)
        clear.clicked.connect(lambda: self.set_path(None))
        lay.addWidget(title)
        lay.addWidget(self.info, 1)
        if template:
            tpl = QPushButton("テンプレート…")
            tpl.setToolTip("記入例つきの空のExcelを保存する")
            tpl.clicked.connect(self.save_template)
            lay.addWidget(tpl)
        lay.addWidget(pick)
        lay.addWidget(clear)
        last = settings.value(f"excel/{key}")
        if last and os.path.exists(last):
            self.set_path(last)

    def save_template(self):
        start = os.path.join(self.settings.value("dir/excel", ""), "香盤表・カット表.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, "テンプレートを保存", start, "Excel (*.xlsx)")
        if path:
            self.template(path)
            self.settings.setValue("dir/excel", os.path.dirname(path))
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def pick(self):
        start = os.path.dirname(self.path) if self.path else self.settings.value("dir/excel", "")
        path, _ = QFileDialog.getOpenFileName(self, "Excelを選択", start, "Excel (*.xlsx *.xlsm)")
        if path:
            self.settings.setValue("dir/excel", os.path.dirname(path))
            self.set_path(path)

    def set_path(self, path):
        if path is None:
            self.path = None
            self.info.setText("未選択")
            self.info.setStyleSheet("color: #888;")
            self.settings.remove(f"excel/{self.key}")
            return
        try:
            summary = self.describe(self.loader(path))
        except Exception as e:
            QMessageBox.warning(self, "Excelを読み込めません", f"{os.path.basename(path)}\n\n{e}")
            return
        self.path = path
        self.info.setText(f"{os.path.basename(path)}　（{summary}）")
        self.info.setStyleSheet("")
        self.settings.setValue(f"excel/{self.key}", path)


class SamplesDialog(QDialog):
    """同梱のサンプル・テンプレートを保存する画面。"""

    def __init__(self, parent, settings):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("サンプル・テンプレート")
        v = QVBoxLayout(self)
        v.addWidget(QLabel("1つずつ「保存…」するか、「すべてZipで保存…」でまとめて保存できます。\n"
                           "記入例は同じ作品の一式で、そのまま①〜④を試せます。"))
        for group in dict.fromkeys(g for g, *_ in samples.FILES):
            box = QGroupBox(group)
            grid = QGridLayout(box)
            for row, (_, name, _, desc) in enumerate(f for f in samples.FILES if f[0] == group):
                title = QLabel(name)
                title.setStyleSheet("font-weight: bold;")
                note = QLabel(desc)
                note.setStyleSheet("color: #666;")
                btn = QPushButton("保存…")
                btn.clicked.connect(lambda _=False, n=name: self.save_one(n))
                grid.addWidget(title, row, 0)
                grid.addWidget(note, row, 1)
                grid.addWidget(btn, row, 2)
            grid.setColumnStretch(1, 1)
            v.addWidget(box)
        buttons = QDialogButtonBox()
        b_all = buttons.addButton("すべてZipで保存…", QDialogButtonBox.ActionRole)
        b_all.clicked.connect(self.save_all)
        buttons.addButton("閉じる", QDialogButtonBox.RejectRole)
        buttons.rejected.connect(self.reject)
        v.addWidget(buttons)

    def _start_dir(self):
        return self.settings.value("dir/samples", os.path.expanduser("~/Downloads"))

    def save_one(self, name):
        ext = os.path.splitext(name)[1]
        kind = "PDF (*.pdf)" if ext == ".pdf" else "Excel (*.xlsx)"
        path, _ = QFileDialog.getSaveFileName(self, f"{name} を保存", os.path.join(self._start_dir(), name), kind)
        if not path:
            return
        if not self._save([(name, path)]):
            return
        self.settings.setValue("dir/samples", os.path.dirname(path))
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def save_all(self):
        start = os.path.join(self._start_dir(), samples.ZIP_NAME + ".zip")
        path, _ = QFileDialog.getSaveFileName(self, "サンプル一式をZipで保存", start, "Zip (*.zip)")
        if not path:
            return
        try:
            samples.save_zip(path)
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "保存できません", str(e))
            return
        self.settings.setValue("dir/samples", os.path.dirname(path))
        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))

    def _save(self, targets):
        try:
            for name, path in targets:
                samples.save(name, path)
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "保存できません", f"{name}\n\n{e}")
            return False
        return True


def version_tuple(v):
    """ "v0.10.2" → (0, 10, 2) """
    return tuple(int(x) for x in re.findall(r"\d+", v)[:3])


class UpdateChecker(QObject):
    """GitHub の Releases に新しい版があるか確認する。通信できなければ何もしない（オフラインでも使えるように）。

    通信は Qt（OS標準の暗号化通信）で行う。Python の urllib だと、アプリ化したときに証明書が見つからず失敗することがある。
    """
    found = Signal(str, str)  # 新しい版の番号, ダウンロードページのURL

    def __init__(self, parent):
        super().__init__(parent)
        self.nam = QNetworkAccessManager(self)

    def check(self):
        req = QNetworkRequest(QUrl(RELEASES_API))
        req.setRawHeader(b"Accept", b"application/vnd.github+json")
        req.setTransferTimeout(8000)
        reply = self.nam.get(req)
        reply.finished.connect(lambda: self._done(reply))

    def _done(self, reply):
        reply.deleteLater()
        if reply.error() != QNetworkReply.NoError:
            return
        try:
            data = json.loads(bytes(reply.readAll()).decode("utf-8"))
            tag, url = data["tag_name"], data["html_url"]
        except (ValueError, KeyError):
            return
        if not str(url).startswith(RELEASES_PAGE + "/"):
            url = RELEASES_PAGE  # 想定外のURLは開かない
        if version_tuple(tag) > version_tuple(__version__):
            self.found.emit(tag, url)


class GuideOverlay(QWidget):
    """初回起動のときだけ画面に重ねる案内。主なボタンを明るく残し、矢印付きの吹き出しで説明する。クリックで閉じる。"""
    GAP = 46  # ボタンと吹き出しの間

    def __init__(self, window, tips):
        super().__init__(window)
        self.tips = tips  # [(ウィジェット（複数ならまとめて囲む）, 文, 吹き出しの位置 "below"/"above"/"left", 横のずらし, 縦のずらし)]
        window.installEventFilter(self)
        self.setGeometry(window.rect())
        self.raise_()
        self.show()

    def eventFilter(self, obj, e):
        if e.type() == QEvent.Resize:
            self.setGeometry(obj.rect())
        return False

    def mousePressEvent(self, e):
        self.parent().removeEventFilter(self)
        self.deleteLater()

    def _target(self, widgets):
        rect = QRect()
        for w in widgets if isinstance(widgets, tuple) else (widgets,):
            rect = rect.united(QRect(w.mapTo(self.parent(), QPoint(0, 0)), w.size()))
        return QRectF(rect).adjusted(-5, -5, 5, 5)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        targets = [self._target(t[0]) for t in self.tips]
        shade = QPainterPath()
        shade.addRect(QRectF(self.rect()))
        for r in targets:
            hole = QPainterPath()
            hole.addRoundedRect(r, 6, 6)
            shade = shade.subtracted(hole)
        p.fillPath(shade, QColor(0, 0, 0, 165))
        yellow = QColor("#ffd54a")
        font = p.font()
        font.setPixelSize(14)
        p.setFont(font)
        fm = QFontMetrics(font)
        for (w, text, side, dx, dy), r in zip(self.tips, targets):
            p.setPen(QPen(yellow, 2))
            p.setBrush(Qt.NoBrush)
            p.drawRoundedRect(r, 6, 6)
            box = QRectF(fm.boundingRect(QRect(0, 0, 260, 400), Qt.TextWordWrap, text)).adjusted(-12, -9, 12, 9)
            cx = min(max(r.center().x() + dx, box.width() / 2 + 12), self.width() - box.width() / 2 - 12)
            if side == "below":
                box.moveCenter(QPointF(cx, r.bottom() + self.GAP + dy + box.height() / 2))
                start, end = QPointF(box.center().x(), box.top()), QPointF(r.center().x(), r.bottom() + 3)
            elif side == "left":
                box.moveCenter(QPointF(r.left() - self.GAP - box.width() / 2, r.center().y() + dy))
                start, end = QPointF(box.right(), box.center().y()), QPointF(r.left() - 3, r.center().y())
            else:
                box.moveCenter(QPointF(cx, r.top() - self.GAP - dy - box.height() / 2))
                start, end = QPointF(box.center().x(), box.bottom()), QPointF(r.center().x(), r.top() - 3)
            p.setPen(QPen(yellow, 2.5))
            p.drawLine(start, end)
            ang = math.atan2(end.y() - start.y(), end.x() - start.x())
            head = [end] + [end - QPointF(12 * math.cos(ang + a), 12 * math.sin(ang + a)) for a in (0.45, -0.45)]
            p.setBrush(yellow)
            p.drawPolygon(QPolygonF(head))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(box, 8, 8)
            p.setPen(QColor("#222"))
            p.drawText(box.adjusted(12, 9, -12, -9), Qt.TextWordWrap, text)
        font.setPixelSize(16)
        font.setBold(True)
        p.setFont(font)
        p.setPen(QColor("white"))
        p.drawText(self.rect().adjusted(0, 0, 0, -self.height() // 5), Qt.AlignCenter, "使い方　—　どこかをクリックすると閉じます")


def show_guide_once(window):
    """初回起動のときだけ、画面に案内を重ねる（あとからは「？ 使い方」ボタンで出せる）。"""
    if window.settings.value("guide/overlay_shown", False, type=bool):
        return
    window.settings.setValue("guide/overlay_shown", True)
    show_guide(window)


def show_guide(window):
    GuideOverlay(window, [
        (window.b_open, "① 描いたカット絵PDFを読み取ります → ② 黄色の「要確認」の行だけ確認します", "below", 120, 0),
        ((window.sheet_format, window.b_sheet), "最初に用紙を作ります。左で「横 / 縦 9:16」を選んでから。印刷は実際のサイズ（100%）で", "below", -60, 0),
        (window.b_samples, "初めてなら、ここの記入例で一度通してみるのがおすすめです", "below", 0, 120),
        (window.excel_box, "③ 香盤表とカット表をまとめたExcelを選びます（省略可）", "above", -150, 0),
        (window.b_export, "④ 絵コンテPDFと香盤表Excelを書き出します", "left", 0, 0),
    ])


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("Qbrick", "EkonteApp")
        self.setWindowTitle(f"{APP_NAME}  v{__version__}")
        self.resize(1200, 780)
        self.results = []     # PageResult（行と同じ順）
        self.edited = {}      # 行 → 手で直した番号
        self.confirmed = set()  # 番号を変えずに「この番号で確定」した行
        self.worker = None
        self._building = False

        root = QWidget()
        self.setCentralWidget(root)
        v = QVBoxLayout(root)

        # ① 読み取り
        top = QHBoxLayout()
        b_open = QPushButton("① カット絵PDFを読み取る…")
        b_open.setMinimumHeight(34)
        b_open.clicked.connect(self.open_pdf)
        self.b_open = b_open
        top.addWidget(b_open)
        self.pdf_label = QLabel("カット絵PDFを選んでください")
        top.addWidget(self.pdf_label, 1)
        top.addWidget(QLabel("用紙"))
        self.sheet_pages = QSpinBox()
        self.sheet_pages.setRange(1, 999)
        self.sheet_pages.setValue(20)
        self.sheet_pages.setSuffix(" ページ")
        top.addWidget(self.sheet_pages)
        self.sheet_format = QComboBox()
        for key, label in FORMATS.items():
            self.sheet_format.addItem(label, key)
        self.sheet_format.setToolTip("縦 9:16 は縦型動画用。読み取るときは用紙の種類を自動で判別します")
        i = self.sheet_format.findData(self.settings.value("sheet/format", "横"))
        self.sheet_format.setCurrentIndex(max(i, 0))
        top.addWidget(self.sheet_format)
        self.b_sheet = QPushButton("カット絵用紙を作成…")
        self.b_sheet.clicked.connect(self.make_sheet)
        top.addWidget(self.b_sheet)
        self.b_samples = QPushButton("サンプル・テンプレート…")
        self.b_samples.setToolTip("記入例（カット絵・香盤表・カット表・絵コンテ）と空のExcelを保存する")
        self.b_samples.clicked.connect(lambda: SamplesDialog(self, self.settings).exec())
        top.addWidget(self.b_samples)
        self.b_update = QPushButton()
        self.b_update.setStyleSheet("QPushButton { background: #e8590c; color: white; font-weight: bold; padding: 4px 12px; border-radius: 4px; }")
        self.b_update.setVisible(False)
        top.addWidget(self.b_update)
        b_guide = QPushButton("？ 使い方")
        b_guide.setToolTip("画面の使い方をもう一度表示する")
        b_guide.clicked.connect(lambda: show_guide(self))
        top.addWidget(b_guide)
        v.addLayout(top)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        v.addWidget(self.progress)

        # ② 確認・修正
        split = QSplitter(Qt.Horizontal)
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        self.summary = QLabel("")
        self.only_review = QCheckBox("要確認・エラーだけ表示")
        self.only_review.toggled.connect(self.apply_filter)
        row = QHBoxLayout()
        row.addWidget(QLabel("② 読み取り結果（番号はダブルクリックで修正）"))
        row.addStretch(1)
        row.addWidget(self.only_review)
        lv.addLayout(row)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ページ", "番号", "状態", "内容"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.table.setItemDelegateForColumn(COL_NAME, NameDelegate(self.table))
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(COL_PAGE, QHeaderView.ResizeToContents)
        h.setSectionResizeMode(COL_NAME, QHeaderView.Interactive)
        h.setSectionResizeMode(COL_STATUS, QHeaderView.ResizeToContents)
        h.setSectionResizeMode(COL_MSG, QHeaderView.Stretch)
        self.table.setColumnWidth(COL_NAME, 90)
        self.table.itemSelectionChanged.connect(self.show_selected)
        self.table.itemChanged.connect(self.on_item_changed)
        lv.addWidget(self.table, 1)
        row2 = QHBoxLayout()
        self.b_confirm = QPushButton("この番号で確定（Enter）")
        self.b_confirm.setToolTip("要確認の行で、読み取った番号が正しいときに押す")
        self.b_confirm.clicked.connect(self.confirm_selected)
        b_next = QPushButton("次の要確認へ（Tab）")
        b_next.clicked.connect(self.select_next_review)
        row2.addWidget(self.b_confirm)
        row2.addWidget(b_next)
        row2.addStretch(1)
        row2.addWidget(self.summary)
        lv.addLayout(row2)
        for key, slot in (("Return", self.confirm_selected), ("Enter", self.confirm_selected), ("Tab", self.select_next_review)):
            sc = QShortcut(QKeySequence(key), self.table)
            sc.setContext(Qt.WidgetShortcut)
            sc.activated.connect(slot)
        split.addWidget(left)

        right = QSplitter(Qt.Vertical)
        self.picture_view = ImageView("カット絵")
        self.mark_view = ImageView("マーク欄")
        right.addWidget(self.picture_view)
        right.addWidget(self.mark_view)
        right.setSizes([380, 420])
        split.addWidget(right)
        split.setSizes([620, 580])
        v.addWidget(split, 1)

        # ③ Excel（香盤表とカット表を1つのファイルに）
        box = QGroupBox("③ 絵コンテに入れるExcel（香盤表とカット表をまとめた1つのファイル。省略可。"
                        "Googleスプレッドシートは .xlsx でダウンロードして選ぶ）")
        g = QVBoxLayout(box)
        if not self.settings.value("excel/book") and self.settings.value("excel/cuts"):
            self.settings.setValue("excel/book", self.settings.value("excel/cuts"))  # 前の版で選んでいたカット表を引き継ぐ
        self.book = FilePicker("book", "香盤表・カット表", load_book, describe_book, self.settings, template=build_template)
        g.addWidget(self.book)
        v.addWidget(box)
        self.excel_box = box

        # ④ 書き出し
        bottom = QHBoxLayout()
        self.export_images = QCheckBox("カット絵の画像（2-3.png など）もフォルダに書き出す")
        bottom.addWidget(self.export_images)
        bottom.addStretch(1)
        self.b_export = QPushButton("④ 絵コンテPDFと香盤表を書き出す…")
        self.b_export.setMinimumHeight(38)
        self.b_export.setStyleSheet("QPushButton { font-weight: bold; padding: 0 18px; }")
        self.b_export.clicked.connect(self.export)
        bottom.addWidget(self.b_export)
        v.addLayout(bottom)

    def check_updates(self):
        self.updater = UpdateChecker(self)
        self.updater.found.connect(self.show_update)
        self.updater.check()

    def show_update(self, tag, url):
        self.b_update.setText(f"新しい版 {tag} があります")
        self.b_update.setToolTip(f"今の版は v{__version__}。押すとダウンロードページを開きます")
        self.b_update.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(url)))
        self.b_update.setVisible(True)

    # ---------- ① 読み取り ----------

    def make_sheet(self):
        n, fmt = self.sheet_pages.value(), self.sheet_format.currentData()
        name = "カット絵用紙.pdf" if fmt == "横" else f"カット絵用紙_{fmt}.pdf"
        path, _ = QFileDialog.getSaveFileName(self, "カット絵用紙を保存", os.path.join(self.settings.value("dir/pdf", ""), name), "PDF (*.pdf)")
        if not path:
            return
        build_sheet(path, n, fmt)
        self.settings.setValue("sheet/format", fmt)
        self._done_dialog("カット絵用紙を作成しました", f"{FORMATS[fmt]}・{n}ページ\n{path}\n\n印刷する場合は「実際のサイズ」で印刷してください。", path)

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "カット絵PDFを選択", self.settings.value("dir/pdf", ""), "PDF (*.pdf)")
        if not path:
            return
        self.settings.setValue("dir/pdf", os.path.dirname(path))
        self.results, self.edited, self.confirmed = [], {}, set()
        self._building = True
        self.table.setRowCount(0)
        self._building = False
        self.picture_view.set_image(None)
        self.mark_view.set_image(None)
        self.pdf_label.setText(os.path.basename(path))
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.b_open.setEnabled(False)
        self.worker = ReadWorker(path)
        self.worker.page_done.connect(self.add_result)
        self.worker.progress.connect(lambda i, n: (self.progress.setMaximum(n), self.progress.setValue(i)))
        self.worker.failed.connect(lambda msg: QMessageBox.critical(self, "読み取りエラー", msg))
        self.worker.finished.connect(self.read_finished)
        self.worker.start()

    def add_result(self, r):
        self.results.append(r)
        self._building = True
        row = self.table.rowCount()
        self.table.insertRow(row)
        item = QTableWidgetItem(str(r.page))
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        item.setTextAlignment(Qt.AlignCenter)
        self.table.setItem(row, COL_PAGE, item)
        self.table.setItem(row, COL_NAME, QTableWidgetItem(r.name or ""))
        for col in (COL_STATUS, COL_MSG):
            it = QTableWidgetItem("")
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, col, it)
        self._building = False
        self.refresh_status()
        if row == 0:
            self.table.selectRow(0)

    def sheet_format_of_results(self):
        """読み取ったカット絵の用紙の種類（最初に読めたページに合わせる。混在は読み取り時に要確認にしている）。"""
        return next((r.format for r in self.results if r.status != "エラー"), "横")

    def read_finished(self):
        if self.results:
            self.pdf_label.setText(f"{self.pdf_label.text()}　（{FORMATS[self.sheet_format_of_results()]}）")
        self.progress.setVisible(False)
        self.b_open.setEnabled(True)
        self.refresh_status()
        self.select_next_review(from_start=True)

    # ---------- ② 確認・修正 ----------

    def current_name(self, row):
        return self.table.item(row, COL_NAME).text().strip()

    def row_status(self, row, counts=None):
        """(状態, 内容) を返す。手で直した行は読み取り時の警告を出さない。"""
        r = self.results[row]
        name = self.current_name(row)
        if counts is None:
            counts = self._name_counts()
        if r.status == "エラー" and not name:
            return "エラー", "；".join(r.warnings)
        if not name:
            return "除外", "番号が空欄のため絵コンテに入れません" + ("（" + "；".join(r.warnings) + "）" if r.warnings else "")
        if not CUT_RE.match(name):
            return "要確認", "番号の形式が正しくありません（例: 2-3, 2-3a）"
        if counts.get(name, 0) > 1:
            pages = [str(self.results[i].page) for i in range(len(self.results)) if self.current_name(i) == name]
            return "要確認", f"番号 {name} が重複しています（p{', p'.join(pages)}）"
        if row in self.edited:
            return "修正済み", f"手で修正（読み取り: {r.name or 'なし'}）"
        if row in self.confirmed:
            return "確認済み", "読み取った番号を目で確認して確定"
        warnings = [w for w in r.warnings if "重複" not in w]
        note = "" if r.orientation in ("", "正位置") else f"［{r.orientation}を補正］"
        if warnings:
            return "要確認", "；".join(warnings) + note
        return "OK", note

    def _name_counts(self):
        counts = {}
        for i in range(self.table.rowCount()):
            n = self.current_name(i)
            if n:
                counts[n] = counts.get(n, 0) + 1
        return counts

    def refresh_status(self):
        self._building = True
        counts = self._name_counts()
        tally = {}
        for row in range(self.table.rowCount()):
            status, msg = self.row_status(row, counts)
            tally[status] = tally.get(status, 0) + 1
            self.table.item(row, COL_STATUS).setText(status)
            self.table.item(row, COL_MSG).setText(msg)
            color = STATUS_COLOR.get(status)
            for col in range(4):
                self.table.item(row, col).setBackground(QColor(color) if color else QColor(0, 0, 0, 0))
        self._building = False
        parts = [f"{k} {v}" for k, v in tally.items()]
        self.summary.setText(f"全 {self.table.rowCount()} ページ：" + "　".join(parts) if parts else "")
        self.apply_filter()

    def on_item_changed(self, item):
        if self._building or item.column() != COL_NAME:
            return
        row = item.row()
        name = item.text().strip()
        if name != (self.results[row].name or ""):
            self.edited[row] = name
        else:
            self.edited.pop(row, None)
        self.confirmed.discard(row)
        self.refresh_status()

    def selected_row(self):
        rows = self.table.selectionModel().selectedRows()
        return rows[0].row() if rows else None

    def confirm_selected(self):
        row = self.selected_row()
        if row is None or not self.current_name(row):
            return
        status, msg = self.row_status(row)
        # 形式の誤り・重複は確定できない（番号を直す必要がある）
        if not CUT_RE.match(self.current_name(row)) or "重複" in msg:
            QMessageBox.information(self, APP_NAME, "番号の形式が正しくないか、他のページと重複しています。\n番号を直してください。")
            return
        if status == "要確認":
            self.confirmed.add(row)
        self.refresh_status()
        self.select_next_review()

    def select_next_review(self, from_start=False):
        n = self.table.rowCount()
        cur = -1 if from_start or self.selected_row() is None else self.selected_row()
        for k in range(1, n + 1):
            i = (cur + k) % n
            if self.table.item(i, COL_STATUS).text() in ("要確認", "エラー"):
                self.table.selectRow(i)
                self.table.scrollToItem(self.table.item(i, COL_NAME))
                return

    def apply_filter(self):
        only = self.only_review.isChecked()
        for row in range(self.table.rowCount()):
            status = self.table.item(row, COL_STATUS).text()
            self.table.setRowHidden(row, only and status not in ("要確認", "エラー"))

    def show_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        r = self.results[rows[0].row()]
        self.picture_view.set_image(r.picture)
        self.mark_view.set_image(r.mark_area)
        if r.mark_area is None:
            self.mark_view.setText("（位置合わせマークが見つからないため表示できません）")

    # ---------- ④ 書き出し ----------

    def export(self):
        if not self.results:
            QMessageBox.information(self, APP_NAME, "先に ① でカット絵PDFを読み取ってください。")
            return
        counts = self._name_counts()
        statuses = [self.row_status(i, counts)[0] for i in range(len(self.results))]
        pending = [str(self.results[i].page) for i, s in enumerate(statuses) if s == "要確認"]
        if pending:
            QMessageBox.warning(self, "要確認のページがあります",
                                f"p{', p'.join(pending)} が要確認のままです。\n番号を修正するか、空欄にして除外してから書き出してください。")
            return
        cuts = [make_cut(self.current_name(i), self.results[i].picture)
                for i, s in enumerate(statuses) if s in ("OK", "修正済み", "確認済み")]
        if not cuts and not self.book.path:
            QMessageBox.warning(self, APP_NAME, "絵コンテに入れるカットがありません。")
            return
        if self.book.path and not self._confirm_cut_check([f"{c['s']}-{c['c']}" for c in cuts]):
            return
        start = os.path.join(self.settings.value("dir/out", self.settings.value("dir/pdf", "")), "絵コンテ.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "絵コンテPDFを保存（香盤表のExcelも隣に保存します）", start, "PDF (*.pdf)")
        if not path:
            return
        kouban = kouban_path_for(path, self.book.path)
        if os.path.exists(kouban) and QMessageBox.question(
                self, APP_NAME, f"{os.path.basename(kouban)} はすでにあります。上書きしますか？") != QMessageBox.Yes:
            return
        self.settings.setValue("dir/out", os.path.dirname(path))
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            result = compose(cuts, path, self.book.path, self.sheet_format_of_results(), kouban_out=kouban)
            if self.export_images.isChecked():
                folder = os.path.splitext(path)[0] + "_カット絵"
                os.makedirs(folder, exist_ok=True)
                for c in cuts:
                    if c["image"] is not None:
                        c["image"].save(os.path.join(folder, f"{c['s']}-{c['c']}.png"))
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "書き出しエラー", str(e))
            return
        finally:
            QApplication.restoreOverrideCursor()
        excluded = statuses.count("除外") + statuses.count("エラー")
        msg = f"{result['cuts']}カット / {result['pages']}ページ\n{path}\n{kouban}"
        if excluded:
            msg += f"\n\n※ 除外・エラーの {excluded} ページは入れていません"
        if result["warnings"]:
            msg += "\n\n⚠ 文字を最小にしても欄に収まらず、はみ出た分を切っています：\n" + "\n".join(result["warnings"])
        self._done_dialog("絵コンテと香盤表を書き出しました", msg, path, kouban)

    def _confirm_cut_check(self, names):
        """カット表とカット絵が食い違っていれば、内容を見せて続けるか確認する。"""
        try:
            missing, extra = check_cuts(names, self.book.path)
        except Exception as e:
            QMessageBox.critical(self, "カット表を読み込めません", str(e))
            return False
        if not missing and not extra:
            return True
        lines = []
        if missing:
            lines.append(f"カット絵がないカット（PICTURE を空けて載せます）：\n  {', '.join(missing)}")
        if extra:
            lines.append(f"カット表にないカット（ACTION/SE などは空欄になります）：\n  {', '.join(extra)}")
        ans = QMessageBox.question(self, "カット表とカット絵が一致しません", "\n\n".join(lines) + "\n\nこのまま書き出しますか？")
        return ans == QMessageBox.Yes

    def _done_dialog(self, title, text, path, kouban=None):
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(text)
        open_btn = box.addButton("開く", QMessageBox.AcceptRole)
        kouban_btn = box.addButton("香盤表を開く", QMessageBox.ActionRole) if kouban else None
        box.addButton("閉じる", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is open_btn:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        elif kouban_btn and box.clickedButton() is kouban_btn:
            QDesktopServices.openUrl(QUrl.fromLocalFile(kouban))


def selftest(pdf, out):
    """アプリ化したあとの動作確認用：画面を出さずに 読み取り→絵コンテ書き出し を行う。
    Ekonte --selftest カット絵.pdf 出力.pdf"""
    results = list(read_pdf(pdf))
    cuts = [make_cut(r.name, r.picture) for r in results if r.name]
    fmt = next((r.format for r in results if r.status != "エラー"), "横")
    r = compose(cuts, out, fmt=fmt, kouban_out=kouban_path_for(out))  # 香盤表のExcelも書き出せるか確認する
    w = MainWindow()  # 画面が組み立てられるかも確認する
    SamplesDialog(w, w.settings)
    samples.save_zip(os.path.splitext(out)[0] + "_サンプル.zip")  # 同梱のサンプルを取り出せるかも確認する
    tls = QSslSocket.supportsSsl()  # 更新確認に使う暗号化通信の部品がアプリに入っているか
    from openpyxl.xml import DEFUSEDXML  # 悪意のあるExcel（XML爆弾）への対策が効いているか
    print(f"selftest: {FORMATS[fmt]} {len(results)}ページ読み取り / {r['cuts']}カット / {r['pages']}ページ書き出し / "
          f"香盤表OK / サンプル{len(samples.FILES)}件保存 / 画面OK / 暗号化通信{'OK' if tls else 'なし'} / "
          f"Excel対策{'OK' if DEFUSEDXML else 'なし'}")
    for r in results:
        print(f"  p{r.page}: {r.name} {r.status} {'；'.join(r.warnings)}")
    return tls and DEFUSEDXML


def main():
    # 出力先の文字コードが日本語に対応していない環境（Windows の自動ビルドなど）でも、print で落ちないようにする
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    if "--selftest" in sys.argv:
        i = sys.argv.index("--selftest")
        app = QApplication(sys.argv[:1])
        try:
            ok = selftest(sys.argv[i + 1], sys.argv[i + 2])
        except Exception:
            # 画面のないアプリでは、捕まえないと Windows でエラーのダイアログが出て止まったままになる
            traceback.print_exc()
            ok = False
        sys.exit(0 if ok else 1)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    w = MainWindow()
    w.show()
    QTimer.singleShot(0, lambda: show_guide_once(w))
    QTimer.singleShot(1500, w.check_updates)  # 起動を遅らせないよう、画面が出てから確認する
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
