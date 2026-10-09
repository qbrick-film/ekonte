"""PDFに埋め込む日本語フォント（BIZ UDゴシック、SIL Open Font License）。

フォントをPDFに埋め込むので、Windows・Macのどちらで開いても同じ見た目になる。
"""
import os, sys
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

JP = "BIZUDGothic"
JP_BOLD = "BIZUDGothic-Bold"


def resource_path(*parts):
    """開発時はリポジトリの resources/、アプリ化後は同梱先の resources/ を指す。"""
    base = getattr(sys, "_MEIPASS", os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base, "resources", *parts)


def register():
    if JP in pdfmetrics.getRegisteredFontNames():
        return
    pdfmetrics.registerFont(TTFont(JP, resource_path("fonts", "BIZUDGothic-Regular.ttf")))
    pdfmetrics.registerFont(TTFont(JP_BOLD, resource_path("fonts", "BIZUDGothic-Bold.ttf")))
    registerFontFamily(JP, normal=JP, bold=JP_BOLD, italic=JP, boldItalic=JP_BOLD)


register()
