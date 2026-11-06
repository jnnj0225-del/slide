"""永山 洵 先生のハウススタイル(python-pptx 用ヘルパー)。

references/house-style.md の仕様をコードにしたもの。症例提示・学会口演とも
これを import して使う:

    import sys; sys.path.insert(0, "<skill>/scripts")
    from nu_style import *
    prs = new_presentation()            # 16:9 の白紙デッキ
    s = blank(prs)
    header(s, "腹部造影CT", sub="症例提示（患者背景）")
    ...
    finish(prs, "out.pptx")             # ページ番号・ロゴを全スライドに付けて保存
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")
LOGO_COLOR = os.path.join(ASSETS, "nu_med_logo_color.png")   # 854x150
LOGO_WHITE = os.path.join(ASSETS, "nu_med_logo_white.png")   # 856x84

# ---- 配色(水色寄りの青を基調、オレンジはアクセントのみ)
BLUE = RGBColor(0x1F, 0x80, 0xC8)        # 基調色(表紙背景・帯・表ヘッダー)
BLUE_DARK = RGBColor(0x13, 0x5C, 0x96)   # タイトル文字
LIGHT = RGBColor(0xEA, 0xF4, 0xFC)       # パネル背景
ORANGE = RGBColor(0xE8, 0x73, 0x1A)      # アクセント(区切り線・強調ボタン)
TEXT = RGBColor(0x26, 0x26, 0x26)
GRAY = RGBColor(0x6E, 0x6E, 0x6E)
RED = RGBColor(0xC0, 0x00, 0x00)         # 異常値・重要所見
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SUBTLE = RGBColor(0xD6, 0xE9, 0xF8)      # 濃い背景上の補助文字

FONT = "メイリオ"
TITLE_PT = 40      # スライドタイトル 40pt 以上
BODY_PT = 22       # 本文 20pt 以上(入る限り大きく)
SMALL_PT = 14      # 出典・注記のみ
CENTER = PP_ALIGN.CENTER
TODO = "[要確認]"

SW, SH = Inches(13.333), Inches(7.5)
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def new_presentation(src=None):
    prs = Presentation(src) if src else Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    return prs


def clear_placeholders(slide):
    for ph in list(slide.placeholders):
        if not hasattr(ph, "image"):
            ph._element.getparent().remove(ph._element)


def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    clear_placeholders(s)
    return s


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.25)
    s.shadow.inherit = False
    return s


def _font(run, size, bold, color):
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("ea", "cs"):
        rPr.append(rPr.makeelement(_A + tag, {"typeface": FONT}))


def text(slide, x, y, w, h, paras, size=BODY_PT, color=TEXT, bold=False,
         align=CENTER, anchor=MSO_ANCHOR.TOP, space_after=6):
    """paras: 文字列 / [(text, {size,bold,color})] のラン配列 / {"runs":…, "bullet":True} の dict。
    既定は中央揃え。"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_after = Pt(space_after)
        if isinstance(p, dict):
            runs, bullet = p["runs"], p.get("bullet")
        else:
            runs, bullet = p, None
        if isinstance(runs, str):
            runs = [(runs, {})]
        if bullet:
            runs = [("・", {})] + list(runs)
        for t, o in runs:
            r = para.add_run()
            r.text = t
            _font(r, o.get("size", size), o.get("bold", bold), o.get("color", color))
    return tb


def header(slide, title, sub=None, size=TITLE_PT):
    """中央揃えの 40pt タイトル + オレンジ/青の二重線。sub は右上の小さな区分ラベル。"""
    text(slide, Inches(0.4), Inches(0.2), Inches(12.53), Inches(0.95), [title],
         size=size, color=BLUE_DARK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, Inches(0.45), Inches(1.2), Inches(12.43), Inches(0.07), ORANGE)
    rect(slide, Inches(0.45), Inches(1.29), Inches(12.43), Inches(0.035), BLUE)
    if sub:
        text(slide, Inches(7.4), Inches(0.0), Inches(5.6), Inches(0.3), [sub],
             size=SMALL_PT, color=GRAY, align=PP_ALIGN.RIGHT)


def band_header(slide, label, title, sub=None):
    """ディスカッションポイント用の青帯ヘッダー。"""
    rect(slide, 0, 0, SW, Inches(1.35), BLUE)
    text(slide, Inches(0.4), Inches(0.2), Inches(12.53), Inches(0.95),
         [[(label + "　", {"color": SUBTLE}), (title, {})]],
         size=TITLE_PT, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if sub:
        text(slide, Inches(7.4), Inches(1.38), Inches(5.6), Inches(0.35), [sub],
             size=SMALL_PT, color=GRAY, align=PP_ALIGN.RIGHT)


def panel(slide, x, y, w, h, paras, size=BODY_PT, title=None, align=CENTER, anchor=MSO_ANCHOR.MIDDLE):
    rect(slide, x, y, w, h, LIGHT)
    rect(slide, x, y, w, Inches(0.07), BLUE)
    body = ([[(title, {"bold": True, "color": BLUE_DARK})]] if title else []) + list(paras)
    return text(slide, x + Inches(0.12), y + Inches(0.15), w - Inches(0.24), h - Inches(0.25),
                body, size=size, align=align, anchor=anchor)


def caption(slide, x, y, w, t, size=20, color=BLUE_DARK):
    text(slide, x, y, w, Inches(0.45), [t], size=size, color=color, bold=True)


def table(slide, x, y, w, rows, colw, size=BODY_PT, rowh=0.55, hl=()):
    """先頭行がヘッダー(青)。hl={(行,列)} は赤太字。全セル中央揃え。"""
    t = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w, Inches(rowh * len(rows))).table
    for j, cw in enumerate(colw):
        t.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        t.rows[i].height = Inches(rowh)
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            c.fill.fore_color.rgb = BLUE if i == 0 else (LIGHT if i % 2 == 0 else WHITE)
            for k, line in enumerate(v.split("\n")):
                p = c.text_frame.paragraphs[0] if k == 0 else c.text_frame.add_paragraph()
                p.alignment = CENTER
                r = p.add_run()
                r.text = line
                _font(r, size, i == 0 or (i, j) in hl,
                      WHITE if i == 0 else (RED if (i, j) in hl else TEXT))
    return t


def placeholder_box(slide, label):
    rect(slide, Inches(0.9), Inches(1.7), Inches(11.5), Inches(4.9), WHITE, line=GRAY)
    text(slide, Inches(0.9), Inches(3.6), Inches(11.5), Inches(1.0), [label], size=26, color=GRAY)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def cover(prs, title_lines, affiliation, name, corner=None):
    """表紙: 青背景、タイトル・所属・氏名は中央揃え、corner(会名・日付)は右上、白ロゴは右下。"""
    s = blank(prs)
    rect(s, 0, 0, SW, SH, BLUE)
    if corner:
        text(s, Inches(7.0), Inches(0.3), Inches(6.0), Inches(0.9), corner, size=18,
             color=SUBTLE, align=PP_ALIGN.RIGHT, space_after=0)
    text(s, Inches(0.6), Inches(1.4), Inches(12.13), Inches(2.75), title_lines,
         size=44, color=WHITE, bold=True, anchor=MSO_ANCHOR.BOTTOM)
    rect(s, Inches(2.5), Inches(4.4), Inches(8.33), Inches(0.07), ORANGE)
    text(s, Inches(0.6), Inches(4.7), Inches(12.13), Inches(1.4),
         [affiliation, [(name, {"size": 32, "bold": True})]], size=24, color=WHITE)
    w = Inches(4.6)
    s.shapes.add_picture(LOGO_WHITE, SW - w - Inches(0.35), SH - Inches(0.85), w, int(w * 84 / 856))
    return s


def finish(prs, out, skip_first=True):
    """全スライド(表紙以外)に、右下ロゴ・左下ページ番号を付けて保存。"""
    for n, s in enumerate(prs.slides, 1):
        if skip_first and n == 1:
            continue
        h = Inches(0.42)
        w = int(h * 854 / 150)
        s.shapes.add_picture(LOGO_COLOR, SW - w - Inches(0.3), SH - h - Inches(0.12), w, h)
        text(s, Inches(0.3), SH - Inches(0.5), Inches(1.0), Inches(0.35), [str(n)],
             size=SMALL_PT, color=GRAY, align=PP_ALIGN.LEFT)
    prs.save(out)
