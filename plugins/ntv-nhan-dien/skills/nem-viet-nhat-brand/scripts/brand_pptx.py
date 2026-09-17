# -*- coding: utf-8 -*-
"""Nệm Việt Nhật — bộ công cụ dựng slide PowerPoint theo nhận diện thương hiệu (Brand Book 2026).

Cách dùng:
    import sys; sys.path.insert(0, r"<skill>/scripts")
    from brand_pptx import *
    prs = new_deck()
    s = add_slide(prs)
    add_bg(s, "bg_title.png")            # nền gradient (asset trong skill)
    cover(s, "TÀI LIỆU HỌP", "TIÊU ĐỀ LỚN", "Phụ đề")
    prs.save("output.pptx")

Tất cả toạ độ tính bằng inch trên khung 16:9 (13.333 x 7.5).
Asset (logo, nền) tự resolve trong thư mục ../assets của skill.
Font mặc định: SVN-Gilroy (Bold heading / Medium subhead / Regular body).
Nấc đậm (weight) trong txt()/bullets(): "thin" "light" "regular" "medium" "semibold" "bold" "heavy" "black"
(+ italic=True cho bản nghiêng). Không truyền weight thì như cũ: bold=True → Bold, không thì Regular.
Mọi run đều khai font ở cả 3 khe latin/ea/cs (xem _set_all_font_slots) để dấu tiếng Việt không rớt nấc.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---- Asset path: <skill>/assets ----
_HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(_HERE, "..", "assets"))

# ---- Bảng màu thương hiệu (Brand Book Nệm Việt Nhật 2026) ----
GREEN    = RGBColor(0x40,0xA6,0x46)   # Xanh chủ đạo  #40A646
GREEN_DK = RGBColor(0x2A,0x78,0x3B)   # Xanh rêu      #2A783B
DEEP     = RGBColor(0x05,0x34,0x21)   # Xanh đậm nền tối / đầu gradient  #053421
LEMON    = RGBColor(0xFF,0xFD,0xA4)   # Vàng chanh nhạt — tiêu đề lớn trên nền tối, đầu gradient phụ #FFFDA4
BLACK    = RGBColor(0x23,0x1F,0x20)
WHITE    = RGBColor(0xFF,0xFF,0xFF)
CREAM    = RGBColor(0xFC,0xFC,0xF5)   # nền kem card (brand book)
GRAY     = RGBColor(0xBB,0xBD,0xBF)
# tên chung để code không phụ thuộc màu: PRIMARY = màu chủ đạo, PRIMARY_DK = đậm, DARK = nền tối
PRIMARY, PRIMARY_DK, DARK, ACCENT = GREEN, GREEN_DK, DEEP, LEMON
# Alias tương thích code cũ viết cho skill Việt Nhật (PINK → GREEN); KHÔNG dùng cho code mới
PINK, PINK_DK, DARKPINK = GREEN, GREEN_DK, DEEP
# Màu logo trong file PNG gốc (bảng màu 11/2025) — KHÁC brand book 2026 một chút, chỉ dùng khi phải khớp logo cũ
LOGO_GREEN_2025, LOGO_GREEN_DK_2025 = RGBColor(0x28,0xA5,0x4A), RGBColor(0x27,0x77,0x14)
# tông hỗ trợ layout
INK      = RGBColor(0x2E,0x33,0x2F)   # body text
MUTE     = RGBColor(0x6B,0x72,0x6C)   # caption
CARD     = RGBColor(0xEA,0xF6,0xEC)   # card xanh nhạt
CARD2    = RGBColor(0xF5,0xF6,0xF5)   # card xám nhạt
LINEC    = RGBColor(0xD9,0xE8,0xDC)   # đường kẻ nhạt
PALEGREEN= RGBColor(0xCD,0xEF,0xD3)   # chữ phụ trên nền tối
ACCENTGR = RGBColor(0x8F,0xE0,0x8A)   # nhãn nhấn trên nền tối
PALEPINK, ACCENTPK = PALEGREEN, ACCENTGR   # alias code cũ

# Font chính thức: SVN-Gilroy (đã cài per-user, 16 weight).
# Dự phòng nếu máy chưa cài: "Be Vietnam Pro" (Google Fonts, SIL OFL) hoặc "Montserrat".
FONT = "SVN-Gilroy"

# ---- Nấc đậm (weight) của SVN-Gilroy ----
# Bộ font đóng gói 16 file. Trong bảng tên (name table) của font:
#   Regular / Bold / Italic / Bold Italic  → family "SVN-Gilroy"          + cờ bold/italic
#   Thin, Light, Medium, SemiBold, Heavy, Black → family riêng "SVN-Gilroy <Nấc>" (+ file Italic riêng)
# Cột thứ 3 = tên PostScript của bản NGHIÊNG. Test 17/09/2026 trên macOS (Quick Look/CoreText):
#   gọi "SVN-Gilroy Heavy" + cờ italic → macOS KHÔNG tìm được bản Heavy Italic, đổ về nghiêng thường;
#   gọi thẳng tên PostScript "SVN-GilroyHeavyItalic" → đúng. Windows PowerPoint chưa test cách này;
#   nếu bên Windows thấy chữ nghiêng nấc lạ rớt font, đổi ITALIC_NAME_MODE = "family".
WEIGHTS = {
    #  nấc        (family thẳng,          cờ bold, tên PostScript bản nghiêng)
    "thin":     ("SVN-Gilroy Thin",     False, "SVN-GilroyThinItalic"),
    "light":    ("SVN-Gilroy Light",    False, "SVN-GilroyLightItalic"),
    "regular":  ("SVN-Gilroy",          False, None),   # None = family + cờ italic (chuẩn, chạy mọi nơi)
    "medium":   ("SVN-Gilroy Medium",   False, "SVN-GilroyMediumItalic"),
    "semibold": ("SVN-Gilroy SemiBold", False, "SVN-GilroySemiBoldItalic"),
    "bold":     ("SVN-Gilroy",          True,  None),
    "heavy":    ("SVN-Gilroy Heavy",    False, "SVN-GilroyHeavyItalic"),
    "black":    ("SVN-Gilroy Black",    False, "SVN-GilroyBlackItalic"),
}
ITALIC_NAME_MODE = "postscript"   # "postscript" (đúng trên macOS) | "family" (kiểu OOXML chuẩn)

def font_for(weight=None, bold=False, italic=False, font=FONT):
    """Trả (tên typeface, cờ bold, cờ italic) cho một nấc đậm. weight=None → hành vi cũ (font + bold + italic)."""
    if weight is None:
        return font, bool(bold), bool(italic)
    w = str(weight).lower().replace(" ", "").replace("-", "")
    if w not in WEIGHTS:
        raise ValueError(f"weight '{weight}' không có; chọn một trong: {', '.join(WEIGHTS)}")
    name, b, ps_italic = WEIGHTS[w]
    if font != FONT:              # người dùng đổi font khác → không áp bảng SVN-Gilroy
        return font, (w in ("bold", "semibold", "heavy", "black")), bool(italic)
    if italic and ps_italic and ITALIC_NAME_MODE == "postscript":
        return ps_italic, b, True   # giữ cờ italic: font đã nghiêng sẵn nên vô hại, còn làm fallback đúng hơn
    return name, b, bool(italic)

def _style_run(r, size, color, weight=None, bold=False, italic=False, font=FONT):
    name, b, i = font_for(weight, bold, italic, font)
    r.font.size = Pt(size); r.font.bold = b; r.font.italic = i
    r.font.name = name; r.font.color.rgb = color
    _set_all_font_slots(r, name)

def _set_all_font_slots(r, name):
    """Khai cùng một font ở cả 3 khe latin / ea / cs của run.
    Lý do (test 17/09/2026, macOS Quick Look/CoreText): file .pptx chỉ khai khe latin thì nhóm chữ Việt
    "Latin mở rộng bổ sung" (ủ ố ề ắ ặ ỡ ữ ệ …) bị render bằng khe ea/cs đang trống → rớt về nấc Regular
    của họ font (Heavy/Black thành mảnh, Thin thành đậm). PowerPoint thật cũng ghi đủ 3 khe khi người dùng
    chọn font, nên cách này chuẩn và vô hại."""
    from pptx.oxml.ns import qn
    from lxml import etree
    rPr = r._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    prev = latin
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.Element(qn(tag))
            if prev is not None: prev.addnext(el)
            else: rPr.append(el)
        el.set("typeface", name); prev = el

set_font_slots = _set_all_font_slots   # tên công khai (import * kéo theo được)

SW, SH = 13.333, 7.5
DOC_TITLE = "Nệm Việt Nhật"   # đổi qua set_doc_title() nếu cần footer khác

def set_doc_title(t):
    global DOC_TITLE; DOC_TITLE = t

# ---- Khởi tạo deck / slide ----
def new_deck():
    prs = Presentation(); prs.slide_width = Inches(SW); prs.slide_height = Inches(SH)
    return prs
def add_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank

# ---- Nền ----
def add_bg(s, img):
    """Nền ảnh full slide. img: 'bg_title.png' | 'bg_section.png' | 'bg_gradient.png' hoặc path tuyệt đối."""
    p = img if os.path.isabs(img) else os.path.join(ASSETS, img)
    s.shapes.add_picture(p, 0, 0, Inches(SW), Inches(SH))
def solid_bg(s, color=WHITE):
    f = s.background.fill; f.solid(); f.fore_color.rgb = color

# ---- Text ----
def txt(s, text, x, y, w, h, size, color, bold=False, italic=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, spacing=None, line=1.0, font=FONT, weight=None):
    """Khung chữ. weight: "thin"|"light"|"regular"|"medium"|"semibold"|"bold"|"heavy"|"black" (xem WEIGHTS)."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = line
        r = p.add_run(); r.text = ln
        _style_run(r, size, color, weight, bold, italic, font)
        if spacing is not None:
            r._r.get_or_add_rPr().set('spc', str(int(spacing*100)))
    return tb

def bullets(s, items, x, y, w, h, size, color, bold_color=GREEN, gap=7, line=1.06, font=FONT,
            weight=None, lead_weight="bold", italic=False):
    """items: list[str] hoặc list[(lead, rest)] (phần lead in đậm theo lead_weight, mặc định Bold).
    weight: nấc đậm cho chữ thường (mặc định Regular); dấu • luôn theo lead_weight."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line; p.space_after = Pt(gap)
        rb = p.add_run(); rb.text = "•  "
        _style_run(rb, size, bold_color, lead_weight, True, False, font)
        if isinstance(it, tuple):
            lead, rest = it
            r1 = p.add_run(); r1.text = lead
            _style_run(r1, size, BLACK, lead_weight, True, italic, font)
            r2 = p.add_run(); r2.text = rest
            _style_run(r2, size, color, weight, False, italic, font)
        else:
            r = p.add_run(); r.text = it
            _style_run(r, size, color, weight, False, italic, font)
    return tb

# ---- Shapes ----
def rrect(s, x, y, w, h, color, rounded=True):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color; sp.line.fill.background(); sp.shadow.inherit = False
    return sp
def oval(s, x, y, w, h, color):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color; sp.line.fill.background(); sp.shadow.inherit = False
    return sp
def chevron(s, x, y, w, h, color):
    sp = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color; sp.line.fill.background(); sp.shadow.inherit = False
    return sp
def line(s, x, y, w, h, color):
    """Đường kẻ mảnh (dùng cho connector sơ đồ). Cho w hoặc h nhỏ (vd 0.024)."""
    rrect(s, x, y, w, h, color, rounded=False)

# ---- Logo ----
def logo_corner(s, dark_bg=False, height=0.40, x=None, y=0.42):
    """Logo nhỏ góc trên phải. dark_bg=True dùng bản trắng (nền gradient/tối)."""
    from PIL import Image
    img = "logo_full_white.png" if dark_bg else "logo_full_transparent.png"
    p = os.path.join(ASSETS, img); lw, lh = Image.open(p).size
    w = height * lw / lh
    if x is None: x = SW - 0.55 - w
    s.shapes.add_picture(p, Inches(x), Inches(y), height=Inches(height))

def logo_big(s, x=0.6, y=0.55, height=0.85, white=True):
    img = "logo_full_white.png" if white else "logo_full_transparent.png"
    s.shapes.add_picture(os.path.join(ASSETS, img), Inches(x), Inches(y), height=Inches(height))

# ---- Khối dựng sẵn ----
def head(s, kicker, title, kicker_color=GREEN, title_color=BLACK, logo=True,
         kicker_weight="semibold", title_weight="semibold"):
    """Tiêu đề trang nội dung (nền trắng): kicker xanh nhỏ (SemiBold) + tiêu đề lớn (SemiBold theo brand book) + logo góc phải."""
    txt(s, kicker, 0.6, 0.5, 11.5, 0.4, 13, kicker_color, weight=kicker_weight, spacing=2)
    txt(s, title, 0.58, 0.9, 11.5, 0.85, 30, title_color, weight=title_weight, line=1.0)
    if logo: logo_corner(s, dark_bg=False)

def footer(s, section, page):
    """Footer chuẩn: tên brand (trái) · 'tài liệu · mục' (giữa) · số trang (phải)."""
    txt(s, DOC_TITLE, 0.55, 7.04, 3, 0.32, 9.5, BLACK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, section, 3.5, 7.04, 6.33, 0.32, 9.5, GRAY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, f"{page:02d}", 9.9, 7.04, 2.9, 0.32, 9.5, BLACK, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def cover(s, kicker, title, subtitle="", meta="", title_color=LEMON):
    """Slide bìa: nền gradient xanh bg_title + logo trắng + tiêu đề lớn màu vàng chanh (brand book)."""
    add_bg(s, "bg_title.png")
    logo_big(s, white=True)
    txt(s, kicker, 0.62, 2.45, 11, 0.5, 17, WHITE, bold=True, spacing=3)
    txt(s, title, 0.6, 2.95, 11.8, 2.0, 50, title_color, bold=True, line=1.02)
    if subtitle: txt(s, subtitle, 0.62, 5.45, 12, 0.5, 15.5, WHITE, italic=True)
    if meta: txt(s, meta, 0.62, 6.85, 9, 0.4, 12, WHITE)

def divider(s, part, title, sub="", num="", title_color=LEMON):
    """Trang ngăn section: nền gradient xanh bg_section + tiêu đề vàng chanh + số phần lớn."""
    add_bg(s, "bg_section.png")
    txt(s, part, 0.62, 2.25, 7, 0.5, 18, WHITE, bold=True, spacing=4)
    txt(s, title, 0.6, 2.78, 10.5, 1.5, 44, title_color, bold=True, line=1.03)
    if sub: txt(s, sub, 0.62, 4.35, 9, 0.7, 15.5, WHITE, line=1.12)
    if num: txt(s, num, 9.7, 4.95, 3.1, 2.0, 120, WHITE, bold=True, align=PP_ALIGN.RIGHT)
