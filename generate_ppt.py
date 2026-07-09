#!/usr/bin/env python3
"""
Generate professional McKinsey-style PPT from RESUME TUPOKSI egov.docx
Bidang e-Government — Diskominfostandi Kota Bekasi

Design principles (based on PPT best practices 2025-2026):
- Pyramid Principle: conclusion first, supporting evidence after
- Action Titles: every title communicates a conclusion
- One Message Per Slide
- 5/5/5 Rule: max 5 words/line, 5 lines/slide
- White space & minimalism
- Color restraint: every color has meaning
- Typography hierarchy
- Precision alignment
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ═══════════════════════════════════════════════════════════════════
#  COLOR PALETTE (McKinsey-inspired restrained palette)
# ═══════════════════════════════════════════════════════════════════
NAVY      = RGBColor(0x1A, 0x36, 0x5D)  # Primary — trustworthy, authoritative
BLUE      = RGBColor(0x2B, 0x6C, 0xB0)  # Secondary — action
LIGHT_BLUE = RGBColor(0xEB, 0xF4, 0xFC) # Background tint
GOLD      = RGBColor(0xD6, 0x9E, 0x2E)  # Accent — highlights
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
NEAR_WHITE = RGBColor(0xFA, 0xFA, 0xFA)
LIGHT_GRAY = RGBColor(0xF0, 0xF0, 0xF0)
MED_GRAY  = RGBColor(0xDD, 0xDD, 0xDD)
DARK_TEXT  = RGBColor(0x1A, 0x20, 0x2C)
MID_TEXT   = RGBColor(0x4A, 0x55, 0x68)
LIGHT_TEXT = RGBColor(0x8B, 0x97, 0xA8)
GREEN     = RGBColor(0x38, 0xA1, 0x69)
TEAL      = RGBColor(0x0B, 0x83, 0x9F)

# ═══════════════════════════════════════════════════════════════════
#  SLIDE DIMENSIONS (Widescreen 16:9)
# ═══════════════════════════════════════════════════════════════════
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# Margins
M = Inches(0.9)    # left/right margin
TOP = Inches(0.6)  # top margin
CONTENT_TOP = Inches(1.8)  # where body content starts
CONTENT_W = SLIDE_W - 2 * M  # usable width

# ═══════════════════════════════════════════════════════════════════
#  CONTENT DATA (from RESUME TUPOKSI egov.docx)
# ═══════════════════════════════════════════════════════════════════

# Slide 1 content
TITLE = "RESUME TUPOKSI"
SUBTITLE = "Bidang e-Government"

# 3 main pillars
PILLARS = [
    {
        "title": "Pengembangan Aplikasi\n& Sistem Informasi",
        "icon": "01",
        "desc": "Supervisi, standarisasi, dan pengelolaan\nsistem penghubung layanan pemerintah"
    },
    {
        "title": "Tata Kelola e-Government\n(SPBE)",
        "icon": "02",
        "desc": "Strategi, roadmap, arsitektur, dan\npeningkatan kapasitas SDM & GCIO"
    },
    {
        "title": "Pengembangan Kota Cerdas\n(Smart City)",
        "icon": "03",
        "desc": "Masterplan, kolaborasi lintas sektor,\ndan evaluasi program kota cerdas"
    },
]

# Sistem/Aplikasi
SISTEM_APLIKASI = [
    ("Web Pemerintah Kota", "Admin Teknis Web Pemkot & Web Diskominfo", TEAL),
    ("Mobile App PSW", "Platform Smart City (Pekan Smart City)", BLUE),
    ("Web Kota Cerdas", "Admin Teknis Portal Kota Cerdas", GREEN),
    ("Aplikasi SPLP", "Sistem Penghubung Layanan Pemerintah — Kemenkomdigi", GOLD),
]

# Detail: Pengembangan Aplikasi
APPS_DETAIL = [
    "Supervisi, analisis, dan standarisasi pengembangan aplikasi perangkat daerah",
    "Mengelola dan mengembangkan Sistem Penghubung Layanan Pemerintah (SPLP)",
    "Pengelolaan domain dan subdomain pemerintah daerah",
    "Sosialisasi dan peningkatan kapasitas SDM pengelola sistem",
]

# Detail: Smart City
SMART_CITY_DETAIL = [
    "Menyusun strategi, rencana aksi, dan masterplan Kota Cerdas",
    "Membangun kolaborasi dengan pemangku kepentingan (pemerintah, swasta, akademisi, masyarakat)",
    "Mengoordinasikan dan mengevaluasi program Kota Cerdas",
    "Menyelaraskan rencana induk dengan dokumen perencanaan daerah",
]

# Detail: SPBE
SPBE_DETAIL = [
    "Menyusun strategi, roadmap, arsitektur, dan peta rencana SPBE",
    "Melaksanakan dan mengoordinasikan program SPBE lintas perangkat daerah",
    "Mengembangkan kebijakan dan tata kelola SPBE",
    "Monitoring, evaluasi, dan rekomendasi perbaikan berkelanjutan",
]

# ═══════════════════════════════════════════════════════════════════
#  HELPER BUILDING BLOCKS
# ═══════════════════════════════════════════════════════════════════

def create_presentation():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs):
    """Add a blank slide."""
    return prs.slides.add_slide(prs.slide_layouts[6])


def solid_bg(slide, color):
    """Set solid background."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, l, t, w, h, color, border=None):
    """Add a filled rectangle."""
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    if border:
        s.line.color.rgb = border
        s.line.width = Pt(0.5)
    return s


def textbox(slide, l, t, w, h, text, size=16, bold=False, color=DARK_TEXT,
            align=PP_ALIGN.LEFT, name="Calibri", spacing_before=0, spacing_after=0):
    """Add a text box, return (shape, text_frame)."""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = name
    p.alignment = align
    if spacing_before:
        p.space_before = Pt(spacing_before)
    if spacing_after:
        p.space_after = Pt(spacing_after)
    return tb, tf


def add_paragraph(tf, text, size=16, bold=False, color=DARK_TEXT,
                  name="Calibri", align=PP_ALIGN.LEFT, space_after=0,
                  space_before=0, level=0):
    """Add a paragraph to existing text frame."""
    p = tf.add_paragraph()
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = name
    p.alignment = align
    p.level = level
    if space_after:
        p.space_after = Pt(space_after)
    if space_before:
        p.space_before = Pt(space_before)
    return p


def slide_header(prs, slide, title, subtitle=None, action_line=True):
    """Standard slide header: top bar + title + optional subtitle.

    Creates a consistent header system across all content slides.
    """
    # Thin gold accent line at very top
    if action_line:
        rect(slide, Inches(0), Inches(0), SLIDE_W, Pt(4), GOLD)

    # Title
    textbox(slide, M, Inches(0.5), CONTENT_W, Inches(0.7),
            title, size=30, bold=True, color=NAVY)

    # Thin separator under title
    rect(slide, M, Inches(1.25), Inches(1.5), Pt(2), GOLD)

    if subtitle:
        textbox(slide, M, Inches(1.35), CONTENT_W, Inches(0.4),
                subtitle, size=14, color=MID_TEXT)


def bullet_list(slide, left, top, width, items, size=15, color=DARK_TEXT,
                spacing=8, bullet_char="▸"):
    """Add a bulleted list."""
    tb, tf = textbox(slide, left, top, width, Inches(0.1), "", size=size)
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"{bullet_char}  {item}"
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.space_after = Pt(spacing)
    return tb


def footer(slide, text="Bidang e-Government — Diskominfostandi Kota Bekasi"):
    """Add page footer."""
    rect(slide, Inches(0), SLIDE_H - Inches(0.45), SLIDE_W, Inches(0.45), NEAR_WHITE)
    rect(slide, Inches(0), SLIDE_H - Inches(0.45), SLIDE_W, Pt(1), MED_GRAY)
    textbox(slide, M, SLIDE_H - Inches(0.38), CONTENT_W, Inches(0.3),
            text, size=9, color=LIGHT_TEXT)


# ═══════════════════════════════════════════════════════════════════
#  SLIDE BUILDERS
# ═══════════════════════════════════════════════════════════════════

def build_title_slide(prs):
    """Slide 1: Title — clean, impactful, minimal."""
    slide = blank_slide(prs)
    solid_bg(slide, NAVY)

    # Large decorative shape — subtle geometric accent
    s = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE,
                                Inches(9.5), Inches(0), Inches(5), Inches(7.5))
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor(0x15, 0x2D, 0x50)
    s.line.fill.background()
    s.rotation = 0.0

    # Smaller triangle overlay
    s2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE,
                                 Inches(10.2), Inches(1.5), Inches(4), Inches(6))
    s2.fill.solid()
    s2.fill.fore_color.rgb = RGBColor(0x12, 0x26, 0x45)
    s2.line.fill.background()

    # Thin gold line top
    rect(slide, Inches(0), Inches(0), SLIDE_W, Pt(3), GOLD)

    # Gold accent line under title area
    rect(slide, M, Inches(4.0), Inches(2.0), Pt(3), GOLD)

    # Main title
    textbox(slide, M, Inches(1.8), Inches(8), Inches(1.0),
            "RESUME TUPOKSI", size=52, bold=True, color=WHITE)

    # Subtitle
    textbox(slide, M, Inches(3.0), Inches(8), Inches(0.7),
            "Bidang e-Government", size=26, bold=False, color=RGBColor(0xBB, 0xD0, 0xE5))

    # Institution
    textbox(slide, M, Inches(4.3), Inches(8), Inches(0.5),
            "Diskominfostandi Kota Bekasi", size=16, color=RGBColor(0x88, 0xA8, 0xC8))

    # Year
    textbox(slide, M, Inches(5.5), Inches(3), Inches(0.4),
            "2025", size=13, color=LIGHT_TEXT)

    return slide


def build_summary_slide(prs):
    """Slide 2: Executive Summary — 3 Main Pillars (Pyramid Principle)."""
    slide = blank_slide(prs)
    slide_header(prs, slide, "Tiga Pilar Utama Tugas Pokok",
                 "Bidang e-Government mencakup 3 area strategis dalam transformasi digital daerah")

    # Three pillar cards
    card_w = Inches(3.6)
    card_h = Inches(4.0)
    gap = Inches(0.6)
    total_w = 3 * card_w + 2 * gap
    start_x = (SLIDE_W - total_w) / 2  # center
    card_top = Inches(2.3)

    colors = [NAVY, BLUE, TEAL]

    for i, pillar in enumerate(PILLARS):
        x = start_x + i * (card_w + gap)

        # Card background
        card = rect(slide, x, card_top, card_w, card_h, WHITE, MED_GRAY)

        # Color top bar
        rect(slide, x, card_top, card_w, Pt(5), colors[i])

        # Number circle
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                         x + Inches(0.25), card_top + Inches(0.4),
                                         Inches(0.55), Inches(0.55))
        circle.fill.solid()
        circle.fill.fore_color.rgb = colors[i]
        circle.line.fill.background()
        # Number text
        tf = circle.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = str(i + 1)
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].space_before = Pt(0)
        tf.paragraphs[0].space_after = Pt(0)

        # Title
        textbox(slide, x + Inches(0.95), card_top + Inches(0.4),
                card_w - Inches(1.2), Inches(0.8),
                pillar["title"], size=15, bold=True, color=NAVY, spacing_after=4)

        # Separator line
        rect(slide, x + Inches(0.25), card_top + Inches(1.4),
             card_w - Inches(0.5), Pt(1), MED_GRAY)

        # Description
        textbox(slide, x + Inches(0.25), card_top + Inches(1.6),
                card_w - Inches(0.5), Inches(1.8),
                pillar["desc"], size=12, color=MID_TEXT, spacing_after=6)

    footer(slide)
    return slide


def build_systems_slide(prs):
    """Slide 3: Sistem & Aplikasi yang Dikelola."""
    slide = blank_slide(prs)
    slide_header(prs, slide, "Empat Sistem & Aplikasi Strategis yang Dikelola",
                 "Layanan digital yang dioperasikan dan dipelihara oleh Bidang e-Government")

    # 2x2 grid of highlight cards
    card_w = Inches(5.5)
    card_h = Inches(2.2)
    gap_x = Inches(0.7)
    gap_y = Inches(0.5)
    total_grid_w = 2 * card_w + gap_x
    start_x = (SLIDE_W - total_grid_w) / 2
    start_y = Inches(2.3)

    for i, (title, desc, accent) in enumerate(SISTEM_APLIKASI):
        col = i % 2
        row = i // 2
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)

        # Card bg
        card = rect(slide, x, y, card_w, card_h, WHITE, MED_GRAY)

        # Left accent bar
        rect(slide, x, y, Pt(6), card_h, accent)

        # Icon placeholder (colored circle with icon text)
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                         x + Inches(0.35), y + Inches(0.5),
                                         Inches(0.7), Inches(0.7))
        circle.fill.solid()
        circle.fill.fore_color.rgb = accent
        circle.line.fill.background()
        # put letter in circle
        letter = title[0]
        tf = circle.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = letter
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER

        # Title
        textbox(slide, x + Inches(1.3), y + Inches(0.4),
                card_w - Inches(1.6), Inches(0.45),
                title, size=18, bold=True, color=NAVY)

        # Description
        textbox(slide, x + Inches(1.3), y + Inches(0.95),
                card_w - Inches(1.6), Inches(0.9),
                desc, size=13, color=MID_TEXT)

    footer(slide)
    return slide


def build_detail_slide(prs, title, subtitle, items, accent_color=NAVY):
    """Generic detail slide: action title + bullet list + visual accent."""
    slide = blank_slide(prs)
    slide_header(prs, slide, title, subtitle)

    # Left accent block for visual interest
    rect(slide, M, Inches(1.8), Pt(4), Inches(4.5), accent_color)

    # Bullet list — clean, minimal
    bullet_list(slide, M + Inches(0.3), Inches(1.9),
                CONTENT_W - Inches(0.3), items,
                size=15, color=DARK_TEXT, spacing=12,
                bullet_char="▸")

    footer(slide)
    return slide


def build_closing_slide(prs):
    """Slide 7: Closing."""
    slide = blank_slide(prs)
    solid_bg(slide, NAVY)

    # Gold top line
    rect(slide, Inches(0), Inches(0), SLIDE_W, Pt(3), GOLD)

    # Gold accent
    rect(slide, M, Inches(3.6), Inches(2.0), Pt(3), GOLD)

    textbox(slide, M, Inches(2.0), SLIDE_W - 2 * M, Inches(1.0),
            "Terima Kasih", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    textbox(slide, M, Inches(3.0), SLIDE_W - 2 * M, Inches(0.6),
            "Bidang e-Government — Diskominfostandi Kota Bekasi",
            size=18, color=RGBColor(0xBB, 0xD0, 0xE5), align=PP_ALIGN.CENTER)

    textbox(slide, M, Inches(4.0), SLIDE_W - 2 * M, Inches(0.5),
            "\"Mewujudkan tata kelola pemerintahan yang cerdas, terpadu, dan berkelanjutan\"",
            size=13, color=LIGHT_TEXT, align=PP_ALIGN.CENTER)

    return slide


# ═══════════════════════════════════════════════════════════════════
#  MAIN — Build and Save
# ═══════════════════════════════════════════════════════════════════

def main():
    prs = create_presentation()

    # Build slides
    build_title_slide(prs)                              # Slide 1
    build_summary_slide(prs)                            # Slide 2
    build_systems_slide(prs)                            # Slide 3
    build_detail_slide(                                  # Slide 4
        prs,
        "Supervisi & Pengembangan Aplikasi Daerah",
        "Lingkup kerja pengembangan sistem informasi dan aplikasi",
        APPS_DETAIL,
        accent_color=NAVY,
    )
    build_detail_slide(                                  # Slide 5
        prs,
        "Akselerasi Program Kota Cerdas",
        "Strategi dan kolaborasi menuju Smart City yang terintegrasi",
        SMART_CITY_DETAIL,
        accent_color=TEAL,
    )
    build_detail_slide(                                  # Slide 6
        prs,
        "Penguatan Tata Kelola SPBE",
        "Kerangka kerja menuju Sistem Pemerintahan Berbasis Elektronik yang matang",
        SPBE_DETAIL,
        accent_color=BLUE,
    )
    build_closing_slide(prs)                            # Slide 7

    # Save
    output_path = "/home/runner/pptx/coba_ppt/RESUME_TUPOKSI_egov.pptx"
    prs.save(output_path)
    print(f"✅ PPT berhasil dibuat: {output_path}")
    print(f"   Total slide: {len(prs.slides)}")
    print(f"   Ukuran: 16:9 Widescreen")


if __name__ == "__main__":
    main()
