#!/usr/bin/env python3
"""
generate_ppt.py — RESUME TUPOKSI e-Government PPT Generator
=============================================================
Pakai engine src/ppt_engine.py — tanpa icon/oval/circle.
Desain: modern, clean, attractive — gradient-style layering.
"""

import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from pptx.dml.color import RGBColor
from ppt_engine import Engine, MSO_SHAPE, Pt, Inches
from pptx.oxml.ns import qn
from lxml import etree


# ─── Color Palette: Ocean-inspired ───────────────────────────────
OCEAN_DARK   = RGBColor(0x0A, 0x16, 0x28)  # deepest navy
OCEAN_MID    = RGBColor(0x0F, 0x2A, 0x4A)  # mid navy
OCEAN_BLUE   = RGBColor(0x1A, 0x56, 0x76)  # blue-teal
TEAL          = RGBColor(0x0D, 0x94, 0x88)  # bright teal
TEAL_LIGHT   = RGBColor(0xE6, 0xF7, 0xF5)  # teal tint
GOLD          = RGBColor(0xF5, 0x9E, 0x0B)  # warm amber
GOLD_LIGHT   = RGBColor(0xFE, 0xF3, 0xC7)  # amber tint
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE    = RGBColor(0xF8, 0xFA, 0xFC)
LIGHT_GRAY   = RGBColor(0xF1, 0xF5, 0xF9)
BORDER_LIGHT = RGBColor(0xE2, 0xE8, 0xF0)
DARK_TEXT     = RGBColor(0x1E, 0x29, 0x3B)
MID_TEXT      = RGBColor(0x47, 0x55, 0x69)
LIGHT_TEXT    = RGBColor(0x94, 0xA3, 0xB8)

# Card accent colors (for variety)
ACCENTS = [OCEAN_DARK, OCEAN_BLUE, TEAL, GOLD]


def build_presentation():
    engine = Engine()
    C = engine.C
    L = engine.L

    from pptx import Presentation
    from pptx.enum.text import PP_ALIGN as PA

    prs = Presentation()
    prs.slide_width = Inches(L.SLIDE_W)
    prs.slide_height = Inches(L.SLIDE_H)
    engine.prs = prs

    MX = L.MARGIN_H       # 0.6"
    CW = L.cw              # 12.133"
    PG = [0]
    BLANK = prs.slide_layouts[6]

    # ─── Helper Functions ──────────────────────────────────────

    def new_slide():
        return prs.slides.add_slide(BLANK)

    def solid_bg(slide, color):
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = color

    def rect(slide, l, t, w, h, color):
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(l), Inches(t),
                                    Inches(w), Inches(h))
        s.fill.solid()
        s.fill.fore_color.rgb = color
        s.line.fill.background()
        return s

    def rrect(slide, l, t, w, h, fill=None, border=None, radius=0.04):
        if fill is None: fill = WHITE
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(l), Inches(t),
                                     Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if border:
            sh.line.color.rgb = border
            sh.line.width = Pt(0.5)
        else:
            sh.line.fill.background()
        sh.adjustments[0] = radius
        return sh

    def box(slide, l, t, w, h, text, size=12, bold=False, color=None,
            align=PA.LEFT, name="Calibri"):
        if color is None: color = DARK_TEXT
        tb = slide.shapes.add_textbox(Inches(l), Inches(t),
                                       Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.font.name = name
        p.alignment = align
        return tb

    def multi_line(slide, l, t, w, h, lines, size=12, color=None,
                   spacing=8, bold_first=False):
        """Text box with multiple lines/paragraphs."""
        if color is None: color = DARK_TEXT
        tb = slide.shapes.add_textbox(Inches(l), Inches(t),
                                       Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(size)
            p.font.bold = (bold_first and i == 0)
            p.font.color.rgb = color
            p.font.name = "Calibri"
            p.space_after = Pt(spacing)
        return tb

    def footer(slide):
        rect(slide, 0, L.SLIDE_H - 0.45, L.SLIDE_W, 0.45, OFF_WHITE)
        rect(slide, 0, L.SLIDE_H - 0.45, L.SLIDE_W, Pt(1.5), BORDER_LIGHT)
        PG[0] += 1
        box(slide, MX, L.SLIDE_H - 0.38, 4, 0.22,
            "Sumber: RESUME TUPOKSI egov.docx", 7, color=LIGHT_TEXT)
        box(slide, L.SLIDE_W - 1.0, L.SLIDE_H - 0.38, 0.8, 0.22,
            str(PG[0]), 8, color=LIGHT_TEXT, align=PA.RIGHT)

    def content_slide(slide_title, subtitle=None):
        """Modern content slide with gradient-style header."""
        s = new_slide()
        solid_bg(s, WHITE)

        # Top gradient effect (layered rectangles)
        rect(s, 0, 0, L.SLIDE_W, 0.04, GOLD)        # gold accent line
        rect(s, 0, 0.04, L.SLIDE_W, 1.15, OCEAN_DARK)  # dark header
        rect(s, 0, 1.19, L.SLIDE_W, 0.03, TEAL)      # teal separator

        # Title
        box(s, MX + 0.15, 0.18, CW - 0.3, 0.55,
            slide_title, 22, bold=True, color=WHITE)

        # Subtitle
        if subtitle:
            box(s, MX + 0.15, 0.72, CW - 0.3, 0.35,
                subtitle, 9, color=TEAL_LIGHT)

        return s

    # ══════════════════════════════════════════════════════════
    #  SLIDE 1 — COVER
    # ══════════════════════════════════════════════════════════
    s = new_slide()
    solid_bg(s, OCEAN_DARK)

    # Geometric layered blocks for visual depth
    rect(s, 0, 0, L.SLIDE_W, 0.04, GOLD)
    rect(s, L.SLIDE_W - 4.5, -0.5, 5, 4.5, OCEAN_MID)
    rect(s, L.SLIDE_W - 3.5, 3.2, 4, 4.5, OCEAN_MID)
    rect(s, -0.5, 5.5, 3, 2.5, OCEAN_MID)

    # Bottom bar
    rect(s, 0, L.SLIDE_H - 0.25, L.SLIDE_W, 0.25, TEAL)

    box(s, MX, 2.0, CW, 1.0, "RESUME TUPOKSI", 46, bold=True, color=WHITE)
    box(s, MX, 3.2, CW, 0.5, "Bidang e-Government", 22, color=TEAL_LIGHT)
    box(s, MX, 3.8, CW, 0.4, "Diskominfostandi Kota Bekasi", 14, color=LIGHT_TEXT)
    rect(s, MX, 4.4, 2.5, Pt(3), GOLD)
    box(s, MX, 4.7, CW, 0.3, "2025", 11, color=LIGHT_TEXT)

    PG[0] += 1

    # ══════════════════════════════════════════════════════════
    #  SLIDE 2 — 3 PILAR (clean, centered, no icons)
    # ══════════════════════════════════════════════════════════
    s = content_slide("Tiga Pilar Utama Tugas Pokok",
                      "Bidang e-Government mencakup 3 area strategis dalam transformasi digital daerah")

    pillars = [
        ("01", "Pengembangan Aplikasi\n& Sistem Informasi",
         "Supervisi, standarisasi, dan\npengelolaan SPLP", OCEAN_DARK),
        ("02", "Tata Kelola e-Government\n(SPBE)",
         "Strategi, roadmap, arsitektur,\ndan kapasitas SDM & GCIO", OCEAN_BLUE),
        ("03", "Pengembangan Kota Cerdas\n(Smart City)",
         "Masterplan, kolaborasi lintas\nsektor, dan evaluasi program", TEAL),
    ]

    n = len(pillars)
    card_w = L.col_width(n, 0.35)
    gap = 0.35
    start_x = (L.SLIDE_W - (n * card_w + (n-1)*gap)) / 2
    card_h = 3.6
    card_y = 1.6

    for i, (num, title, desc, clr) in enumerate(pillars):
        cx = start_x + i * (card_w + gap)

        # Card
        rrect(s, cx, card_y, card_w, card_h, fill=WHITE, border=BORDER_LIGHT)

        # Top color block (like a header band)
        rect(s, cx, card_y, card_w, 0.65, clr)

        # Number in white on the color block
        box(s, cx + 0.2, card_y + 0.12, 0.5, 0.35, num, 20, bold=True, color=WHITE)

        # Sub-label "PILAR" 
        box(s, cx + 0.7, card_y + 0.15, 1.5, 0.25, "PILAR UTAMA", 7, color=TEAL_LIGHT)

        # Title below the block
        box(s, cx + 0.25, card_y + 0.85, card_w - 0.5, 0.85,
            title, 16, bold=True, color=OCEAN_DARK)

        # Thin separator
        rect(s, cx + 0.25, card_y + 1.8, card_w - 0.5, Pt(1.5), BORDER_LIGHT)

        # Description
        box(s, cx + 0.25, card_y + 2.0, card_w - 0.5, 1.2,
            desc, 12, color=MID_TEXT)

    footer(s)

    # ══════════════════════════════════════════════════════════
    #  SLIDE 3 — SISTEM & APLIKASI (2x2 grid, refined)
    # ══════════════════════════════════════════════════════════
    s = content_slide("Empat Sistem & Aplikasi Strategis",
                      "Layanan digital yang dioperasikan dan dipelihara oleh Bidang e-Government")

    systems = [
        ("Web Pemerintah Kota", "Admin Teknis Web Pemkot\n& Web Diskominfo", OCEAN_DARK),
        ("Mobile App PSW", "Platform Smart City\n— Pekan Smart City", OCEAN_BLUE),
        ("Web Kota Cerdas", "Admin Teknis\nPortal Kota Cerdas", TEAL),
        ("Aplikasi SPLP", "Sistem Penghubung Layanan\nPemerintah (Kemenkomdigi)", GOLD),
    ]

    cw2 = L.col_width(2, 0.4)
    rh2 = 1.7
    gh2 = 0.4
    gv2 = 0.3
    grid_sx = (L.SLIDE_W - (2*cw2 + gh2)) / 2
    ac_y = 1.6

    for i, (title, desc, clr) in enumerate(systems):
        col = i % 2
        row = i // 2
        cx = grid_sx + col * (cw2 + gh2)
        cy = ac_y + row * (rh2 + gv2)

        # Card
        rrect(s, cx, cy, cw2, rh2, fill=WHITE, border=BORDER_LIGHT)

        # Left accent strip
        rect(s, cx, cy, 0.08, rh2, clr)

        # Title
        box(s, cx + 0.3, cy + 0.25, cw2 - 0.5, 0.4,
            title, 18, bold=True, color=clr)

        # Description
        box(s, cx + 0.3, cy + 0.75, cw2 - 0.5, 0.7,
            desc, 12, color=MID_TEXT)

    footer(s)

    # ══════════════════════════════════════════════════════════
    #  SLIDE 4-6 — DETAIL CONTENT (modern numbered layout)
    # ══════════════════════════════════════════════════════════

    details = [
        ("Supervisi & Pengembangan Aplikasi Daerah",
         "Lingkup kerja pengembangan sistem informasi dan aplikasi perangkat daerah",
         ["Supervisi, analisis, dan standarisasi pengembangan\n aplikasi perangkat daerah",
          "Mengelola dan mengembangkan Sistem Penghubung\n Layanan Pemerintah (SPLP)",
          "Pengelolaan domain dan subdomain\n pemerintah daerah",
          "Sosialisasi dan peningkatan kapasitas\n SDM pengelola sistem"]),
        ("Akselerasi Program Kota Cerdas",
         "Strategi dan kolaborasi menuju Smart City yang terintegrasi",
         ["Menyusun strategi, rencana aksi,\n dan masterplan Kota Cerdas",
          "Membangun kolaborasi dengan pemangku\n kepentingan lintas sektor",
          "Mengoordinasikan dan mengevaluasi\n program Kota Cerdas",
          "Menyelaraskan rencana induk dengan\n dokumen perencanaan daerah"]),
        ("Penguatan Tata Kelola SPBE",
         "Kerangka kerja menuju Sistem Pemerintahan\n Berbasis Elektronik yang matang",
         ["Menyusun strategi, roadmap, arsitektur,\n dan peta rencana SPBE",
          "Melaksanakan dan mengoordinasikan program\n SPBE lintas perangkat daerah",
          "Mengembangkan kebijakan dan\n tata kelola SPBE",
          "Monitoring, evaluasi, dan rekomendasi\n perbaikan berkelanjutan"]),
    ]

    accent_colors = [OCEAN_DARK, TEAL, OCEAN_BLUE]

    for slide_idx, (title, subtitle, items) in enumerate(details):
        s = content_slide(title, subtitle)
        accent = accent_colors[slide_idx]

        # Left decorative panel
        rect(s, MX, 1.5, 0.12, 4.8, accent)
        rect(s, MX, 1.5, 1.8, 0.06, accent)  # horizontal accent at top

        # Numbered items with blocks
        for i, item in enumerate(items):
            iy = 1.8 + i * 1.1

            # Number block
            rect(s, MX + 0.4, iy, 0.5, 0.5, accent)

            # Number text
            box(s, MX + 0.4, iy + 0.06, 0.5, 0.4, str(i+1),
                18, bold=True, color=WHITE, align=PA.CENTER)

            # Item text
            box(s, MX + 1.15, iy + 0.02, CW - 1.4, 0.8,
                item, 13, color=DARK_TEXT)

            # Subtle connector line
            if i < len(items) - 1:
                rect(s, MX + 0.65, iy + 0.55, Pt(1.5), 0.5, BORDER_LIGHT)

        footer(s)

    # ══════════════════════════════════════════════════════════
    #  SLIDE 7 — CLOSING
    # ══════════════════════════════════════════════════════════
    s = new_slide()
    solid_bg(s, OCEAN_DARK)

    # Geometric accents
    rect(s, 0, 0, L.SLIDE_W, 0.04, GOLD)
    rect(s, L.SLIDE_W - 4, -0.5, 5, 4, OCEAN_MID)
    rect(s, L.SLIDE_W - 3, 4, 3.5, 4, OCEAN_MID)
    rect(s, 0, L.SLIDE_H - 0.25, L.SLIDE_W, 0.25, TEAL)

    box(s, MX, 2.2, CW, 1.0, "Terima Kasih", 44, bold=True, color=WHITE, align=PA.CENTER)
    box(s, MX, 3.4, CW, 0.5,
        "Bidang e-Government — Diskominfostandi Kota Bekasi",
        18, color=TEAL_LIGHT, align=PA.CENTER)
    rect(s, L.SLIDE_W/2 - 1.0, 4.1, 2.0, Pt(2), GOLD)
    box(s, MX, 4.4, CW, 0.5,
        "\"Mewujudkan tata kelola pemerintahan yang cerdas,\n terpadu, dan berkelanjutan\"",
        12, color=LIGHT_TEXT, align=PA.CENTER)

    PG[0] += 1

    # ── Save ──
    output = "RESUME_TUPOKSI_egov.pptx"
    prs.save(output)
    print(f"✅ PPT berhasil: {output}")
    print(f"   {len(prs.slides)} slide — modern attractive design")
    return output


if __name__ == "__main__":
    build_presentation()
