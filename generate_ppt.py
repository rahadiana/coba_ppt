#!/usr/bin/env python3
"""
generate_ppt.py — RESUME TUPOKSI e-Government PPT Generator
=============================================================
Pakai engine src/ppt_engine.py.
Creative design dengan icon shapes + Unicode symbols.
NO circles/ovals — pakai kotak, diamond, line art.
"""

import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from pptx.dml.color import RGBColor
from ppt_engine import Engine, MSO_SHAPE, Pt, Inches
from pptx.enum.text import PP_ALIGN as PA

# ─── Color Palette: Ocean-inspired ───────────────────────────────
DARK    = RGBColor(0x0A, 0x16, 0x28)
MID     = RGBColor(0x0F, 0x2A, 0x4A)
BLUE    = RGBColor(0x1A, 0x56, 0x76)
TEAL    = RGBColor(0x0D, 0x94, 0x88)
TEAL_L  = RGBColor(0xE6, 0xF7, 0xF5)
GOLD    = RGBColor(0xF5, 0x9E, 0x0B)
GOLD_L  = RGBColor(0xFE, 0xF3, 0xC7)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
OFF_W   = RGBColor(0xF8, 0xFA, 0xFC)
LGRAY   = RGBColor(0xF1, 0xF5, 0xF9)
BORDER  = RGBColor(0xE2, 0xE8, 0xF0)
TDARK   = RGBColor(0x1E, 0x29, 0x3B)
TMID    = RGBColor(0x47, 0x55, 0x69)
TLIGHT  = RGBColor(0x94, 0xA3, 0xB8)
GREEN   = RGBColor(0x10, 0xB9, 0x81)  # emerald
PURPLE  = RGBColor(0x8B, 0x5C, 0xF6)  # violet accent
ORANGE  = RGBColor(0xF9, 0x73, 0x16)  # orange accent

# ─── Unicode Icons ──────────────────────────────────────────────
ICO_APP   = "▣"   # application
ICO_GEAR  = "⚙"   # settings/SPBE
ICO_CITY  = "🏛"   # city/smart city
ICO_WEB   = "🌐"   # web
ICO_PHONE = "📱"   # mobile
ICO_LINK  = "🔗"   # link/connection
ICO_STAR  = "◆"    # diamond star
ICO_BOX   = "▣"    # box
ICO_ARROW = "▶"    # arrow
ICO_CHECK = "✓"    # check
ICO_LIST  = "▸"    # list bullet

# ─── Shape Icon Builders ────────────────────────────────────────
def icon_box(slide, l, t, size, color, symbol="", sym_size=14, sym_color=None):
    """Square icon box with symbol inside."""
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(l), Inches(t),
                                Inches(size), Inches(size))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.adjustments[0] = 0.15

    if symbol:
        if sym_color is None: sym_color = WHITE
        tb = slide.shapes.add_textbox(Inches(l), Inches(t + size*0.15),
                                       Inches(size), Inches(size*0.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = symbol
        p.font.size = Pt(sym_size)
        p.font.color.rgb = sym_color
        p.font.name = "Calibri"
        p.alignment = PA.CENTER
    return s


def icon_diamond(slide, l, t, size, color):
    """Diamond shape icon."""
    # Use rotated square as diamond - python-pptx doesn't have native diamond
    s = slide.shapes.add_shape(MSO_SHAPE.DIAMOND,
                                Inches(l), Inches(t),
                                Inches(size), Inches(size))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def accent_bar(slide, l, t, w, h, color):
    """Thin accent bar."""
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(l), Inches(t),
                                Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def card(slide, l, t, w, h, fill=None):
    """Rounded card."""
    if fill is None: fill = WHITE
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(l), Inches(t),
                                Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = BORDER
    s.line.width = Pt(0.5)
    s.adjustments[0] = 0.06
    return s


def txt(slide, l, t, w, h, text, size=12, bold=False, color=None,
        align=PA.LEFT, name="Calibri"):
    if color is None: color = TDARK
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


def multi_txt(slide, l, t, w, h, lines, size=12, color=None,
              spacing=4, first_bold=False, first_size=None):
    if color is None: color = TDARK
    tb = slide.shapes.add_textbox(Inches(l), Inches(t),
                                   Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(first_size if (i == 0 and first_size) else size)
        p.font.bold = (first_bold and i == 0)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.space_after = Pt(spacing)
    return tb


# ══════════════════════════════════════════════════════════════════
#  MAIN BUILDER
# ══════════════════════════════════════════════════════════════════

def build():
    engine = Engine()
    L = engine.L
    from pptx import Presentation
    prs = Presentation()
    prs.slide_width = Inches(L.SLIDE_W)
    prs.slide_height = Inches(L.SLIDE_H)
    engine.prs = prs

    MX = L.MARGIN_H
    CW = L.cw
    PG = [0]

    def ns():
        return prs.slides.add_slide(prs.slide_layouts[6])

    def solid_bg(sl, c):
        bg = sl.background; bg.fill.solid(); bg.fill.fore_color.rgb = c

    def content_slide(title, sub=None):
        sl = ns()
        solid_bg(sl, WHITE)
        # Header gradient layers
        accent_bar(sl, 0, 0, L.SLIDE_W, 0.04, GOLD)
        accent_bar(sl, 0, 0.04, L.SLIDE_W, 1.15, DARK)
        accent_bar(sl, 0, 1.19, L.SLIDE_W, 0.03, TEAL)
        txt(sl, MX+0.15, 0.18, CW-0.3, 0.55, title, 22, bold=True, color=WHITE)
        if sub:
            txt(sl, MX+0.15, 0.72, CW-0.3, 0.35, sub, 9, color=TEAL_L)
        return sl

    def ft(sl):
        accent_bar(sl, 0, L.SLIDE_H-0.45, L.SLIDE_W, 0.45, OFF_W)
        accent_bar(sl, 0, L.SLIDE_H-0.45, L.SLIDE_W, Pt(1.5), BORDER)
        PG[0] += 1
        txt(sl, MX, L.SLIDE_H-0.38, 4, 0.22, "Sumber: RESUME TUPOKSI egov.docx", 7, color=TLIGHT)
        txt(sl, L.SLIDE_W-1.0, L.SLIDE_H-0.38, 0.8, 0.22, str(PG[0]), 8, color=TLIGHT, align=PA.RIGHT)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 1 — COVER
    # ══════════════════════════════════════════════════════════════
    sl = ns()
    solid_bg(sl, DARK)

    # Geometric pattern
    accent_bar(sl, 0, 0, L.SLIDE_W, 0.04, GOLD)

    # Diagonal-like blocks
    accent_bar(sl, L.SLIDE_W-5, -0.5, 5.5, 5, MID)
    accent_bar(sl, L.SLIDE_W-3.8, 3.5, 4.5, 4.5, MID)
    accent_bar(sl, -0.5, 6, 3.5, 2, MID)
    accent_bar(sl, -0.5, 0, 2.5, 1.5, MID)

    # Bottom bar
    accent_bar(sl, 0, L.SLIDE_H-0.28, L.SLIDE_W, 0.28, TEAL)

    # Main title with icon block
    icon_box(sl, MX, 1.6, 0.7, TEAL, ICO_BOX, 22)
    txt(sl, MX+1.0, 1.6, CW-1.5, 0.8, "RESUME TUPOKSI", 46, bold=True, color=WHITE)

    txt(sl, MX, 2.8, CW, 0.5, "Bidang e-Government", 22, color=TEAL_L)
    txt(sl, MX, 3.4, CW, 0.4, "Diskominfostandi Kota Bekasi", 14, color=TLIGHT)
    accent_bar(sl, MX, 4.0, 2.5, Pt(3), GOLD)
    txt(sl, MX, 4.3, CW, 0.3, "2025", 11, color=TLIGHT)

    PG[0] += 1

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 2 — 3 PILAR DENGAN ICON
    # ══════════════════════════════════════════════════════════════
    sl = content_slide("Tiga Pilar Strategis e-Government",
                       "Bidang e-Government menggerakkan 3 pilar utama transformasi digital Kota Bekasi")

    pillars = [
        ("01", ICO_APP, "Pengembangan Aplikasi\n& Sistem Informasi",
         "Mengawal standar & kualitas\npengembangan aplikasi serta\nmenghubungkan layanan daerah\nmelalui SPLP", DARK),
        ("02", ICO_GEAR, "Tata Kelola SPBE",
         "Merancang roadmap, arsitektur &\nkebijakan SPBE sekaligus\nmemperkuat kapasitas SDM\ndan peran Government CIO", BLUE),
        ("03", ICO_CITY, "Pengembangan\nKota Cerdas",
         "Menyusun masterplan, membangun\nkolaborasi lintas sektor, serta\nmengevaluasi program Smart City\nsecara berkelanjutan", TEAL),
    ]

    n = len(pillars)
    cw = L.col_width(n, 0.35)
    gap = 0.35
    sx = (L.SLIDE_W - (n*cw + (n-1)*gap)) / 2
    ch = 4.0
    cy = 1.55

    for i, (num, ico, title, desc, clr) in enumerate(pillars):
        cx = sx + i * (cw + gap)

        # Card
        card(sl, cx, cy, cw, ch)

        # Colored top bar
        accent_bar(sl, cx, cy, cw, 0.08, clr)

        # Icon square
        icon_box(sl, cx+0.25, cy+0.3, 0.55, clr, ico, 20)

        # Number label
        txt(sl, cx+0.95, cy+0.35, 0.6, 0.3, f"0{i+1}", 14, bold=True, color=clr)

        # Title
        txt(sl, cx+0.25, cy+1.05, cw-0.5, 0.85, title, 17, bold=True, color=DARK)

        # Separator
        accent_bar(sl, cx+0.25, cy+2.0, cw-0.5, Pt(1.5), BORDER)

        # Description
        txt(sl, cx+0.25, cy+2.2, cw-0.5, 1.5, desc, 11, color=TMID)

    ft(sl)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 3 — SISTEM & APLIKASI (with icons)
    # ══════════════════════════════════════════════════════════════
    sl = content_slide("Empat Sistem & Aplikasi Strategis",
                       "Layanan digital yang dioperasikan dan dipelihara oleh Bidang e-Government")

    systems = [
        (ICO_WEB,   "Web Pemerintah Kota", "Portal utama pemerintah kota\nsebagai pusat informasi & layanan publik", DARK),
        (ICO_PHONE, "Mobile App PSW", "Aplikasi mobile Pekan Smart City\nsebagai layanan kota cerdas terintegrasi", BLUE),
        (ICO_CITY,  "Web Kota Cerdas", "Portal informasi & layanan\nprogram Smart City Kota Bekasi", TEAL),
        (ICO_LINK,  "Aplikasi SPLP", "Sistem Penghubung Layanan Pemerintah\nterintegrasi dengan Kemenkomdigi", GOLD),
    ]

    cw2 = L.col_width(2, 0.4)
    rh2 = 1.8
    gh2 = 0.4
    gv2 = 0.35
    gsx = (L.SLIDE_W - (2*cw2+gh2))/2
    acy = 1.55

    for i, (ico, title, desc, clr) in enumerate(systems):
        col = i % 2
        row = i // 2
        cx = gsx + col * (cw2+gh2)
        cy = acy + row * (rh2+gv2)

        card(sl, cx, cy, cw2, rh2)

        # Icon box
        icon_box(sl, cx+0.25, cy+0.3, 0.65, clr, ico, 22)

        # Title
        txt(sl, cx+1.1, cy+0.3, cw2-1.4, 0.45, title, 18, bold=True, color=clr)

        # Description
        txt(sl, cx+1.1, cy+0.85, cw2-1.4, 0.7, desc, 12, color=TMID)

    ft(sl)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 4-6 — DETAIL CONTENT with icon blocks
    # ══════════════════════════════════════════════════════════════

    details = [
        ("Supervisi & Pengembangan Aplikasi Daerah",
         "Lingkup kerja pengembangan sistem informasi dan aplikasi perangkat daerah",
         [ICO_CHECK, ICO_LINK, ICO_WEB, ICO_GEAR],
         ["Supervisi, analisis, dan standarisasi\npengembangan aplikasi perangkat daerah",
          "Mengelola & mengembangkan SPLP\nsebagai tulang punggung integrasi layanan",
          "Pengelolaan domain dan subdomain\npemerintah daerah secara terpusat",
          "Sosialisasi & peningkatan kapasitas\nSDM pengelola sistem informasi"]),
        ("Akselerasi Program Kota Cerdas",
         "Strategi dan kolaborasi menuju Smart City yang terintegrasi",
         [ICO_STAR, ICO_CITY, ICO_APP, ICO_LINK],
         ["Menyusun strategi, rencana aksi,\ndan masterplan Kota Cerdas",
          "Membangun kolaborasi dengan pemangku\nkepentingan lintas sektor",
          "Mengoordinasikan & mengevaluasi\nprogram Kota Cerdas secara berkala",
          "Menyelaraskan rencana induk dengan\ndokumen perencanaan daerah"]),
        ("Penguatan Tata Kelola SPBE",
         "Kerangka kerja menuju Sistem Pemerintahan Berbasis Elektronik yang matang",
         [ICO_GEAR, ICO_LIST, ICO_BOX, ICO_STAR],
         ["Menyusun strategi, roadmap, arsitektur,\ndan peta rencana SPBE",
          "Melaksanakan & mengoordinasikan\nprogram SPBE lintas perangkat daerah",
          "Mengembangkan kebijakan &\ntata kelola SPBE yang adaptif",
          "Monitoring, evaluasi, & rekomendasi\nperbaikan SPBE berkelanjutan"]),
    ]

    accent_colors = [DARK, TEAL, BLUE]

    for si, (title, sub, icons, items) in enumerate(details):
        sl = content_slide(title, sub)
        accent = accent_colors[si]

        # Left panel
        accent_bar(sl, MX, 1.5, 0.1, 5.0, accent)

        # Items with icon boxes
        for i, (ico, item) in enumerate(zip(icons, items)):
            iy = 1.75 + i * 1.15

            # Number box
            icon_box(sl, MX+0.35, iy, 0.5, accent, ico, 16)

            # Vertical connector
            if i < len(icons)-1:
                accent_bar(sl, MX+0.58, iy+0.55, Pt(1.5), 0.55, BORDER)

            # Title (bold first line)
            lines = item.split('\n')
            multi_txt(sl, MX+1.1, iy+0.02, CW-1.4, 0.8, lines,
                      13, color=TDARK, spacing=1, first_bold=True, first_size=14)

        ft(sl)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 7 — CLOSING with icon
    # ══════════════════════════════════════════════════════════════
    sl = ns()
    solid_bg(sl, DARK)

    accent_bar(sl, 0, 0, L.SLIDE_W, 0.04, GOLD)
    accent_bar(sl, L.SLIDE_W-4.5, -0.5, 5, 5, MID)
    accent_bar(sl, L.SLIDE_W-3.2, 4, 4, 4, MID)
    accent_bar(sl, 0, L.SLIDE_H-0.28, L.SLIDE_W, 0.28, TEAL)

    # Big icon
    icon_box(sl, L.SLIDE_W/2-0.45, 1.3, 0.9, TEAL, ICO_STAR, 30)

    txt(sl, MX, 2.5, CW, 1.0, "Terima Kasih", 44, bold=True, color=WHITE, align=PA.CENTER)
    txt(sl, MX, 3.6, CW, 0.5, "Bidang e-Government — Diskominfostandi Kota Bekasi",
        18, color=TEAL_L, align=PA.CENTER)
    accent_bar(sl, L.SLIDE_W/2-1.0, 4.3, 2.0, Pt(2), GOLD)
    txt(sl, MX, 4.6, CW, 0.5,
        "\"Mewujudkan tata kelola pemerintahan yang cerdas,\nterpadu, dan berkelanjutan\"",
        12, color=TLIGHT, align=PA.CENTER)

    PG[0] += 1

    # ── Save ──
    out = "RESUME_TUPOKSI_egov.pptx"
    prs.save(out)
    print(f"✅ PPT selesai: {out}")
    print(f"   {len(prs.slides)} slide — creative design with icons!")
    return out


if __name__ == "__main__":
    build()
