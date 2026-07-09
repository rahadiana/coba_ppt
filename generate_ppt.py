#!/usr/bin/env python3
"""
generate_ppt.py — RESUME TUPOKSI e-Government
==============================================
Pakai engine src/ppt_engine.py.
Icon pake EMOJI — simple, colorful, render di PowerPoint mana aja.
NO shape icons, NO circles/ovals.
"""

import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from pptx.dml.color import RGBColor
from ppt_engine import Engine, MSO_SHAPE, Pt, Inches
from pptx.enum.text import PP_ALIGN as PA

# ─── Color Palette ───────────────────────────────────────────────
DARK    = RGBColor(0x0A, 0x16, 0x28)
MID     = RGBColor(0x0F, 0x2A, 0x4A)
BLUE    = RGBColor(0x1A, 0x56, 0x76)
TEAL    = RGBColor(0x0D, 0x94, 0x88)
TEAL_L  = RGBColor(0xE6, 0xF7, 0xF5)
GOLD    = RGBColor(0xF5, 0x9E, 0x0B)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
OFF_W   = RGBColor(0xF8, 0xFA, 0xFC)
BORDER  = RGBColor(0xE2, 0xE8, 0xF0)
TDARK   = RGBColor(0x1E, 0x29, 0x3B)
TMID    = RGBColor(0x47, 0x55, 0x69)
TLIGHT  = RGBColor(0x94, 0xA3, 0xB8)


def add_shape(slide, st, l, t, w, h, fill):
    s = slide.shapes.add_shape(st, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s


def rect(slide, l, t, w, h, c):
    return add_shape(slide, MSO_SHAPE.RECTANGLE, l, t, w, h, c)


def card(slide, l, t, w, h):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = WHITE
    s.line.color.rgb = BORDER; s.line.width = Pt(0.5)
    s.adjustments[0] = 0.06
    return s


def txt(slide, l, t, w, h, text, size=12, bold=False, color=None,
        align=PA.LEFT, font="Calibri"):
    if color is None: color = TDARK
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text; p.font.size = Pt(size)
    p.font.bold = bold; p.font.color.rgb = color
    p.font.name = font; p.alignment = align
    return tb


def icon_txt(slide, l, t, s, emoji, bg_color, txt_color=None):
    """Emoji icon inside a rounded square background."""
    if txt_color is None: txt_color = WHITE
    # Background square
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, l, t, s, s, bg_color)
    # Emoji on top
    tb = slide.shapes.add_textbox(Inches(l), Inches(t + s*0.08),
                                   Inches(s), Inches(s*0.85))
    tf = tb.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = emoji; p.font.size = Pt(int(s * 36))
    p.font.color.rgb = txt_color; p.alignment = PA.CENTER
    # Use Segoe UI Emoji for better emoji rendering
    p.font.name = "Segoe UI Emoji"
    return l + s + 0.2


def multi_txt(slide, l, t, w, h, lines, size=12, color=None,
              spacing=2, first_bold=False, first_size=None):
    if color is None: color = TDARK
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(first_size if (i == 0 and first_size) else size)
        p.font.bold = (first_bold and i == 0)
        p.font.color.rgb = color; p.font.name = "Calibri"
        p.space_after = Pt(spacing)
    return tb


def build():
    engine = Engine()
    L = engine.L

    from pptx import Presentation
    prs = Presentation()
    prs.slide_width = Inches(L.SLIDE_W)
    prs.slide_height = Inches(L.SLIDE_H)

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
        rect(sl, 0, 0, L.SLIDE_W, 0.04, GOLD)
        rect(sl, 0, 0.04, L.SLIDE_W, 1.15, DARK)
        rect(sl, 0, 1.19, L.SLIDE_W, 0.03, TEAL)
        txt(sl, MX+0.15, 0.18, CW-0.3, 0.55, title, 22, bold=True, color=WHITE)
        if sub:
            txt(sl, MX+0.15, 0.72, CW-0.3, 0.35, sub, 9, color=TEAL_L)
        return sl

    def ft(sl):
        rect(sl, 0, L.SLIDE_H-0.45, L.SLIDE_W, 0.45, OFF_W)
        rect(sl, 0, L.SLIDE_H-0.45, L.SLIDE_W, Pt(1.5), BORDER)
        PG[0] += 1
        txt(sl, MX, L.SLIDE_H-0.38, 4, 0.22, "Sumber: RESUME TUPOKSI egov.docx", 7, color=TLIGHT)
        txt(sl, L.SLIDE_W-1.0, L.SLIDE_H-0.38, 0.8, 0.22, str(PG[0]), 8, color=TLIGHT, align=PA.RIGHT)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 1 — COVER
    # ══════════════════════════════════════════════════════════════
    sl = ns()
    solid_bg(sl, DARK)
    rect(sl, 0, 0, L.SLIDE_W, 0.04, GOLD)
    rect(sl, L.SLIDE_W-5, -0.5, 5.5, 5, MID)
    rect(sl, L.SLIDE_W-3.8, 3.5, 4.5, 4.5, MID)
    rect(sl, 0, L.SLIDE_H-0.28, L.SLIDE_W, 0.28, TEAL)

    icon_txt(sl, MX+0.3, 1.5, 0.8, "📊", TEAL)
    txt(sl, MX+1.5, 1.6, CW-2, 0.8, "RESUME TUPOKSI", 46, bold=True, color=WHITE)
    txt(sl, MX, 2.9, CW, 0.5, "Bidang e-Government", 22, color=TEAL_L)
    txt(sl, MX, 3.5, CW, 0.4, "Diskominfostandi Kota Bekasi", 14, color=TLIGHT)
    rect(sl, MX, 4.1, 2.5, Pt(3), GOLD)
    txt(sl, MX, 4.4, CW, 0.3, "2025", 11, color=TLIGHT)
    PG[0] += 1

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 2 — 3 PILAR
    # ══════════════════════════════════════════════════════════════
    sl = content_slide("Tiga Pilar Strategis e-Government",
                       "Bidang e-Government menggerakkan 3 pilar utama transformasi digital Kota Bekasi")

    pillars = [
        ("01", "💻", "Pengembangan Aplikasi\n& Sistem Informasi",
         "Mengawal standar & kualitas\npengembangan aplikasi serta\nmenghubungkan layanan daerah\nmelalui SPLP", DARK),
        ("02", "⚙️", "Tata Kelola SPBE",
         "Merancang roadmap, arsitektur &\nkebijakan SPBE sekaligus\nmemperkuat kapasitas SDM\ndan peran Government CIO", BLUE),
        ("03", "🌆", "Pengembangan\nKota Cerdas",
         "Menyusun masterplan, membangun\nkolaborasi lintas sektor, serta\nmengevaluasi program Smart City\nsecara berkelanjutan", TEAL),
    ]

    n = len(pillars)
    cw = L.col_width(n, 0.35)
    gap = 0.35
    sx = (L.SLIDE_W - (n*cw + (n-1)*gap)) / 2
    ch = 4.0
    cy = 1.55

    for i, (num, emoji, title, desc, clr) in enumerate(pillars):
        cx = sx + i * (cw + gap)
        card(sl, cx, cy, cw, ch)
        rect(sl, cx, cy, cw, 0.08, clr)
        icon_txt(sl, cx+0.25, cy+0.3, 0.55, emoji, clr)
        txt(sl, cx+0.95, cy+0.38, 0.5, 0.3, f"0{i+1}", 14, bold=True, color=clr)
        txt(sl, cx+0.25, cy+1.1, cw-0.5, 0.85, title, 17, bold=True, color=DARK)
        rect(sl, cx+0.25, cy+2.05, cw-0.5, Pt(1.5), BORDER)
        txt(sl, cx+0.25, cy+2.25, cw-0.5, 1.5, desc, 11, color=TMID)

    ft(sl)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 3 — SISTEM & APLIKASI
    # ══════════════════════════════════════════════════════════════
    sl = content_slide("Empat Sistem & Aplikasi Strategis",
                       "Layanan digital yang dioperasikan dan dipelihara oleh Bidang e-Government")

    systems = [
        ("🌐", "Web Pemerintah Kota",  "Portal utama pemerintah kota\nsebagai pusat info & layanan publik", DARK),
        ("📱", "Mobile App PSW",        "Aplikasi mobile Pekan Smart City\n— layanan kota cerdas terintegrasi", BLUE),
        ("🌆", "Web Kota Cerdas",       "Portal informasi & layanan\nprogram Smart City Kota Bekasi", TEAL),
        ("🔗", "Aplikasi SPLP",         "Sistem Penghubung Layanan\nPemerintah — terintegrasi Kemenkomdigi", GOLD),
    ]

    cw2 = L.col_width(2, 0.4)
    rh2 = 1.8
    gh2 = 0.4; gv2 = 0.35
    gsx = (L.SLIDE_W - (2*cw2+gh2))/2
    acy = 1.55

    for i, (emoji, title, desc, clr) in enumerate(systems):
        col = i % 2; row = i // 2
        cx = gsx + col*(cw2+gh2); cy = acy + row*(rh2+gv2)
        card(sl, cx, cy, cw2, rh2)
        icon_txt(sl, cx+0.25, cy+0.3, 0.65, emoji, clr)
        txt(sl, cx+1.15, cy+0.3, cw2-1.5, 0.45, title, 18, bold=True, color=clr)
        txt(sl, cx+1.15, cy+0.85, cw2-1.5, 0.7, desc, 12, color=TMID)

    ft(sl)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 4-6 — DETAIL with emoji icons
    # ══════════════════════════════════════════════════════════════

    details = [
        ("Supervisi & Pengembangan Aplikasi Daerah",
         "Lingkup kerja pengembangan sistem informasi dan aplikasi perangkat daerah",
         DARK,
         ["✅", "🔗", "🌐", "👥"],
         ["Supervisi, analisis, dan standarisasi\npengembangan aplikasi perangkat daerah",
          "Mengelola & mengembangkan SPLP\nsebagai tulang punggung integrasi layanan",
          "Pengelolaan domain & subdomain\npemerintah daerah secara terpusat",
          "Sosialisasi & peningkatan kapasitas\nSDM pengelola sistem informasi"]),
        ("Akselerasi Program Kota Cerdas",
         "Strategi dan kolaborasi menuju Smart City yang terintegrasi",
         TEAL,
         ["⭐", "🌆", "💻", "🔗"],
         ["Menyusun strategi, rencana aksi,\ndan masterplan Kota Cerdas",
          "Membangun kolaborasi dengan pemangku\nkepentingan lintas sektor",
          "Mengoordinasikan & mengevaluasi\nprogram Kota Cerdas secara berkala",
          "Menyelaraskan rencana induk dengan\ndokumen perencanaan daerah"]),
        ("Penguatan Tata Kelola SPBE",
         "Kerangka kerja menuju SPBE yang matang",
         BLUE,
         ["⚙️", "📋", "📝", "📊"],
         ["Menyusun strategi, roadmap, arsitektur,\ndan peta rencana SPBE",
          "Melaksanakan & mengoordinasikan\nprogram SPBE lintas perangkat daerah",
          "Mengembangkan kebijakan &\ntata kelola SPBE yang adaptif",
          "Monitoring, evaluasi, & rekomendasi\nperbaikan SPBE berkelanjutan"]),
    ]

    for si, (title, sub, accent, emojis, items) in enumerate(details):
        sl = content_slide(title, sub)
        rect(sl, MX, 1.5, 0.1, 5.2, accent)

        for i, (emoji, item) in enumerate(zip(emojis, items)):
            iy = 1.8 + i * 1.15
            icon_txt(sl, MX+0.35, iy, 0.5, emoji, accent)
            if i < len(emojis)-1:
                rect(sl, MX+0.6, iy+0.6, Pt(1.5), 0.5, BORDER)
            lines = item.split('\n')
            multi_txt(sl, MX+1.15, iy+0.04, CW-1.5, 0.8, lines,
                      13, color=TDARK, spacing=1, first_bold=True, first_size=14)

        ft(sl)

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 7 — CLOSING
    # ══════════════════════════════════════════════════════════════
    sl = ns()
    solid_bg(sl, DARK)
    rect(sl, 0, 0, L.SLIDE_W, 0.04, GOLD)
    rect(sl, L.SLIDE_W-4.5, -0.5, 5, 5, MID)
    rect(sl, 0, L.SLIDE_H-0.28, L.SLIDE_W, 0.28, TEAL)

    icon_txt(sl, L.SLIDE_W/2-0.4, 1.2, 0.8, "⭐", TEAL)
    txt(sl, MX, 2.5, CW, 1.0, "Terima Kasih", 44, bold=True, color=WHITE, align=PA.CENTER)
    txt(sl, MX, 3.6, CW, 0.5, "Bidang e-Government — Diskominfostandi Kota Bekasi",
        18, color=TEAL_L, align=PA.CENTER)
    rect(sl, L.SLIDE_W/2-1.0, 4.3, 2.0, Pt(2), GOLD)
    txt(sl, MX, 4.6, CW, 0.5,
        "\"Mewujudkan tata kelola pemerintahan yang cerdas,\nterpadu, dan berkelanjutan\"",
        12, color=TLIGHT, align=PA.CENTER)
    PG[0] += 1

    # ── Save ──
    out = "RESUME_TUPOKSI_egov.pptx"
    prs.save(out)
    print(f"✅ PPT selesai: {out}")
    print(f"   {len(prs.slides)} slide — emoji icons! 🎉")
    return out


if __name__ == "__main__":
    build()
