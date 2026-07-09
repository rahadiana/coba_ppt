#!/usr/bin/env python3
"""
generate_ppt.py — RESUME TUPOKSI e-Government PPT Generator
=============================================================
Menggunakan engine dari src/ppt_engine.py sebagai foundation.
Tanpa icon, tanpa oval, tanpa decorative circles.
Desain: clean typography-driven, McKinsey-style.
"""

import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from ppt_engine import Engine, MSO_SHAPE, Pt, Inches


def build_presentation():
    """Build PPT menggunakan engine helpers — zero icons/ovals."""
    engine = Engine()
    C = engine.C
    L = engine.L

    # ── Init presentation ──
    from pptx import Presentation
    prs = Presentation()
    prs.slide_width = Inches(L.SLIDE_W)
    prs.slide_height = Inches(L.SLIDE_H)
    engine.prs = prs

    # Constants
    MX = L.MARGIN_H      # 0.6"
    CW = L.cw             # 12.133"
    PG = [0]              # page counter
    BLANK = prs.slide_layouts[6]

    def new_slide():
        return prs.slides.add_slide(BLANK)

    def solid_bg(slide, color):
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = color

    def add_bar(slide, l, t, w, h, color):
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
        s.fill.solid()
        s.fill.fore_color.rgb = color
        s.line.fill.background()
        return s

    def add_box(slide, l, t, w, h, text, size=12, bold=False, color=None,
                align=None, name="Calibri"):
        if color is None:
            color = C.TEXT_D
        if align is None:
            from pptx.enum.text import PP_ALIGN as PA
            align = PA.LEFT
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
        return tb

    def add_rrect(slide, l, t, w, h, fill=None, line=None, radius=0.04):
        if fill is None:
            fill = C.WHITE
        if line is None:
            line = C.ICE
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        sh.line.color.rgb = line
        sh.line.width = Pt(0.5)
        sh.adjustments[0] = radius
        return sh

    def content_header(slide, title, subtitle=None):
        """Standard content slide header: gold bar + navy header."""
        add_bar(slide, 0, 0, L.SLIDE_W, Pt(3), C.GOLD)
        hdr_h = L.HEADER_H - 0.04
        add_bar(slide, 0, Pt(3), L.SLIDE_W, Inches(hdr_h), C.NAVY)
        add_bar(slide, MX, 0.12, 0.07, 0.55, C.GOLD)
        add_box(slide, MX + 0.22, 0.15, CW - 0.5, 0.48,
                title, 20, bold=True, color=C.WHITE)
        if subtitle:
            add_box(slide, MX + 0.22, 0.63, CW - 0.5, 0.28,
                    subtitle, 9, color=C.TEXT_L)

    def content_slide(title, subtitle=None):
        s = new_slide()
        solid_bg(s, C.OFF_W)
        content_header(s, title, subtitle)
        return s

    def render_bullets(slide, items):
        """Clean numbered items with left accent bar."""
        add_bar(slide, MX, L.cy, Pt(4), Inches(3.8), C.NAVY)
        bx = MX + 0.3
        bw = CW - 0.3
        for i, item in enumerate(items):
            y = L.cy + i * 0.55
            add_box(slide, bx, y, 0.35, 0.35, f"{i+1}.", 16, bold=True, color=C.GOLD)
            add_box(slide, bx + 0.4, y + 0.02, bw - 0.4, 0.45, item, 14, color=C.TEXT_D)
            if i < len(items) - 1:
                add_bar(slide, bx + 0.4, y + 0.45, bw - 0.4, Pt(0.5), C.ICE_D)

    def footer(slide):
        add_bar(slide, 0, L.SLIDE_H - L.FOOTER_H, L.SLIDE_W, Pt(2), C.NAVY)
        PG[0] += 1
        add_box(slide, MX, L.SLIDE_H - L.FOOTER_H + 0.06, 4, 0.22,
                "Sumber: RESUME TUPOKSI egov.docx", 7, color=C.TEXT_L)
        add_box(slide, L.SLIDE_W - 1.0, L.SLIDE_H - L.FOOTER_H + 0.06, 0.8, 0.22,
                str(PG[0]), 8, color=C.TEXT_L, align=PA_RIGHT)

    from pptx.enum.text import PP_ALIGN as PA
    global PA_RIGHT
    PA_RIGHT = PA.RIGHT
    PA_CENTER = PA.CENTER

    # ══════════════════════════════════════════════════════════════
    # SLIDE 1 — Cover
    # ══════════════════════════════════════════════════════════════
    s = new_slide()
    solid_bg(s, C.NAVY)
    add_bar(s, 0, L.SLIDE_H - 0.22, L.SLIDE_W, 0.22, C.GOLD)

    add_box(s, MX, 2.0, CW, 1.2, "RESUME TUPOKSI", 48, bold=True, color=C.WHITE)
    add_box(s, MX, 3.4, CW, 0.5, "Bidang e-Government", 22, color=C.TEXT_M)
    add_box(s, MX, 4.0, CW, 0.4, "Diskominfostandi Kota Bekasi", 14, color=C.TEXT_L)
    add_bar(s, MX, 4.6, 2.5, Pt(3), C.GOLD)
    add_box(s, MX, 5.0, CW, 0.3, "2025", 11, color=C.TEXT_L)
    PG[0] += 1

    # ══════════════════════════════════════════════════════════════
    # SLIDE 2 — Tiga Pilar
    # ══════════════════════════════════════════════════════════════
    s = content_slide("Tiga Pilar Utama Tugas Pokok",
                      "Bidang e-Government mencakup 3 area strategis dalam transformasi digital daerah")

    pillars = [
        ("01", "Pengembangan Aplikasi\n& Sistem Informasi",
         "Supervisi, standarisasi, dan\npengelolaan SPLP", C.NAVY),
        ("02", "Tata Kelola e-Government\n(SPBE)",
         "Strategi, roadmap, arsitektur,\ndan kapasitas SDM & GCIO", C.BLUE),
        ("03", "Pengembangan Kota Cerdas\n(Smart City)",
         "Masterplan, kolaborasi lintas\nsektor, dan evaluasi program", C.TEAL),
    ]

    n = len(pillars)
    card_w = L.col_width(n, 0.4)
    gap = 0.4
    start_x = (L.SLIDE_W - (n * card_w + (n - 1) * gap)) / 2
    card_h = 3.8

    for i, (num, title, desc, clr) in enumerate(pillars):
        cx = start_x + i * (card_w + gap)
        add_rrect(s, cx, L.cy + 0.3, card_w, card_h)
        add_bar(s, cx, L.cy + 0.3, card_w, Pt(5), clr)
        add_box(s, cx + 0.2, L.cy + 0.65, 0.5, 0.35, num, 16, bold=True, color=clr)
        add_box(s, cx + 0.2, L.cy + 1.1, card_w - 0.4, 0.9,
                title, 15, bold=True, color=C.NAVY)
        add_bar(s, cx + 0.2, L.cy + 2.1, card_w - 0.4, Pt(1), C.ICE)
        add_box(s, cx + 0.2, L.cy + 2.3, card_w - 0.4, 1.2,
                desc, 12, color=C.TEXT_M)

    footer(s)

    # ══════════════════════════════════════════════════════════════
    # SLIDE 3 — Sistem & Aplikasi (2x2 grid cards)
    # ══════════════════════════════════════════════════════════════
    s = content_slide("Empat Sistem & Aplikasi Strategis yang Dikelola",
                      "Layanan digital yang dioperasikan dan dipelihara oleh Bidang e-Government")

    systems = [
        ("Web Pemerintah Kota", "Admin Teknis Web Pemkot & Web Diskominfo", C.NAVY),
        ("Mobile App PSW", "Platform Smart City — Pekan Smart City", C.BLUE),
        ("Web Kota Cerdas", "Admin Teknis Portal Kota Cerdas", C.TEAL),
        ("Aplikasi SPLP", "Sistem Penghubung Layanan Pemerintah (Kemenkomdigi)", C.GOLD),
    ]

    cw2 = L.col_width(2, 0.5)
    rh2 = 1.6
    gh2 = 0.5
    gv2 = 0.4
    grid_sx = (L.SLIDE_W - (2 * cw2 + gh2)) / 2

    for i, (title, desc, clr) in enumerate(systems):
        col = i % 2
        row = i // 2
        cx = grid_sx + col * (cw2 + gh2)
        cy = L.cy + row * (rh2 + gv2)

        add_rrect(s, cx, cy, cw2, rh2)
        add_bar(s, cx, cy, Pt(5), rh2, clr)
        add_box(s, cx + 0.25, cy + 0.25, cw2 - 0.5, 0.4,
                title, 17, bold=True, color=clr)
        add_box(s, cx + 0.25, cy + 0.75, cw2 - 0.5, 0.6,
                desc, 13, color=C.TEXT_D)

    footer(s)

    # ══════════════════════════════════════════════════════════════
    # SLIDE 4-6 — Detail slides with bullets
    # ══════════════════════════════════════════════════════════════

    details = [
        ("Supervisi & Pengembangan Aplikasi Daerah",
         "Lingkup kerja pengembangan sistem informasi dan aplikasi perangkat daerah", [
             "Supervisi, analisis, dan standarisasi pengembangan aplikasi perangkat daerah",
             "Mengelola dan mengembangkan Sistem Penghubung Layanan Pemerintah (SPLP)",
             "Pengelolaan domain dan subdomain pemerintah daerah",
             "Sosialisasi dan peningkatan kapasitas SDM pengelola sistem",
         ]),
        ("Akselerasi Program Kota Cerdas",
         "Strategi dan kolaborasi menuju Smart City yang terintegrasi", [
             "Menyusun strategi, rencana aksi, dan masterplan Kota Cerdas",
             "Membangun kolaborasi dengan pemangku kepentingan lintas sektor",
             "Mengoordinasikan dan mengevaluasi program Kota Cerdas",
             "Menyelaraskan rencana induk dengan dokumen perencanaan daerah",
         ]),
        ("Penguatan Tata Kelola SPBE",
         "Kerangka kerja menuju Sistem Pemerintahan Berbasis Elektronik yang matang", [
             "Menyusun strategi, roadmap, arsitektur, dan peta rencana SPBE",
             "Melaksanakan dan mengoordinasikan program SPBE lintas perangkat daerah",
             "Mengembangkan kebijakan dan tata kelola SPBE",
             "Monitoring, evaluasi, dan rekomendasi perbaikan berkelanjutan",
         ]),
    ]

    for title, subtitle, items in details:
        s = content_slide(title, subtitle)
        render_bullets(s, items)
        footer(s)

    # ══════════════════════════════════════════════════════════════
    # SLIDE 7 — Closing
    # ══════════════════════════════════════════════════════════════
    s = new_slide()
    solid_bg(s, C.NAVY)
    add_bar(s, 0, L.SLIDE_H - 0.22, L.SLIDE_W, 0.22, C.GOLD)

    add_box(s, MX, 2.5, CW, 1.0, "Terima Kasih", 44, bold=True, color=C.WHITE, align=PA_CENTER)
    add_box(s, MX, 3.6, CW, 0.6, "Bidang e-Government — Diskominfostandi Kota Bekasi",
            18, color=C.TEXT_M, align=PA_CENTER)
    add_box(s, MX, 4.4, CW, 0.5,
            "\"Mewujudkan tata kelola pemerintahan yang cerdas, terpadu, dan berkelanjutan\"",
            12, color=C.TEXT_L, align=PA_CENTER)
    PG[0] += 1

    # ── Save ──
    output = "RESUME_TUPOKSI_egov.pptx"
    prs.save(output)
    print(f"✅ PPT berhasil: {output}")
    print(f"   {len(prs.slides)} slide — clean design, no icons/ovals")
    return output


if __name__ == "__main__":
    build_presentation()
