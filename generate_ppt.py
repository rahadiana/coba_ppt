#!/usr/bin/env python3
"""
generate_ppt.py — RESUME TUPOKSI e-Government
==============================================
2 Versi: Indonesia + English.
Pakai engine src/ppt_engine.py.
Icon pake emoji — no circles/ovals.
"""

import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from pptx.dml.color import RGBColor
from ppt_engine import Engine, MSO_SHAPE, Pt, Inches
from pptx.enum.text import PP_ALIGN as PA

# ─── Colors ──────────────────────────────────────────────────────
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

# ─── Helpers ─────────────────────────────────────────────────────

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

def icon_txt(slide, l, t, s, emoji, bg_color):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, l, t, s, s, bg_color)
    tb = slide.shapes.add_textbox(Inches(l), Inches(t + s*0.08),
                                   Inches(s), Inches(s*0.85))
    tf = tb.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = emoji; p.font.size = Pt(int(s * 36))
    p.font.color.rgb = WHITE; p.alignment = PA.CENTER
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


# ══════════════════════════════════════════════════════════════════
#  CONTENT — Indonesia & English
# ══════════════════════════════════════════════════════════════════

CONTENT = {
    "id": {
        "name": "Indonesia",
        "suffix": "",
        "label": "Bidang e-Government DISKOMINFOSTANDI Kota Bekasi",
        "source": "Sumber: RESUME TUPOKSI egov.docx",
        "motto": "\"Mewujudkan tata kelola pemerintahan yang cerdas,\nterpadu, dan berkelanjutan\"",
        "cover": {
            "title": "RESUME TUPOKSI",
            "sub": "Bidang e-Government",
            "inst": "DISKOMINFOSTANDI Kota Bekasi",
        },
        "pillar": {
            "title": "Tiga Pilar Strategis e-Government",
            "sub": "Bidang e-Government menggerakkan 3 pilar utama transformasi digital Kota Bekasi",
            "items": [
                ("01", "💻", "Pengembangan Aplikasi\n& Sistem Informasi",
                 "Mengawal standar & kualitas\npengembangan aplikasi serta\nmenghubungkan layanan daerah\nmelalui SPLP", DARK),
                ("02", "⚙️", "Tata Kelola SPBE",
                 "Merancang roadmap, arsitektur &\nkebijakan SPBE sekaligus\nmemperkuat kapasitas SDM\ndan peran Government CIO", BLUE),
                ("03", "🌆", "Pengembangan\nKota Cerdas",
                 "Menyusun masterplan, membangun\nkolaborasi lintas sektor, serta\nmengevaluasi program Smart City\nsecara berkelanjutan", TEAL),
            ],
        },
        "systems": {
            "title": "Empat Sistem & Aplikasi Strategis",
            "sub": "Layanan digital yang dioperasikan dan dipelihara oleh Bidang e-Government",
            "items": [
                ("🌐", "Web Pemerintah Kota",  "Portal utama pemerintah kota\nsebagai pusat info & layanan publik", DARK),
                ("📱", "Mobile App PSW",        "Aplikasi mobile Pekan Smart City\n— layanan kota cerdas terintegrasi", BLUE),
                ("🌆", "Web Kota Cerdas",       "Portal informasi & layanan\nprogram Smart City Kota Bekasi", TEAL),
                ("🔗", "Aplikasi SPLP",         "Sistem Penghubung Layanan\nPemerintah — terintegrasi Kemenkomdigi", GOLD),
            ],
        },
        "details": [
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
        ],
        "closing": {
            "thank": "Terima Kasih",
        },
    },
    "en": {
        "name": "English",
        "suffix": "_EN",
        "label": "e-Government Division DISKOMINFOSTANDI Bekasi City",
        "source": "Source: RESUME TUPOKSI egov.docx",
        "motto": "\"Realizing smart, integrated, and\nsustainable government governance\"",
        "cover": {
            "title": "RESUME OF MAIN TASKS\n& FUNCTIONS",
            "sub": "e-Government Division",
            "inst": "DISKOMINFOSTANDI Bekasi City",
        },
        "pillar": {
            "title": "Three Strategic Pillars of e-Government",
            "sub": "The e-Government Division drives 3 main pillars in Bekasi City's digital transformation",
            "items": [
                ("01", "💻", "Application Development\n& Information Systems",
                 "Overseeing standards & quality\nof application development and\nconnecting regional services\nthrough SPLP", DARK),
                ("02", "⚙️", "SPBE Governance",
                 "Designing roadmap, architecture\n& SPBE policies while\nstrengthening HR capacity and\nGovernment CIO roles", BLUE),
                ("03", "🌆", "Smart City\nDevelopment",
                 "Developing masterplan, building\ncross-sector collaboration, and\nevaluating Smart City programs\nsustainably", TEAL),
            ],
        },
        "systems": {
            "title": "Four Strategic Systems & Applications",
            "sub": "Digital services operated and maintained by the e-Government Division",
            "items": [
                ("🌐", "City Government Website", "Main city government portal\nas the center of public\ninformation & services", DARK),
                ("📱", "Mobile App PSW",         "Pekan Smart City mobile app\n— integrated smart city services", BLUE),
                ("🌆", "Smart City Website",     "Information & service portal\nfor Bekasi Smart City program", TEAL),
                ("🔗", "SPLP Application",       "Government Service Gateway\nSystem — integrated with\nKemenkomdigi", GOLD),
            ],
        },
        "details": [
            ("Supervision & Application Development",
             "Scope of work for information system and application development",
             DARK,
             ["✅", "🔗", "🌐", "👥"],
             ["Supervision, analysis, and standardization\nof regional application development",
              "Managing & developing SPLP as the\nbackbone of service integration",
              "Centralized management of regional\ngovernment domains & subdomains",
              "Socialization & capacity building\nfor information system managers"]),
            ("Smart City Program Acceleration",
             "Strategy and collaboration toward an integrated Smart City",
             TEAL,
             ["⭐", "🌆", "💻", "🔗"],
             ["Developing strategies, action plans,\nand Smart City masterplan",
              "Building collaboration with\ncross-sector stakeholders",
              "Coordinating & evaluating Smart City\nprograms on a regular basis",
              "Aligning Smart City masterplan with\nregional planning documents"]),
            ("Strengthening SPBE Governance",
             "Framework toward a mature e-Government system",
             BLUE,
             ["⚙️", "📋", "📝", "📊"],
             ["Developing SPBE strategy, roadmap,\narchitecture, and master plan",
              "Implementing & coordinating SPBE\nprograms across regional agencies",
              "Developing adaptive SPBE policies\nand governance frameworks",
              "Monitoring, evaluation, and continuous\nimprovement recommendations"]),
        ],
        "closing": {
            "thank": "Thank You",
        },
    },
}


# ══════════════════════════════════════════════════════════════════
#  BUILD ENGINE
# ══════════════════════════════════════════════════════════════════

def build_ppt(lang="id"):
    """Generate PPT for given language."""
    c = CONTENT[lang]
    engine = Engine()
    L = engine.L

    from pptx import Presentation
    prs = Presentation()
    prs.slide_width = Inches(L.SLIDE_W)
    prs.slide_height = Inches(L.SLIDE_H)

    MX = L.MARGIN_H; CW = L.cw; PG = [0]

    def ns():
        return prs.slides.add_slide(prs.slide_layouts[6])

    def solid_bg(sl, clr):
        bg = sl.background; bg.fill.solid(); bg.fill.fore_color.rgb = clr

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
        txt(sl, MX, L.SLIDE_H-0.38, 4, 0.22, c["source"], 7, color=TLIGHT)
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
    txt(sl, MX+1.5, 1.6, CW-2, 0.8, c["cover"]["title"], 46, bold=True, color=WHITE)
    txt(sl, MX, 2.9, CW, 0.5, c["cover"]["sub"], 22, color=TEAL_L)
    txt(sl, MX, 3.5, CW, 0.4, c["cover"]["inst"], 14, color=TLIGHT)
    rect(sl, MX, 4.1, 2.5, Pt(3), GOLD)
    txt(sl, MX, 4.4, CW, 0.3, "2025", 11, color=TLIGHT)
    PG[0] += 1

    # ══════════════════════════════════════════════════════════════
    #  SLIDE 2 — PILLARS
    # ══════════════════════════════════════════════════════════════
    sl = content_slide(c["pillar"]["title"], c["pillar"]["sub"])
    pillars = c["pillar"]["items"]

    n = len(pillars)
    cw = L.col_width(n, 0.35)
    gap = 0.35
    sx = (L.SLIDE_W - (n*cw + (n-1)*gap)) / 2
    ch = 4.0; cy = 1.55

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
    #  SLIDE 3 — SYSTEMS
    # ══════════════════════════════════════════════════════════════
    sl = content_slide(c["systems"]["title"], c["systems"]["sub"])
    systems = c["systems"]["items"]

    cw2 = L.col_width(2, 0.4)
    rh2 = 1.8; gh2 = 0.4; gv2 = 0.35
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
    #  SLIDE 4-6 — DETAILS
    # ══════════════════════════════════════════════════════════════
    for si, (title, sub, accent, emojis, items) in enumerate(c["details"]):
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
    txt(sl, MX, 2.5, CW, 1.0, c["closing"]["thank"], 44, bold=True, color=WHITE, align=PA.CENTER)
    txt(sl, MX, 3.6, CW, 0.5, c["label"], 18, color=TEAL_L, align=PA.CENTER)
    rect(sl, L.SLIDE_W/2-1.0, 4.3, 2.0, Pt(2), GOLD)
    txt(sl, MX, 4.6, CW, 0.5, c["motto"], 12, color=TLIGHT, align=PA.CENTER)
    PG[0] += 1

    # ── Save ──
    suffix = c["suffix"]
    out = f"RESUME_TUPOKSI_egov{suffix}.pptx"
    prs.save(out)
    print(f"✅ [{c['name']:>8}] {out}  ({len(prs.slides)} slide)")
    return out


if __name__ == "__main__":
    build_ppt("id")
    build_ppt("en")
    print("\n🎉 2 versi selesai! 🇮🇩🇬🇧")
