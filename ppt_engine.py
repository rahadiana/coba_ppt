#!/usr/bin/env python3
"""
ppt_engine.py — Reusable Presentation Engine
=============================================
LayoutFrame class + draw helpers + generic archetype functions.
Terpisah dari content — bisa dipakai untuk PPT apa saja.

Arsitektur:
    ppt_engine.py       ← engine reusable (file ini)
    content_*.py        ← data konten terpisah
    generate_ppt.py     ← entry point

Usage:
    from ppt_engine import Engine
    engine = Engine()
    prs = engine.build(slides_data)
    prs.save("output.pptx")
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import math, os


# ═══════════════════════════════════════════════════════════
# COLOR UTILITY
# ═══════════════════════════════════════════════════════════

class Colors:
    """Palette 60-30-10 default. Bisa di- extend atau override."""
    # 60% Dominant
    NAVY   = RGBColor(0x0A, 0x16, 0x28)
    NAVY_L = RGBColor(0x12, 0x29, 0x4A)
    NAVY_M = RGBColor(0x1B, 0x3A, 0x6B)
    NAVY_D = RGBColor(0x0D, 0x1F, 0x3C)
    # 30% Secondary
    WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    OFF_W  = RGBColor(0xF5, 0xF7, 0xFA)
    ICE    = RGBColor(0xE8, 0xED, 0xF5)
    ICE_D  = RGBColor(0xD0, 0xD8, 0xE8)
    # 10% Accent
    GOLD   = RGBColor(0xC8, 0x96, 0x2E)
    GOLD_L = RGBColor(0xD4, 0xA0, 0x17)
    # Semantic
    TEXT_D = RGBColor(0x1A, 0x1A, 0x2E)
    TEXT_M = RGBColor(0x6B, 0x72, 0x88)
    TEXT_L = RGBColor(0x9C, 0xA3, 0xAF)
    BLUE   = RGBColor(0x25, 0x63, 0xEB)
    TEAL   = RGBColor(0x0D, 0x94, 0x88)
    WARM   = RGBColor(0xB8, 0x86, 0x0B)
    RED    = RGBColor(0xDC, 0x26, 0x26)

    @staticmethod
    def from_hex(hex_str):
        """Convert '#2563EB' or '2563EB' to RGBColor."""
        h = hex_str.lstrip('#')
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    @staticmethod
    def named(name, fallback=None):
        """Cari warna by name string, fallback ke NAVY."""
        if fallback is None:
            fallback = Colors.NAVY
        return getattr(Colors, name.upper(), fallback)


# ═══════════════════════════════════════════════════════════
# LAYOUT FRAME — Semua kalkulasi layout
# ═══════════════════════════════════════════════════════════

class LayoutFrame:
    """
    Kalkulator posisi & ukuran untuk elemen slide.
    
    Slide widescreen 16:9 = 13.333" × 7.5"
    
    Zona layout:
        MARGIN_H  = 0.6"      (kiri/kanan)
        HEADER_H  = 0.94"     (gold_bar 0.04" + navy 0.9")
        HEADER_GAP = 0.21"    (jarak header → konten)
        FOOTER_H  = 0.50"     (navy bar 0.03" + text 0.22" + gap)
    
    Content area:
        cx = 0.6",  cy = 1.15"
        cw = 12.133",  ch = 5.85"
    """
    
    SLIDE_W = 13.333
    SLIDE_H = 7.5
    
    MARGIN_H   = 0.6
    HEADER_H   = 0.94
    HEADER_GAP = 0.21
    FOOTER_H   = 0.50
    
    GAP_H      = 0.30
    GAP_V      = 0.30
    GAP_ITEM   = 0.02
    CARD_PAD   = 0.15
    
    # Characters-per-inch (CPI) untuk Calibri
    CPI = { 7:17, 8:15, 9:14, 10:13, 11:12, 12:11,
            13:10, 14:9, 16:8, 18:7, 20:6.5, 24:5.5,
            32:4, 40:3, 48:2.5 }
    
    def __init__(self):
        self.cx = self.MARGIN_H
        self.cw = self.SLIDE_W - 2 * self.MARGIN_H
        self.cy = self.HEADER_H + self.HEADER_GAP
        self.ch = (self.SLIDE_H - self.FOOTER_H) - self.cy
    
    # ─── Grid ───
    
    def col_width(self, n, gap=None):
        """col_w = (cw - (n-1) × gap) / n"""
        g = gap if gap is not None else self.GAP_H
        return (self.cw - (n - 1) * g) / n
    
    def row_height(self, n, gap=None):
        """row_h = (ch - (n-1) × gap_v) / n"""
        g = gap if gap is not None else self.GAP_V
        return (self.ch - (n - 1) * g) / n
    
    def grid_pos(self, col, row, col_w, row_h, gap_h=None, gap_v=None):
        """x = cx + col × (col_w + gap_h)"""
        gh = gap_h if gap_h is not None else self.GAP_H
        gv = gap_v if gap_v is not None else self.GAP_V
        return self.cx + col * (col_w + gh), self.cy + row * (row_h + gv)
    
    def calc_grid(self, n_items, cols=0):
        """Kalkulasi grid multi-baris."""
        per_row = cols if cols > 0 else (
            2 if n_items <= 2 else (3 if n_items <= 3 else 4))
        n_rows = (n_items + per_row - 1) // per_row
        col_w = self.col_width(per_row)
        row_h = self.row_height(n_rows)
        return {'per_row': per_row, 'n_rows': n_rows,
                'col_w': col_w, 'row_h': row_h,
                'gap_h': self.GAP_H, 'gap_v': self.GAP_V}
    
    def grid_cell(self, i, grid):
        """Posisi cell ke-i dalam grid."""
        col = i % grid['per_row']
        row = i // grid['per_row']
        x, y = self.grid_pos(col, row, grid['col_w'], grid['row_h'],
                              grid['gap_h'], grid['gap_v'])
        return x, y, grid['col_w'], grid['row_h']
    
    # ─── Card Internal ───
    
    def card_item_layout(self, card_h, has_icon=True, n_items=1, title_h=0.30):
        icon_size = 0.42
        icon_y = self.CARD_PAD
        if has_icon:
            title_y = icon_y + icon_size + 0.13
        else:
            title_y = self.CARD_PAD
        item_h = 0.35
        item_start_y = title_y + title_h + 0.08
        content_used = item_start_y + n_items * (item_h + self.GAP_ITEM)
        overflow = max(0, content_used - card_h)
        return {'icon_size': icon_size, 'icon_y': icon_y,
                'title_y': title_y, 'item_start_y': item_start_y,
                'item_h': item_h, 'overflow': overflow}
    
    # ─── Estimasi Teks ───
    
    def text_height(self, text, font_size, box_width):
        """Tinggi teks estimated: ceil(len/(cpi*box_w)) * pt*1.2/72"""
        cpi = self.CPI.get(font_size, 12)
        max_ch = box_width * cpi
        if max_ch <= 0:
            return font_size * 1.2 / 72
        n_lines = math.ceil(len(text) / max_ch)
        return n_lines * font_size * 1.2 / 72 + 0.02
    
    def safe_item_height(self, text, font_size, box_width, min_h=0.25):
        return max(self.text_height(text, font_size, box_width) + 0.05, min_h)
    
    # ─── Verifikasi ───
    
    def check_bounds(self, x, y, w, h):
        warns = []
        if x < 0: warns.append(f"x={x:.2f}<0")
        if y < 0: warns.append(f"y={y:.2f}<0")
        if x + w > self.SLIDE_W: warns.append(f"right={x+w:.2f}>{self.SLIDE_W}")
        if y + h > self.SLIDE_H: warns.append(f"bottom={y+h:.2f}>{self.SLIDE_H}")
        return (len(warns) == 0, warns)
    
    def check_overlap(self, a, b):
        ax, ay, aw, ah = a
        bx, by, bw, bh = b
        ox = max(0, min(ax+aw, bx+bw) - max(ax, bx))
        oy = max(0, min(ay+ah, by+bh) - max(ay, by))
        overlap = ox * oy
        min_area = min(aw * ah, bw * bh)
        return 0 if min_area == 0 else (overlap / min_area) * 100


# ═══════════════════════════════════════════════════════════
# ENGINE — Menggabungkan LayoutFrame + Draw Helpers + Archetypes
# ═══════════════════════════════════════════════════════════

class Engine:
    """
    Presentation Engine — generic, content-agnostic.
    
    Cara pakai:
        engine = Engine()
        prs = engine.build(slides_data)   # slides_data = list of slide dicts
        prs.save("output.pptx")
    
    Slide dict format:
        {
            "type": "cover|toc|section|content|card_grid|two_col|callout|flow|table|closing",
            "data": { ... key/value sesuai type ... }
        }
    """
    
    def __init__(self, colors=None):
        self.L = LayoutFrame()
        self.C = colors or Colors()
        self.prs = None
        self.pg_counter = [0]
        self.source_text = ""
    
    # ════════════════════════════════════════════════════════
    # BUILD — main entry point
    # ════════════════════════════════════════════════════════
    
    def build(self, slides_data, source_text="Sumber: ...", output_path=None):
        """
        Generate presentation dari array slide dicts.
        
        Args:
            slides_data: list of dict — lihat STRUCTURE.md atau contoh
            source_text: teks footer sumber
            output_path: path output (None = return Presentation object)
        
        Returns:
            Presentation object (panggil .save() untuk simpan)
        """
        self.source_text = source_text
        self.pg_counter = [0]
        
        self.prs = Presentation()
        self.prs.slide_width = Inches(self.L.SLIDE_W)
        self.prs.slide_height = Inches(self.L.SLIDE_H)
        
        dispatcher = {
            "cover":     self._build_cover,
            "toc":       self._build_toc,
            "section":   self._build_section,
            "content":   self._build_content,
            "card_grid": self._build_card_grid,
            "two_col":   self._build_two_col,
            "callout":   self._build_callout,
            "flow":        self._build_flow,
            "table":       self._build_table,
            "nsr_factors": self._build_nsr_factors,
            "closing":     self._build_closing,
        }
        
        for slide_def in slides_data:
            stype = slide_def.get("type", "content")
            data = slide_def.get("data", {})
            builder = dispatcher.get(stype)
            if builder:
                builder(data)
            else:
                # fallback: content slide
                self._build_content(data)
        
        if output_path:
            self.prs.save(output_path)
        
        return self.prs
    
    # ─── Helpers ───
    
    def _blank(self):
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])
    
    def _resolve_color(self, clr):
        """clr bisa string hex '#2563EB', string name 'BLUE', atau RGBColor."""
        if isinstance(clr, RGBColor):
            return clr
        if isinstance(clr, str):
            if clr.startswith('#'):
                return Colors.from_hex(clr)
            return Colors.named(clr)
        return self.C.NAVY
    
    def _add_box(self, slide, left, top, width, height, text,
                 size=12, bold=False, color=None, align=PP_ALIGN.LEFT,
                 font="Calibri", vAlign=MSO_ANCHOR.TOP):
        if color is None:
            color = self.C.TEXT_D
        else:
            color = self._resolve_color(color)
        tb = slide.shapes.add_textbox(Inches(left), Inches(top),
                                       Inches(width), Inches(height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.font.name = font
        p.alignment = align
        p.space_after = Pt(0)
        p.space_before = Pt(0)
        self.L.check_bounds(left, top, width, height)
        return tb
    
    def _add_shape(self, slide, stype, left, top, width, height,
                   fill=None, line=None, lw=None):
        sh = slide.shapes.add_shape(stype, Inches(left), Inches(top),
                                     Inches(width), Inches(height))
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = self._resolve_color(fill)
        else:
            sh.fill.background()
        if line:
            sh.line.color.rgb = self._resolve_color(line)
            if lw: sh.line.width = Pt(lw)
        else:
            sh.line.fill.background()
        self.L.check_bounds(left, top, width, height)
        return sh
    
    def _add_rrect(self, slide, left, top, width, height,
                   fill=None, line=None, lw=0.5, radius=0.04):
        if fill is None: fill = self.C.WHITE
        if line is None: line = self.C.ICE
        sh = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top), Inches(width), Inches(height))
        sh.fill.solid()
        sh.fill.fore_color.rgb = self._resolve_color(fill)
        if line:
            sh.line.color.rgb = self._resolve_color(line)
            sh.line.width = Pt(lw)
        else:
            sh.line.fill.background()
        sh.adjustments[0] = radius
        self.L.check_bounds(left, top, width, height)
        return sh
    
    def _add_oval(self, slide, left, top, size, fill=None):
        if fill is None: fill = self.C.BLUE
        sh = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(size), Inches(size))
        sh.fill.solid()
        sh.fill.fore_color.rgb = self._resolve_color(fill)
        sh.line.fill.background()
        return sh
    
    def _text_to_shape(self, shape, text, size=12, bold=False, color=None,
                       align=PP_ALIGN.LEFT, font="Calibri"):
        if color is None: color = self.C.TEXT_D
        else: color = self._resolve_color(color)
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.font.name = font
        p.alignment = align
        return tf
    
    # ─── Header & Footer ───
    
    def _gold_top_bar(self, slide):
        self._add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0,
                        self.L.SLIDE_W, 0.035, fill=self.C.GOLD)
    
    def _navy_header(self, slide, title, subtitle=None):
        self._gold_top_bar(slide)
        hdr_h = self.L.HEADER_H - 0.04
        self._add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0.035,
                        self.L.SLIDE_W, hdr_h, fill=self.C.NAVY)
        self._add_shape(slide, MSO_SHAPE.RECTANGLE,
                        self.L.MARGIN_H, 0.12, 0.07, 0.55, fill=self.C.GOLD)
        self._add_box(slide, self.L.MARGIN_H + 0.22, 0.15,
                      self.L.cw - 0.5, 0.48, title, 20, bold=True, color=self.C.WHITE)
        if subtitle:
            self._add_box(slide, self.L.MARGIN_H + 0.22, 0.63,
                          self.L.cw - 0.5, 0.28, subtitle, 9, color=self.C.TEXT_L)
    
    def _footer(self, slide):
        h = 0.03
        self._add_shape(slide, MSO_SHAPE.RECTANGLE, 0,
                        self.L.SLIDE_H - self.L.FOOTER_H,
                        self.L.SLIDE_W, h, fill=self.C.NAVY)
        self._add_box(slide, self.L.MARGIN_H,
                      self.L.SLIDE_H - self.L.FOOTER_H + 0.06,
                      4, 0.22, self.source_text, 7, color=self.C.TEXT_L)
        self.pg_counter[0] += 1
        self._add_box(slide, self.L.SLIDE_W - 1.0,
                      self.L.SLIDE_H - self.L.FOOTER_H + 0.06,
                      0.8, 0.22, str(self.pg_counter[0]), 8,
                      color=self.C.TEXT_L, align=PP_ALIGN.RIGHT)
    
    def _content_slide(self, title, subtitle=None):
        """Bikin blank slide + navy header."""
        s = self._blank()
        bg = s.background; bg.fill.solid(); bg.fill.fore_color.rgb = self.C.OFF_W
        self._navy_header(s, title, subtitle)
        return s
    
    # ════════════════════════════════════════════════════════
    # ARCHETYPE BUILDERS — generic, content-driven
    # ════════════════════════════════════════════════════════
    
    def _build_cover(self, data):
        """
        Slide cover — full-bleed navy + dekoratif ovals.
        
        Data:
            pre_title, city, main_title, main_subtitle,
            display_title, badge_text, badge_subtext
        """
        s = self._blank()
        bg = s.background; bg.fill.solid(); bg.fill.fore_color.rgb = self.C.NAVY
        
        # Dekoratif ovals
        self._add_oval(s, -1, -1.5, 4.5, self.C.NAVY_D)
        self._add_oval(s, 8.5, -2, 7, self.C.NAVY_D)
        self._add_oval(s, 10, 4.5, 5, self.C.NAVY_D)
        self._add_oval(s, 0.5, 5.5, 2.5, self.C.NAVY_L)
        
        mx = self.L.MARGIN_H
        cw = self.L.cw
        
        # Gold bar
        self._add_shape(s, MSO_SHAPE.RECTANGLE, mx, 2.6, 4, 0.04, fill=self.C.GOLD)
        
        # Teks
        pre = data.get("pre_title", "")
        if pre:
            self._add_box(s, mx, 0.5, 5, 0.35, pre, 12, bold=True, color=self.C.GOLD)
        
        city = data.get("city", "")
        if city:
            self._add_box(s, mx, 0.85, 5, 0.35, city, 14, bold=True, color=self.C.WHITE)
        
        main_title = data.get("main_title", "")
        if main_title:
            self._add_box(s, mx, 1.8, cw, 0.45, main_title, 20, bold=True, color=self.C.WHITE)
        
        main_sub = data.get("main_subtitle", "")
        if main_sub:
            self._add_box(s, mx, 2.15, cw, 0.4, main_sub, 15, bold=True, color=self.C.GOLD)
        
        display = data.get("display_title", "")
        if display:
            self._add_box(s, mx, 3.1, cw, 2.0, display, 40, bold=True, color=self.C.WHITE)
        
        # Badge
        badge_text = data.get("badge_text", "")
        if badge_text:
            self._add_rrect(s, mx, 5.4, 7.5, 0.9, fill=self.C.NAVY_D)
            self._add_box(s, mx + 0.2, 5.45, 7, 0.4, badge_text, 12, color=self.C.ICE)
        
        badge_sub = data.get("badge_subtext", "")
        if badge_sub:
            self._add_box(s, mx + 0.2, 5.75, 7, 0.35, badge_sub, 10, color=self.C.TEXT_L)
        
        # Bottom gold bar
        self._add_shape(s, MSO_SHAPE.RECTANGLE, 0,
                        self.L.SLIDE_H - 0.22, self.L.SLIDE_W, 0.22, fill=self.C.GOLD)
    
    def _build_toc(self, data):
        """
        Daftar isi — tabel konten dengan badge.
        
        Data:
            title, subtitle,
            items: [{num, label, color}]
            cols: 2 (default)
        """
        title = data.get("title", "Daftar Isi")
        subtitle = data.get("subtitle")
        items = data.get("items", [])
        cols = data.get("cols", 2)
        
        s = self._content_slide(title, subtitle)
        
        if not items:
            self._footer(s)
            return
        
        cw_col = self.L.col_width(cols, 0.35)
        n_per_col = (len(items) + cols - 1) // cols
        row_h = min(0.62, (self.L.ch - 0.2) / n_per_col)
        
        for i, item in enumerate(items):
            col = i // n_per_col
            row = i % n_per_col
            x = self.L.cx + col * (cw_col + 0.35)
            y = self.L.cy + row * row_h
            
            clr = self._resolve_color(item.get("color", self.C.BLUE))
            num = str(item.get("num", ""))
            lb = item.get("label", "")
            
            self._add_rrect(s, x, y, cw_col, row_h - 0.05)
            badge = self._add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE,
                                     x + 0.1, y + 0.06, 0.5, 0.5, fill=clr)
            self._text_to_shape(badge, num, 14, bold=True,
                                color=self.C.WHITE, align=PP_ALIGN.CENTER)
            self._add_box(s, x + 0.75, y + 0.06, cw_col - 0.9, 0.5,
                          lb, 13, color=self.C.TEXT_D,
                          vAlign=MSO_ANCHOR.MIDDLE)
        
        self._footer(s)
    
    def _build_section(self, data):
        """
        Section divider — full-bleed dark.
        
        Data:
            title, subtitle, action_text
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        action_text = data.get("action_text")
        
        s = self._blank()
        bg = s.background; bg.fill.solid(); bg.fill.fore_color.rgb = self.C.NAVY
        self._add_oval(s, -1.5, -1.5, 5, self.C.NAVY_L)
        self._add_oval(s, -0.5, -0.5, 3.5, self.C.NAVY_D)
        self._add_oval(s, 9.5, 4, 5, self.C.NAVY_D)
        self._add_oval(s, 10.5, 3, 3, self.C.NAVY_L)
        self._add_shape(s, MSO_SHAPE.RECTANGLE,
                        self.L.MARGIN_H, 2.3, 2.5, 0.04, fill=self.C.GOLD)
        self._add_box(s, self.L.MARGIN_H, 2.6, self.L.cw, 1.6,
                      title, 34, bold=True, color=self.C.WHITE)
        
        y_sub = 4.2
        if action_text:
            self._add_box(s, self.L.MARGIN_H, y_sub, self.L.cw, 0.4,
                          action_text, 14, bold=True, color=self.C.GOLD)
            y_sub += 0.35
        if subtitle:
            self._add_box(s, self.L.MARGIN_H, y_sub, self.L.cw, 0.3,
                          subtitle, 11, color=self.C.TEXT_L)
        
        self._add_shape(s, MSO_SHAPE.RECTANGLE, 0,
                        self.L.SLIDE_H - 0.22, self.L.SLIDE_W, 0.22, fill=self.C.GOLD)
        self.pg_counter[0] += 1
    
    def _build_content(self, data):
        """
        Content slide polos — header + footer doang.
        
        Data:
            title, subtitle
        """
        s = self._content_slide(data.get("title", ""), data.get("subtitle"))
        self._footer(s)
    
    def _build_card_grid(self, data):
        """
        Grid of cards — multi-baris, icon, bullet items.
        
        Data:
            title, subtitle,
            cards: [{icon, title, color, items: [str]}]
            cols: 0 (auto)
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        cards = data.get("cards", [])
        cols = data.get("cols", 0)
        
        s = self._content_slide(title, subtitle)
        n = len(cards)
        if n == 0:
            self._footer(s)
            return
        
        grid = self.L.calc_grid(n, cols)
        
        for i, cd in enumerate(cards):
            cx, cy, cw, ch = self.L.grid_cell(i, grid)
            
            clr_s = cd.get('color', self.C.BLUE)
            clr = self._resolve_color(clr_s)
            ic = cd.get('icon', '')
            t = cd.get('title', '')
            items = cd.get('items', [])
            
            self._add_rrect(s, cx, cy, cw, ch)
            self._add_shape(s, MSO_SHAPE.RECTANGLE, cx, cy, 0.05, ch, fill=clr)
            
            item_layout = self.L.card_item_layout(ch, has_icon=bool(ic), n_items=len(items))
            
            if ic:
                self._add_oval(s, cx + self.L.CARD_PAD, cy + item_layout['icon_y'],
                               item_layout['icon_size'], clr)
                self._add_box(s, cx + self.L.CARD_PAD, cy + item_layout['icon_y'],
                              item_layout['icon_size'], item_layout['icon_size'],
                              ic, 13, color=self.C.WHITE, align=PP_ALIGN.CENTER)
                title_y = cy + item_layout['title_y']
            else:
                title_y = cy + self.L.CARD_PAD
            
            self._add_box(s, cx + self.L.CARD_PAD, title_y, cw - self.L.CARD_PAD * 2, 0.30,
                          t, 13, bold=True, color=clr)
            
            ay = cy + item_layout['item_start_y']
            box_w = cw - self.L.CARD_PAD * 2
            for item in items:
                item_h = self.L.safe_item_height(f"• {item}", 9, box_w, min_h=0.30)
                self._add_box(s, cx + self.L.CARD_PAD, ay, box_w, item_h,
                              f"• {item}", 9, color=self.C.TEXT_D)
                ay += item_h + self.L.GAP_ITEM
        
        self._footer(s)
    
    def _build_two_col(self, data):
        """
        Two column layout — 2 card side by side.
        
        Data:
            title, subtitle,
            left: {color, lines: [str]},    # $ prefix = highlight
            right: {color, lines: [str]}
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        left_data = data.get("left", {})
        right_data = data.get("right", {})
        
        s = self._content_slide(title, subtitle)
        gap = self.L.GAP_H
        cw = self.L.col_width(2, gap)
        sy = self.L.cy
        ch = self.L.ch
        
        for data_block, clr_s, x in [
            (left_data.get('lines', []), left_data.get('color', self.C.BLUE), self.L.cx),
            (right_data.get('lines', []), right_data.get('color', self.C.TEAL),
             self.L.cx + cw + gap),
        ]:
            clr = self._resolve_color(clr_s)
            self._add_rrect(s, x, sy, cw, ch)
            self._add_shape(s, MSO_SHAPE.RECTANGLE, x, sy, 0.05, ch, fill=clr)
            ay = sy + 0.15
            for line in data_block:
                if line.startswith("$"):
                    self._add_box(s, x + 0.2, ay, cw - 0.35, 0.30,
                                  line[1:], 14, bold=True, color=clr)
                    ay += 0.32
                else:
                    self._add_box(s, x + 0.2, ay, cw - 0.35, 0.25,
                                  f"• {line}", 11, color=self.C.TEXT_D)
                    ay += 0.26
        
        self._footer(s)
    
    def _build_callout(self, data):
        """
        Big number callout cards.
        
        Data:
            title, subtitle,
            callouts: [{number, label, color}]
            note: str (optional)
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        callouts = data.get("callouts", [])
        note_text = data.get("note")
        
        s = self._content_slide(title, subtitle)
        n = len(callouts)
        if n == 0:
            self._footer(s)
            return
        
        cw = self.L.col_width(n)
        sy = self.L.cy
        
        for i, co in enumerate(callouts):
            cx = self.L.cx + i * (cw + self.L.GAP_H)
            clr = self._resolve_color(co.get("color", self.C.BLUE))
            num = str(co.get("number", ""))
            lb = co.get("label", "")
            
            self._add_rrect(s, cx, sy, cw, 2.2)
            self._add_shape(s, MSO_SHAPE.RECTANGLE, cx, sy, cw, 0.05, fill=clr)
            self._add_box(s, cx, sy + 0.25, cw, 0.7, num, 32, bold=True,
                          color=clr, align=PP_ALIGN.CENTER)
            self._add_box(s, cx + 0.1, sy + 1.0, cw - 0.2, 0.7, lb, 11,
                          color=self.C.TEXT_M, align=PP_ALIGN.CENTER,
                          vAlign=MSO_ANCHOR.TOP)
        
        if note_text:
            ny = sy + 2.5
            self._add_rrect(s, self.L.cx, ny, self.L.cw, 2.8, fill=self.C.ICE)
            self._add_box(s, self.L.cx + 0.25, ny + 0.15, self.L.cw - 0.5, 2.5,
                          note_text, 11, color=self.C.TEXT_D)
        
        self._footer(s)
    
    def _build_flow(self, data):
        """
        Horizontal flow dengan step dan arrow.
        
        Data:
            title, subtitle,
            steps: [{num, title, desc, color}]
            note: str (optional)
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        steps = data.get("steps", [])
        note_text = data.get("note")
        
        s = self._content_slide(title, subtitle)
        n = len(steps)
        if n == 0:
            self._footer(s)
            return
        
        bw = self.L.col_width(n)
        bgap = self.L.GAP_H
        sy = self.L.cy + 0.15
        
        for i, step in enumerate(steps):
            cx = self.L.cx + i * (bw + bgap)
            clr = self._resolve_color(step.get("color", self.C.BLUE))
            num = str(step.get("num", ""))
            step_title = step.get("title", "")
            desc = step.get("desc", "")
            
            self._add_rrect(s, cx, sy, bw, 2.0)
            self._add_shape(s, MSO_SHAPE.RECTANGLE, cx, sy, bw, 0.05, fill=clr)
            circ_x = cx + bw / 2 - 0.25
            self._add_oval(s, circ_x, sy + 0.15, 0.5, clr)
            self._add_box(s, circ_x, sy + 0.15, 0.5, 0.5, num, 18, bold=True,
                          color=self.C.WHITE, align=PP_ALIGN.CENTER)
            self._add_box(s, cx + 0.1, sy + 0.75, bw - 0.2, 0.35, step_title,
                          14, bold=True, color=clr, align=PP_ALIGN.CENTER)
            self._add_box(s, cx + 0.1, sy + 1.1, bw - 0.2, 0.7, desc, 10,
                          color=self.C.TEXT_M, align=PP_ALIGN.CENTER,
                          vAlign=MSO_ANCHOR.TOP)
            if i < n - 1:
                self._add_box(s, cx + bw, sy + 0.7, bgap, 0.5, "›", 24,
                              color=self.C.GOLD, align=PP_ALIGN.CENTER,
                              vAlign=MSO_ANCHOR.MIDDLE)
        
        if note_text:
            ny = sy + 2.4
            self._add_rrect(s, self.L.cx, ny, self.L.cw, 2.8, fill=self.C.ICE)
            self._add_box(s, self.L.cx + 0.25, ny + 0.15, self.L.cw - 0.5, 2.5,
                          note_text, 11, color=self.C.TEXT_D)
        
        self._footer(s)
    
    def _set_cell(self, cell, text, size=11, bold=False, color=None,
                  align=PP_ALIGN.LEFT, fill=None):
        if color is None: color = self.C.TEXT_D
        else: color = self._resolve_color(color)
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        p.text = str(text)
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.alignment = align
        cell.text_frame.word_wrap = True
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill:
            cell.fill.solid()
            cell.fill.fore_color.rgb = self._resolve_color(fill)
    
    def _build_table(self, data):
        """
        Native table slide.
        
        Data:
            title, subtitle,
            headers: [str],
            rows: [[str], ...]
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        headers = data.get("headers", [])
        rows = data.get("rows", [])
        
        s = self._content_slide(title, subtitle)
        
        if headers and rows:
            nc = len(headers)
            nr = len(rows) + 1
            rh = 0.42
            ts = s.shapes.add_table(nr, nc, Inches(self.L.cx), Inches(self.L.cy),
                                     Inches(self.L.cw), Inches(rh * nr))
            tbl = ts.table
            for i in range(nc):
                tbl.columns[i].width = int(Inches(self.L.cw / nc))
            for j, hdr in enumerate(headers):
                self._set_cell(tbl.cell(0, j), hdr, bold=True, color=self.C.WHITE,
                               align=PP_ALIGN.CENTER, fill=self.C.NAVY)
            for ri, row in enumerate(rows):
                bg_c = self.C.ICE if ri % 2 == 0 else self.C.WHITE
                for j, val in enumerate(row):
                    al = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
                    self._set_cell(tbl.cell(ri + 1, j), val, color=self.C.TEXT_D,
                                   align=al, fill=bg_c)
        
        self._footer(s)
    
    def _build_nsr_factors(self, data):
        """
        NSR Factors — 7 faktor dengan oval numbers + classification box.
        
        Data:
            title, subtitle,
            factors: [{num, label, color}],
            classification_note: str
        """
        title = data.get("title", "")
        subtitle = data.get("subtitle")
        factors = data.get("factors", [])
        class_note = data.get("classification_note", "")
        
        s = self._content_slide(title, subtitle)
        n = len(factors)
        if n == 0:
            self._footer(s)
            return
        
        per_row = 4
        fw = self.L.col_width(per_row, 0.3)
        
        for i, f in enumerate(factors):
            col = i % per_row
            row = i // per_row
            x, y = self.L.grid_pos(col, row, fw, 1.5, gap_v=0.3)
            clr = self._resolve_color(f.get("color", self.C.BLUE))
            num = str(f.get("num", ""))
            lb = f.get("label", "")
            
            self._add_rrect(s, x, y, fw, 1.3)
            self._add_oval(s, x + 0.15, y + 0.25, 0.55, clr)
            self._add_box(s, x + 0.15, y + 0.25, 0.55, 0.55, num,
                          18, bold=True, color=self.C.WHITE, align=PP_ALIGN.CENTER)
            self._add_box(s, x + 0.8, y + 0.2, fw - 1, 0.9, lb,
                          13, color=self.C.TEXT_D, vAlign=MSO_ANCHOR.MIDDLE)
        
        if class_note:
            ny = self.L.cy + 1.5 + 0.3
            self._add_rrect(s, self.L.cx, ny, self.L.cw, 1.7, fill=self.C.ICE)
            self._add_box(s, self.L.cx + 0.2, ny + 0.05, 5, 0.3,
                          "KLASIFIKASI KELAS JALAN", 11, bold=True, color=self.C.NAVY)
            self._add_box(s, self.L.cx + 0.2, ny + 0.4, self.L.cw - 0.4, 0.9,
                          class_note, 10, color=self.C.TEXT_D)
        
        self._footer(s)
    
    def _build_closing(self, data):
        """
        Closing slide — full-bleed navy.
        
        Data:
            pre_title, main_title, subtitle, source
        """
        s = self._blank()
        bg = s.background; bg.fill.solid(); bg.fill.fore_color.rgb = self.C.NAVY
        self._add_oval(s, -1, -1.5, 4.5, self.C.NAVY_D)
        self._add_oval(s, 8.5, -2, 7, self.C.NAVY_D)
        self._add_oval(s, 10, 4.5, 5, self.C.NAVY_D)
        
        mx = self.L.MARGIN_H
        
        self._add_shape(s, MSO_SHAPE.RECTANGLE, mx, 3.6, 3.5, 0.04, fill=self.C.GOLD)
        
        pre = data.get("pre_title", "")
        if pre:
            self._add_box(s, mx, 1.6, self.L.cw, 0.4, pre, 14, bold=True, color=self.C.GOLD)
        
        main = data.get("main_title", "")
        if main:
            self._add_box(s, mx, 2.4, self.L.cw, 1.5, main, 48, bold=True, color=self.C.WHITE)
        
        sub = data.get("subtitle", "")
        if sub:
            self._add_box(s, mx, 4.1, self.L.cw, 0.8, sub, 14, color=self.C.ICE)
        
        src = data.get("source", "")
        if src:
            self._add_box(s, mx, 5.1, self.L.cw, 0.35, src, 10, color=self.C.TEXT_L)
        
        self._add_shape(s, MSO_SHAPE.RECTANGLE, 0,
                        self.L.SLIDE_H - 0.22, self.L.SLIDE_W, 0.22, fill=self.C.GOLD)


# ═══════════════════════════════════════════════════════════
# SHORTCUT — langsung jalan kalau dipanggil
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Demo / test
    slides = [
        {"type": "cover", "data": {
            "pre_title": "DEMO", "city": "TEST CITY",
            "main_title": "PERATURAN DEMO", "main_subtitle": "NOMOR 1 TAHUN 2026",
            "display_title": "DEMO\nGENERIC ENGINE", "badge_text": "Test  ·  2026"
        }},
        {"type": "closing", "data": {
            "pre_title": "DEMO", "main_title": "SELESAI",
            "subtitle": "Engine generic berhasil", "source": "ppt_engine.py"
        }},
    ]
    engine = Engine()
    engine.build(slides, source_text="Sumber: Demo", output_path="/tmp/ppt_engine_demo.pptx")
    print(f"✅ Demo OK: /tmp/ppt_engine_demo.pptx ({len(engine.prs.slides)} slide)")
