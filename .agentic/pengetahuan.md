# Pengetahuan — Proyek PPT Generator

## Arsitektur Terkini (src/)

```
📁 coba_ppt/
├── src/                          ← Semua skrip
│   ├── __init__.py               ← Biar from src.ppt_engine import Engine
│   ├── ppt_engine.py             ← Engine utama (LayoutFrame + archetypes + PaletteGenerator)
│   ├── buat_ppt_generik.py       ← Entry point (CONTENT_MODULE=... python3 src/buat_ppt_generik.py)
│   ├── create_ppt.py             ← Legacy standalone script
│   ├── qa.py                     ← Visual QA: render PPTX → images (butuh LibreOffice + pdftoppm)
│   ├── pptx_tools.py             ← Template editing: unpack/list/pack
│   ├── fix_pptx_zip.py           ← Fix ZIP entry order PPTX corrupt (PptxGenJS issue)
├── agent.md                      ← Panduan LLM (di root)
├── *_tmp.py                      ← Content module (di-ignore git via .gitignore)
├── .gitignore                    ← *_tmp.py diabaikan
└── .agentic/
    └── pengetahuan.md             ← File ini
```

### Cara Pakai Berbagai Mode

```bash
# Mode 1 — LLM inline (tanpa file)
python3 -c "from src.ppt_engine import Engine; Engine().build(SLIDES, ...)"

# Mode 2 — Content module (_tmp.py)
CONTENT_MODULE=content_xxx_tmp python3 src/buat_ppt_generik.py

# Mode 3 — Auto palette dari 1 warna
python3 -c "from src.ppt_engine import Engine; Engine(primary_color='#E91E63')"

# Mode 4 — Font style
python3 -c "from src.ppt_engine import Engine; Engine(font_style='modern')"

# Mode 5 — QA render
python3 src/qa.py output.pptx --inspect

# Mode 6 — Template editing
python3 src/pptx_tools.py list template.pptx
```

### Naming Convention
- Content module → **`*_tmp.py`** (di-ignore git)
- Skrip utama → **`src/*.py`** (tracked)

---

## Root Cause: PptxGenJS Corrupt di PowerPoint

**Bukan** masalah XML/shape. Masalahnya **ZIP entry order**.

- ECMA-376 OPC §9.1.2.1: `[Content_Types].xml` WAJIB sebagai stream pertama dalam ZIP.
- PptxGenJS menaruh `[Content_Types].xml` di entry #19 → PowerPoint tolak.
- **Fix**: `python3 src/fix_pptx_zip.py <file.pptx>` — tulis ulang ZIP dengan entry order benar.
- **Alternative**: `python-pptx` selalu generate `[Content_Types].xml` di posisi 0 ✅

---

## LayoutFrame Class — Rumus Kalkulasi

Semua ukuran dalam **inches**. Slide 16:9 = **13.333" × 7.5"**.

### Zona Layout
```
Margin H      = 0.6"
Header        = 0.94" (gold_bar 0.04" + navy 0.9")
Header Gap    = 0.21"
Footer        = 0.50"
─────────────────────────────────
Content Area: cx=0.6, cy=1.15, cw=12.133, ch=5.85
```

### Rumus Utama

| Rumus | Formula |
|-------|---------|
| **Lebar Kolom** | `col_w = (cw - (n-1) × gap) / n` |
| **Tinggi Baris** | `row_h = (ch - (n-1) × gap_v) / n` |
| **Posisi Grid** | `x = cx + col × (col_w + gap_h)` |
| | `y = cy + row × (row_h + gap_v)` |
| **Estimasi Tinggi Teks** | `n_lines = ceil(len(text) / (cpi × box_w))` |
| | `height = n_lines × pt × 1.2 / 72` |
| **Safe Item Height** | `max(text_height + 0.05, min_h)` |

---

## Font Pairings — 8 Style

| Style | Header | Body | Kesan |
|-------|--------|------|-------|
| `classic` | Georgia | Calibri | Profesional, formal |
| `modern` | Arial Black | Arial | Berani, kontemporer |
| `clean` | Calibri | Calibri Light | Minimalis |
| `formal` | Cambria | Calibri | Resmi, akademik |
| `tech` | Consolas | Calibri | Teknis |
| `elegant` | Palatino | Garamond | Elegan, klasik |
| `bold` | Impact | Arial | Sangat berani |
| `friendly` | Trebuchet MS | Calibri | Ramah |

```python
Engine(font_style="modern")
Engine(font_style="formal", primary_color="#E91E63")
```

---

## Icon System — 8 Built-in Shapes + Emoji

Engine method `_add_icon(slide, x, y, size, icon="...", fill=..., icon_color=...)`

| Nama | Simbol | Penggunaan |
|------|--------|------------|
| `'check'` | ✓ | Centang / selesai |
| `'x'` | ✗ | Silang / salah |
| `'arrow'` | → | Panah / next |
| `'star'` | ★ | Bintang / favorit |
| `'circle'` | ● | Circle / default |
| `'info'` | ℹ | Informasi |
| `'warning'` | ⚠ | Peringatan |
| `'question'` | ? | Pertanyaan |

Support juga emoji langsung: `{"icon": "🔵", ...}`

---

## Color Theory — Auto WCAG AA Palette

### `PaletteGenerator` class
Menerima 1 warna utama (hex) → generate full 60-30-10 + 4 semantic.

### Rules
| Rule | Output | Detail |
|------|--------|--------|
| **Monochromatic** | NAVY, NAVY_L, ICE (60%+30%) | Same hue, varian lightness 6%–92% |
| **Split-complementary** | GOLD accent (10%) | h + 150° dan h + 210° |
| **Tetradic** | BLUE, TEAL, WARM, RED | 4 hues 90° apart |
| **WCAG binary-search** | Semua teks | Adjust lightness sampai ≥4.5:1 |

### Acuan
- Rathore et al., VIS 2019 — arXiv:1908.00220
- Li et al., 2021 — arXiv:2102.05231
- WCAG 2.1 / sRGB ITU-R BT.709

```python
engine = Engine(primary_color="#2563EB")
# → otomatis: TEXT_D 17.1:1, BLUE 8.0:1, TEAL 5.0:1, WARM 4.6:1, RED 4.5:1
```

Juga dari CLI: `python3 src/ppt_engine.py "#E91E63"`

---

## 11 Slide Archetype

| Type | Fungsi | Elemen Kunci |
|------|--------|--------------|
| `cover` | Halaman sampul | Full-bleed navy, ovals, gold bar, judul 40pt |
| `toc` | Daftar isi | 2 kolom, badge nomor + label, per item |
| `section` | Pembatas bab | Full-bleed dark, ovals dekoratif, action title |
| `content` | Slide standar | Navy header + gold bar + footer |
| `card_grid` | Grid card multi-baris | Icon circle, title, bullet items, auto item height |
| `two_col` | 2 card bersebelahan | Accent bar kiri, $ untuk highlight, bullet |
| `callout` | Big number cards | Angka 32pt, label, note box |
| `flow` | Horizontal flow | Numbered circles, arrow `›`, note box |
| `table` | Native table | Header navy, zebra striping (ICE/WHITE) |
| `nsr_factors` | Faktor NSR custom | Oval numbers 4+3 grid + classification box |
| `closing` | Penutup | Full-bleed navy, mirip cover |

---

## Color Palette Default — WCAG AA Verified ✅

### 60-30-10
| Peran | Warna | Hex | Kontras Putih |
|-------|-------|-----|:-------------:|
| **60%** Dominan | Navy | `#0A1628` | 18.1:1 🏆 |
| **30%** Sekunder | Off White | `#F5F7FA` | 1.1:1 |
| **10%** Aksen | Gold | `#C8962E` | 2.7:1 (dark bg only) |

### Semantic (≥4.5:1 ✅)
- BLUE `#2563EB` (5.2:1) — info
- TEAL `#0B7C72` (5.1:1) — prosedur
- WARM `#A0522D` (5.6:1) — peringatan
- RED `#DC2626` (4.8:1) — bahaya

### Text
- TEXT_D `#1A1A2E` (17.1:1 on white)
- TEXT_M `#8899B0` (6.2:1 on navy)
- TEXT_L `#64748B` (4.8:1 on white)

---

## QA Workflow

```bash
# Text check
python3 src/qa.py --text-only output.pptx

# Render to images (butuh LibreOffice + poppler-utils)
python3 src/qa.py --render-only output.pptx --dpi 150

# Full with subagent inspect prompt
python3 src/qa.py --inspect output.pptx
```

### Verification Loop (WAJIB)
1. Generate → Render → Subagent inspect
2. Fix issues → Re-verify
3. Ulang sampai no new issues

Subagent dipanggil via tool `task()` atau `agentic_delegate()`.

---

## Template Editing

```bash
python3 src/pptx_tools.py unpack template.pptx unpacked/
python3 src/pptx_tools.py list template.pptx
# edit XML di unpacked/ppt/slides/slide*.xml
python3 src/pptx_tools.py pack unpacked/ output.pptx
```
