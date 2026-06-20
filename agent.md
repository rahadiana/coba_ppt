# 🎯 PPT Generator — Agent Navigation

> **Skill**: Membuat presentasi PowerPoint profesional dengan layout presisi, action titles (McKinsey style), visual hierarchy, bebas overlap, kompatibel Microsoft PowerPoint.
>
> **Cocok untuk**: Regulasi, laporan, proposal, data — apapun yang perlu presentasi cepat.

---

## 🤖 LLM Workflow (Cara Agent Menggunakan Skill Ini)

**Tidak perlu** menulis file `content_*.py` dulu. LLM bisa langsung:

```
[Baca dokumen sumber]
       ↓
[Analisis & ekstrak konten → susun slide]
       ↓
[Generate dict SLIDES langsung di memori]
       ↓
[Panggil engine via satu perintah python]
```

### Step-by-step untuk LLM:

**Step 1 — Baca dokumen sumber**
```python
# Baca file PDF / teks / URL
with open("dokumen.pdf", "rb") as f:
    text = extract_text(f)  # atau pakai webfetch / read tool
```

**Step 2 — Analisis & susun slide**
- Ekstrak: judul, bab, pasal, tabel, definisi, poin-poin penting
- Buat action title untuk tiap slide (judul = **kesimpulan**, bukan topik)
- Pilih archetype yang tepat (lihat panduan di bawah)

**Step 3 — Generate langsung**
```python
from src.ppt_engine import Engine

SLIDES = [
    {"type": "cover", "data": {...}},
    {"type": "toc", "data": {...}},
    {"type": "section", "data": {...}},
    {"type": "card_grid", "data": {...}},
    # ... dst sesuai hasil analisis
]

engine = Engine()
engine.build(SLIDES,
             source_text="Sumber: Dokumen X",
             output_path="output.pptx")
print("✅ PPT siap: output.pptx")
```

Jalankan dengan:
```bash
python3 -c "
from src.ppt_engine import Engine
# ... paste SLIDES di sini ...
engine = Engine()
engine.build(SLIDES, source_text='...', output_path='output.pptx')
"
```

Atau simpan sebagai file temp `_gen.py`, jalankan, lalu hapus.

---

## 📋 Pilih Archetype yang Tepat

| Konten | Archetype | Alasan |
|--------|-----------|--------|
| Judul besar + info dasar | `cover` | 1 slide pembuka |
| Daftar bab/sesi | `toc` | 1 slide navigasi |
| Pembatas antar bab | `section` | tiap bab 1 divider |
| Definisi, istilah, poin pendek | `card_grid` | max 8 card per slide |
| Pro/kontra, bandingan 2 sisi | `two_col` | 2 kolom (kiri/kanan) |
| Data statistik, angka besar | `callout` | max 4 angka besar |
| Proses, alur, tahapan | `flow` | max 5 step |
| Data tabular, perbandingan | `table` | banyak baris/kolom |
| Faktor-faktor dengan angka | `nsr_factors` | khusus faktor + catatan |
| Penutup | `closing` | 1 slide akhir |

### Aturan Pemilihan Archetype

1. **Tiap bab** → `section` divider dulu, lalu slide isi
2. **Definisi 3-8 item** → `card_grid` (max 8 card, auto 4+4 atau 4+3)
3. **Definisi 2 item** → `card_grid` atau `two_col`
4. **Tabel data** → `table` (header navy, zebra stripe)
5. **Alur/proses** → `flow` (numbered steps + arrow)
6. **Angka menonjol** → `callout` (angka 32pt bold)
7. **Informasi umum** → `content` (header + footer aja)

---

## 🧩 11 Slide Archetype + Data Format

### 1. `cover` — Halaman Sampul
```python
{"type": "cover", "data": {
    "pre_title": "BERITA DAERAH",        # atas, gold 12pt bold
    "city": "KOTA BEKASI",               # 14pt white bold
    "main_title": "PERATURAN...",         # 20pt white bold
    "main_subtitle": "NOMOR ...",        # 15pt gold bold
    "display_title": "TENTANG\n...",     # 40pt white bold — judul visual
    "badge_text": "Kota Bekasi · 2024",  # 12pt ice (badge navy_d)
    "badge_subtext": "Berlaku sejak...", # 10pt text_l
}}
```

### 2. `toc` — Daftar Isi
```python
{"type": "toc", "data": {
    "title": "Action Title",                    # kesimpulan daftar isi
    "subtitle": "Subtitle",
    "cols": 2,                                  # jumlah kolom
    "items": [
        {"num": "1", "label": "Bab 1",        "color": "#2563EB"},
        {"num": "2", "label": "Bab 2",        "color": "#0D9488"},
        # ... bisa sampai 12 item (2 kolom × 6 baris)
    ],
}}
```
⚠️ Max ~12 item. Lebih? split jadi 2 slide TOC.

### 3. `section` — Section Divider (full-bleed navy)
```python
{"type": "section", "data": {
    "title": "BAB I\nKETENTUAN UMUM",     # 34pt bold white (bisa multi-line)
    "subtitle": "Pasal 1",                # 11pt text_l
    "action_text": "7 Definisi Kunci",    # 14pt gold bold — inti bab
}}
```
**Wajib** sebelum tiap bab. Action text = ringkasan 1 kalimat bab.

### 4. `content` — Standar Header + Footer
```python
{"type": "content", "data": {
    "title": "Action Title",
    "subtitle": "Subtitle",
}}
```
Untuk konten yang dibuat manual dengan shapes tambahan.

### 5. `card_grid` — Multi-row Cards
```python
{"type": "card_grid", "data": {
    "title": "Action Title",                    # kesimpulan
    "subtitle": "Pasal 1",
    "cols": 0,                                  # 0=auto, atau 2/3/4
    "cards": [
        {"icon": "🏛️", "title": "Judul",
         "color": "#2563EB",
         "items": ["Item 1", "Item 2"]},        # jadi "• Item 1"
        {"icon": "📊", "title": "Judul 2",
         "color": "#0D9488",
         "items": ["A", "B", "C"]},
    ],
}}
```
- Max **4 item** per card (biar muat)
- Max **8 card** per slide (4+4 grid)
- Auto calculate item height → aman dari overlap

### 6. `two_col` — 2 Column Cards
```python
{"type": "two_col", "data": {
    "title": "Action Title",
    "subtitle": "Pasal 6–8",
    "left": {
        "color": "#2563EB",
        "lines": [
            "$JUDUL SECTION",              # $ → highlight 14pt bold
            "",                             # baris kosong = spacer
            "teks biasa",                   # → "• teks biasa"
            "teks lain",
        ],
    },
    "right": {
        "color": "#0D9488",
        "lines": ["$SECTION 2", "", "item 1", "item 2"],
    },
}}
```
- Cocok untuk perbandingan kiri/kanan
- Baris mulai `$` = highlight title
- Baris `""` = spacer

### 7. `callout` — Big Number Cards
```python
{"type": "callout", "data": {
    "title": "Action Title",
    "subtitle": "Pasal 5",
    "callouts": [
        {"number": "12", "label": "Bulan\n(Permanen)",   "color": "#2563EB"},
        {"number": "30", "label": "Hari\n(Insidentil)",  "color": "#0D9488"},
    ],
    "note": "• Note 1\n• Note 2",   # optional — box di bawah callout
}}
```
- Max **4 callout** per slide
- Number = 32pt bold, label = 11pt
- `\n` untuk multi-line label

### 8. `flow` — Horizontal Flow
```python
{"type": "flow", "data": {
    "title": "Alur 4 Langkah",
    "subtitle": "Pasal 11–14",
    "steps": [
        {"num": "1", "title": "SKPD",
         "desc": "Diterbitkan Bapenda\nMasa 5 tahun",
         "color": "#2563EB"},
        {"num": "2", "title": "Bayar",
         "desc": "Lunas 1 bln",
         "color": "#0D9488"},
    ],
    "note": "• Jatuh tempo: 1 bulan",   # optional
}}
```
- Max **5 step** per slide
- Step = numbered circle 0.5" + title + desc
- Arrow `›` gold di antara step

### 9. `table` — Native Table
```python
{"type": "table", "data": {
    "title": "Tabel NSR",
    "subtitle": "Pasal 10",
    "headers": ["Kelas Jalan", "Zona", "NSR"],    # navy background
    "rows": [
        ["Kelas Khusus", "Tol", "23.575"],          # zebra stripe
        ["Kelas I", "Ketat", "13.225"],
    ],
}}
```
- Header navy, row bergantian ICE/WHITE
- Kolom pertama left-align, sisanya center

### 10. `nsr_factors` — Faktor NSR
```python
{"type": "nsr_factors", "data": {
    "title": "7 Faktor Penentu NSR",
    "subtitle": "Pasal 9",
    "factors": [
        {"num": "1", "label": "Jenis Reklame",    "color": "#2563EB"},
        {"num": "2", "label": "Bahan",             "color": "#0D9488"},
        # max 7 item (4+3 grid)
    ],
    "classification_note": "🏛️ Kelas Jalan Khusus\n🚗 Kelas Jalan I",
}}
```
Khusus untuk faktor-faktor bernomor + kotak catatan di bawah.

### 11. `closing` — Penutup
```python
{"type": "closing", "data": {
    "pre_title": "BERITA DAERAH KOTA BEKASI",  # 14pt gold
    "main_title": "TERIMA KASIH",               # 48pt white bold
    "subtitle": "Peraturan ...\nTentang ...",   # 14pt ice
    "source": "Sumber: https://...",             # 10pt text_l
}}
```

---

## 🎨 Color System (60-30-10 Rule) — WCAG AA Verified ✅

Semua warna sudah diuji kontras dengan formula WCAG 2.1 — **no color bias**.

### Palette Inti

| Nama | Hex | Peran | Kontras Putih | Kontras Navy | Penggunaan |
|------|-----|-------|--------------|--------------|------------|
| Navy | `#0A1628` | **60%** dominan | 18.1:1 🏆 | — | Background, header, section divider |
| Off White | `#F5F7FA` | **30%** sekunder | 1.1:1 | 16.9:1 🏆 | Background content slide |
| White | `#FFFFFF` | card bg | — | 18.1:1 🏆 | Card background, teks header |
| Gold | `#C8962E` | **10%** aksen | 2.7:1 | 6.8:1 ✅ | Bar, highlight, arrow (hanya di dark bg) |

### Semantic Colors

| Nama | Hex | Kontras Putih | WCAG | Penggunaan |
|------|-----|:------------:|:----:|------------|
| Blue | `#2563EB` | 5.2:1 | ✅ AA | Info, definisi, card 1 |
| **Teal** | **`#0B7C72`** | **5.1:1** | **✅ AA** | **Prosedur, data, card 2** |
| **Warm** | **`#A0522D`** | **5.6:1** | **✅ AA** | **Peringatan, faktor, card 3** |
| Red | `#DC2626` | 4.8:1 | ✅ AA | Sanksi, bahaya, card 4 |

### Text Colors

| Nama | Hex | Kontras Putih | Kontras Navy | Untuk |
|------|-----|:------------:|:------------:|-------|
| TEXT_D | `#1A1A2E` | 17.1:1 🏆 | — | Body text di light bg |
| **TEXT_M** | **`#8899B0`** | 2.9:1 | **6.2:1 ✅** | **Subtitle di navy header** |
| **TEXT_L** | **`#64748B`** | **4.8:1 ✅** | 3.8:1 | **Footer source (7pt)** |

> ⚡ **Perubahan dari palette sebelumnya** (hasil analisis WCAG):
> - `TEAL` #0D9488 → **#0B7C72** (digelapkan: 3.7→5.1:1 on white ✅)
> - `WARM` #B8860B → **#A0522D** (ganti hue ke sienna ~20°, tidak clash dengan gold 41°)
> - `TEXT_M` #6B7288 → **#8899B0** (dicerahkan: 3.8→6.2:1 on navy ✅)
> - `TEXT_L` #9CA3AF → **#64748B** (digelapkan: 2.5→4.8:1 on white ✅)

### Rotasi Warna untuk Card
```python
# Rotasi 4 warna — semua ≥ 4.5:1 on white ✅
colors = ["#2563EB", "#0B7C72", "#A0522D", "#1B3A6B"]
for i, card in enumerate(cards):
    card["color"] = colors[i % len(colors)]
```

### Aturan Pakai Warna (anti bias)

| Background | Teks Aman | Teks Hindari |
|------------|-----------|-------------|
| Navy `#0A1628` | White, Gold, TEXT_M, TEAL, WARM | TEXT_D, BLUE, TEXT_L |
| White `#FFFFFF` | TEXT_D, BLUE, TEAL, WARM, RED | Gold, TEXT_M, TEXT_L |
| Off White `#F5F7FA` | TEXT_D, BLUE, TEAL, WARM, RED | Gold, TEXT_M, TEXT_L |
| ICE `#E8EDF5` | TEXT_D, BLUE, WARM | Gold, TEAL, TEXT_M |

---

## 🎨 Auto Palette — Dari Satu Warna Utama 🆕

Jika user hanya memberi **satu warna utama**, engine bisa otomatis generate seluruh palette 60-30-10 + 4 warna semantik, semuanya WCAG AA verified.

### Cara Pakai

```python
from src.ppt_engine import Engine

engine = Engine(primary_color="#2563EB")
# → otomatis generate palette, print WCAG report:
#   ✅ TEXT_D on WHITE: 17.1:1
#   ✅ BLUE on WHITE: 6.3:1
#   ✅ TEAL on WHITE: 5.0:1
#   ✅ WARM on WHITE: 4.6:1
#   ✅ RED on WHITE: 4.5:1
```

Atau langsung dari CLI:

```bash
python3 src/ppt_engine.py "#E91E63"
# → generate 3-slide PPT + palette report
```

### Color Harmony Rules

| Rule | Deskripsi | Digunakan Untuk |
|------|-----------|----------------|
| **Split-complementary** | Base hue + 2 adjacent to complement (150°, 210°) | **10% Accent** (GOLD) — kontras maksimal di navy |
| **Tetradic** | 4 hues 90° apart | **Semantic colors** (BLUE, TEAL, WARM, RED) — 4 slot berbeda |
| **Monochromatic** | Same hue, varied lightness (6%-92%) | **60% Dominant** (NAVY variants) + **30% Secondary** (ICE variants) |
| **WCAG Auto-Adjust** | Binary search lightness hingga ≥4.5:1 | Semua teks dan warna semantik |

### Output Palette (dict → `Colors`-like object)

```python
engine.C.NAVY   # → RGBColor — dark background
engine.C.GOLD   # → RGBColor — accent
engine.C.BLUE   # → RGBColor — semantic 1
engine.C.TEAL   # → RGBColor — semantic 2
engine.C.WARM   # → RGBColor — semantic 3
engine.C.RED    # → RGBColor — semantic 4
engine.C.TEXT_D # → RGBColor — body text
engine.C.TEXT_M # → RGBColor — subtitle on navy
engine.C.TEXT_L # → RGBColor — footer
```

> **Catatan**: Label `BLUE`/`TEAL`/`WARM`/`RED` adalah nama slot semantik. Hue aktual tergantung primary color (tetradic rotation). Misal primary pink → BLUE adalah merah-jambu, RED adalah ungu-biru.

### Acuan

- WCAG 2.1 Relative Luminance & Contrast Ratio (Web Content Accessibility Guidelines)
- Color-by-concept association (Rathore et al., VIS 2019, [arXiv:1908.00220](https://arxiv.org/abs/1908.00220))
- Culture-inspired palette generation (Li et al., 2021, [arXiv:2102.05231](https://arxiv.org/abs/2102.05231))
- ITU-R BT.709 sRGB linearization for luminance (Rec. 709 / IEC 61966-2-1)

---

## 📐 Layout Rules — Jangan Dilanggar!

### Zona Slide (16:9 = 13.333" × 7.5")
```
┌─ 0.035" gold_bar ────────────────────────────┐
│ 0.9" navy header   [Judul 20pt bold white]   │ ← HEADER (0.94")
│ 0.21" gap                                     │
├───────────────────────────────────────────────┤
│                                               │
│    CONTENT AREA: 12.133" × 5.85"              │
│    margin kiri/kanan = 0.6"                   │
│    cx=0.6", cy=1.15"                          │
│                                               │
├───────────────────────────────────────────────┤
│ 0.03" navy bar   [source]            [page #] │ ← FOOTER (0.50")
└───────────────────────────────────────────────┘
```

### Rumus Otomatis (dihitung engine, LLM tidak perlu hafal)
```
col_width(N)      = (12.133 - (N-1) × 0.3) / N
row_height(N)     = (5.85 - (N-1) × 0.3) / N
text_height()     = ceil(len / (cpi × box_w)) × pt × 1.2 / 72
safe_item_height  = max(text_height + 0.05, 0.25)
```

### Action Title — WAJIB
Setiap slide (kecuali cover & closing) harus punya **action title**:
- ✅ **Benar**: "7 Definisi Kunci Menjadi Landasan Pengelolaan Pajak"
- ❌ **Salah**: "Definisi" / "Bab I" / "Pendahuluan"
- Action title = **kesimpulan**, bukan topik

---

## 📐 Panduan Ekstraksi Konten untuk LLM

Saat membaca dokumen sumber, ikuti pola ini:

### 1. Identifikasi Bab
```
Dokumen → Bab I, Bab II, ... → tiap bab = 1 section divider + slide isi
```

### 2. Pilih Slide Isi per Bab
```
Bab kecil (1 pasal)    → 1 card_grid
Bab sedang (2-3 pasal) → 2-3 slide (card_grid / two_col / table)
Bab besar (4+ pasal)   → 3-5 slide (mix archetype)
Ada tabel data         → table_slide
Ada proses/alur        → flow_slide
Ada angka penting      → callout_slide
```

### 3. Struktur Presentasi Ideal (33 slide untuk dokumen 11 bab)
```
Cover (1)
TOC (1)
Untuk tiap bab:
  Section divider (1)
  Slide isi (1-3)
Closing (1)
────────────────
Total: ~20-40 slide
```

### 4. Aturan Jumlah Item
| Archetype | Max Item | Notes |
|-----------|----------|-------|
| card_grid | 8 card | 4+4 grid, auto |
| two_col | ~15 baris per kolom | tergantung panjang teks |
| callout | 4 angka | 4 kolom |
| flow | 5 step | horizontal |
| table | unlimited | native scroll |
| nsr_factors | 7 faktor | 4+3 grid |

### 5. Icon Emoji untuk Card
Gunakan emoji yang relevan:
```
🏛️ pemerintah/daerah     📊 data/statistik
📢 promosi/reklame        💰 pajak/keuangan
📐 ukuran/nilai           🆔 identitas/NPWPD
👤 orang/wajib pajak      🚫 larangan/pengecualian
🔍 pemeriksaan            📬 surat/tagihan
⏳ waktu/kadaluwarsa      🎯 keringanan/fasilitas
🤝 kemudahan/kerjasama    🏆 penghargaan
🎈 udara                  🌊 apung/air
🎬 film/media             🎭 peragaan/event
🏷️ stiker/label           🧱 melekat
📄 selebaran              🚌 kendaraan/berjalan
```

---

## 🚀 Cara Generate (3 Opsi)

### Opsi 1 — Langsung dari CLI (paling cepat)
```bash
python3 -c "
from src.ppt_engine import Engine
SLIDES = [  # ← paste hasil analisis di sini
    {'type':'cover','data':{...}},
    {'type':'section','data':{...}},
]
engine = Engine()
engine.build(SLIDES, source_text='Sumber: ...', output_path='output.pptx')
"
```

### Opsi 2 — Lewat file temp (untuk konten besar)
```python
# _gen.py — buat, jalankan, hapus
from src.ppt_engine import Engine
SLIDES = [...]  # konten
Engine().build(SLIDES, source_text='...', output_path='output.pptx')
```
```bash
python3 _gen.py && rm _gen.py
```

### Opsi 3 — Lewat src/buat_ppt_generik.py (dengan file content)
Hanya jika konten akan dipakai berulang:
```bash
CONTENT_MODULE=content_xxx python3 src/buat_ppt_generik.py
```

---

## ✅ Referensi File

| File | Fungsi | Wajib? |
|------|--------|--------|
| `src/ppt_engine.py` | Engine — LayoutFrame + 11 archetype builders | ✅ Ya |
| `src/buat_ppt_generik.py` | Entry point untuk content file | 🔧 Optional |
| `src/fix_pptx_zip.py` | Utility fix ZIP order PPTX corrupt | 🔧 Optional |

> **LLM tidak perlu** menyentuh `src/ppt_engine.py`. Cukup baca `src/agent.md` ini, extract konten dari sumber, lalu panggil engine langsung.

---

## 🛠️ Stack Requirements — Cek Installasi

Sebelum generate, LLM harus cek apakah environment sudah siap:

### Required Stack
| Komponen | Minimal | Untuk |
|----------|---------|-------|
| Python | 3.8+ | Menjalankan engine |
| `python-pptx` | 1.0.0+ | Generate PPTX |

### Cek Installasi

**Cek Python:**
```bash
python3 --version
# Harus: Python 3.8.x atau lebih baru
```

**Cek python-pptx:**
```bash
python3 -c "import pptx; print(pptx.__version__)"
# Harus: 1.0.0 atau lebih baru
```

**Cek engine:**
```bash
python3 -c "from src.ppt_engine import Engine; print('✅ Engine siap')"


# Harus: ✅ Engine siap
```

### Install Jika Belum Ada

```bash
# Install python-pptx
pip install python-pptx

# Atau via pip3
pip3 install python-pptx
```

### Verifikasi Lengkap (satu perintah)
```bash
python3 -c "
import sys, pptx
from src.ppt_engine import Engine, LayoutFrame, Colors
print(f'✅ Python   : {sys.version}')
print(f'✅ python-pptx: {pptx.__version__}')
print(f'✅ Engine   : OK — {len(Engine().build.__code__.co_varnames)} params')
L = LayoutFrame()
print(f'✅ Layout   : cw={L.cw:.3f}, ch={L.ch:.3f}')
print('🎯 Stack siap — bisa generate PPT')
"
```

### Jika Gagal
| Error | Penyebab | Solusi |
|-------|----------|--------|
| `ModuleNotFoundError: pptx` | python-pptx belum install | `pip install python-pptx` |
| `ModuleNotFoundError: src.ppt_engine` | engine file tidak ada | Cek `ls src/ppt_engine.py` |
| `SyntaxError` | Python < 3.8 | Upgrade Python |

---

## 🎨 Visual Design System — Background Shapes, Color & Logo Patterns

Semua pattern design sudah di-engine, LLM cukup paham agar konsisten.

### 1. Background per Tipe Slide

| Slide Type | Background | Warna | Pattern |
|------------|-----------|-------|---------|
| **Cover** | Full-bleed navy | `#0A1628` | 4 decorative ovals di pojok (`NAVY_D` & `NAVY_L`) |
| **Section** | Full-bleed navy | `#0A1628` | 4 decorative ovals (posisi berbeda dari cover) |
| **Content** | Ice solid | `#F5F7FA` | Header navy + gold bar + footer |
| **Card** | Rounded rect white | `#FFFFFF` | Left accent bar 0.05" warna semantic |
| **Callout** | Rounded rect white | `#FFFFFF` | Top accent bar 0.05" warna semantic |
| **Closing** | Full-bleed navy | `#0A1628` | 3 decorative ovals (mirip cover) |

### 2. Decorative Oval Pattern

Oval digunakan sebagai elemen dekoratif di slide gelap (cover, section, closing):

```
Cover:         Section:           Closing:
┌──────────┐   ┌──────────┐       ┌──────────┐
│◉ ○       │   │◉◌        │       │◉         │
│          │   │          │       │          │
│         ◉│   │    ◉     │       │     ◉    │
│ ○        │   │     ◉   ◌│       │         ◉│
└──────────┘   └──────────┘       └──────────┘

Ukuran: 2.5" - 7"
Warna: NAVY_D (#0D1F3C) dan NAVY_L (#12294A)
Posisi: random di luar content area
```

Detail posisi ovals (dalam inches):
```
Cover:
  oval(-1, -1.5, 4.5, NAVY_D)     // kiri atas besar
  oval(8.5, -2, 7, NAVY_D)        // kanan atas besar
  oval(10, 4.5, 5, NAVY_D)        // kanan bawah
  oval(0.5, 5.5, 2.5, NAVY_L)     // kiri bawah kecil

Section:
  oval(-1.5, -1.5, 5, NAVY_L)     // kiri atas
  oval(-0.5, -0.5, 3.5, NAVY_D)   // kiri atas kecil
  oval(9.5, 4, 5, NAVY_D)         // kanan bawah
  oval(10.5, 3, 3, NAVY_L)        // kanan tengah

Closing:
  oval(-1, -1.5, 4.5, NAVY_D)     // kiri atas
  oval(8.5, -2, 7, NAVY_D)        // kanan atas
  oval(10, 4.5, 5, NAVY_D)        // kanan bawah
```

### 3. Header Pattern (Content Slides)

```
┌─────────────────────────────────────────────┐
│ ═══ gold_bar 0.035" (#C8962E)              │ ← y=0
│                                              │
│ ████████████████████████████████████████████ │ ← navy 0.9" (#0A1628)
│ ████│                                        │
│ ██│█│  Judul 20pt bold white                 │
│ ██│█│  Subtitle 9pt TEXT_L                   │
│ ██│█│                                        │
│ ██│█│  0.07" gold accent bar (#C8962E)       │
│ ██│█│  di x=0.6", y=0.12", height=0.55"     │
│                                              │
│ ═══ gap 0.21"                                │
├─────────────────────────────────────────────┤
```

### 4. Card Pattern

Setiap card dalam card_grid punya pattern konsisten:

```
┌───────────────────────────┐
│ ┃  ◯ icon emoji 0.42"    │ ← icon circle (colored, 13pt white)
│ ┃                        │
│ ┃  Title 13pt bold       │ ← colored sesuai card
│ ┃                        │
│ ┃  • Item 1 9pt          │ ← bullet items
│ ┃  • Item 2              │
│ ┃  • Item 3              │
│ ┃                        │
│  0.05" accent bar        │ ← warna semantic di kiri
└───────────────────────────┘
  rounded rectangle
  radius=0.04", border=ICE
  padding internal=0.15"
```

**Ukuran Card:**
```
Card width  (4 kolom) = (12.133 - 3×0.3) / 4 = 2.808"
Card height (2 baris) = (5.85 - 1×0.3) / 2 = 2.775"
Icon circle = 0.42"
Left accent bar = 0.05"
```

### 5. Footer Pattern

```
┌─────────────────────────────────────────────┐
│ ═══ navy bar 0.03" (#0A1628)               │
│ Sumber: ...                    [page #]     │
│ 7pt TEXT_L                     8pt TEXT_L   │
│ x=0.6"                         x=12.333"    │
└─────────────────────────────────────────────┘
  y = 7.0" (SLIDE_H - FOOTER_H)
```

### 6. Logo / Icon System

Tidak ada logo gambar. Semua ikon pakai **emoji** dalam lingkaran warna:

| Ukuran | Diameter | Font Size Emoji | Untuk |
|--------|----------|----------------|-------|
| Icon card | 0.42" | 13pt | card_grid |
| Badge TOC | 0.5" | 14pt | daftar isi |
| Circle flow | 0.5" | 18pt | flow step number |
| Oval NSR | 0.55" | 18pt | nsr_factors |

**Posisi icon dalam card:**
```python
icon_y = 0.15"     # dari atas card
icon_size = 0.42"  # diameter
title_y = icon_y + icon_size + 0.13"  # title di bawah icon
```

**Emoji yang umum dipakai:**
```
🏛️📊📢💰📐🆔👤🚫🔍📬⏳🎯🤝🏆🎈🌊🎬🎭🏷️🧱📄🚌
```

### 7. Warna Rotasi untuk Multi-Card

Saat punya banyak card, rotasi warna otomatis:
```python
colors = [
    "#2563EB",  # Blue  — definisi, info (5.2:1 ✅)
    "#0B7C72",  # Teal  — prosedur, data (5.1:1 ✅)
    "#A0522D",  # Warm  — sienna, peringatan (5.6:1 ✅)
    "#1B3A6B",  # Navy_M — pendukung (11.3:1 🏆)
]
# Untuk 7 card: Blue, Teal, Warm, Navy_M, Blue, Teal, Warm
```

### 8. Visual Hierarchy (Typography)

| Elemen | Font | Size | Weight | Color |
|--------|------|------|--------|-------|
| Cover title | Calibri | 40pt | Bold | White |
| Section title | Calibri | 34pt | Bold | White |
| Header title | Calibri | 20pt | Bold | White |
| Action title | Calibri | 20pt | Bold | White |
| Card title | Calibri | 13pt | Bold | Semantic |
| Card items | Calibri | 9pt | Regular | TEXT_D |
| Callout number | Calibri | 32pt | Bold | Semantic |
| Table header | Calibri | 11pt | Bold | White |
| Table cell | Calibri | 11pt | Regular | TEXT_D |
| Source footer | Calibri | 7pt | Regular | TEXT_L |
| Page number | Calibri | 8pt | Regular | TEXT_L |

### 9. Logo / Icon Placement Mathematics

Semua logo/icon di PPT ini adalah **emoji dalam lingkaran warna** (bukan file gambar). Berikut matematika posisi untuk tiap jenis icon:

#### a) Card Icon — Paling Sering Dipakai

```
Card Grid (setiap card):
┌──────────────────────────┐
│ ┃  ◌  (0.42")            │ ← icon di (cx+0.15, cy+0.15)
│ ┃                        │
│ ┃  Judul Card            │ ← title_y = cy + 0.15 + 0.42 + 0.13
│ ┃  • Item                │
│ ┃                        │
│ 0.05" accent bar         │ ← di (cx, cy) tinggi penuh = row_h
└──────────────────────────┘

Rumus:
  icon_x      = cx + CARD_PAD          = cx + 0.15"
  icon_y      = cy + CARD_PAD          = cy + 0.15"
  icon_size   = 0.42"                  ← diameter lingkaran
  icon_center = icon_x + icon_size/2   ← untuk text alignment CENTER
  
  title_x     = cx + CARD_PAD          = cx + 0.15"
  title_y     = icon_y + icon_size + 0.13"  = cy + 0.70"
  title_width = col_w - CARD_PAD×2     = col_w - 0.30"
```

**Contoh 4 kolom:**
```
col_w = (12.133 - 3×0.3) / 4 = 2.808"
Card 1: icon di (0.6+0.15, cy+0.15) = (0.75, cy+0.15)
Card 2: icon di (0.6+2.808+0.3+0.15, cy+0.15) = (3.858, cy+0.15)
Card 3: icon di (0.6+2×(2.808+0.3)+0.15, cy+0.15) = (6.966, cy+0.15)
Card 4: icon di (0.6+3×(2.808+0.3)+0.15, cy+0.15) = (10.074, cy+0.15)
```

#### b) TOC Badge

```
┌────────────────────┐
│ ┌────┐             │
│ │  1 │  Bab I      │ ← badge di (x+0.10, y+0.06), 0.5×0.5"
│ └────┘             │
│                    │
└────────────────────┘

Rumus:
  badge_x      = x + 0.10"            ← 0.1" dari kiri card
  badge_y      = y + 0.06"            ← 0.06" dari atas card
  badge_size   = 0.5"                 ← square rounded rectangle
  badge_color  = item["color"]
  text_inside  = nomor, 14pt bold white, CENTER
  
  label_x      = x + 0.75"            ← setelah badge + gap
  label_width  = cw_col - 0.9"        ← sisa width card
```

#### c) Flow Step Circle

```
┌──────────────────────┐
│══════ (accent bar)   │ ← y=sy=1.30"
│                      │
│       ◌ (0.5")       │ ← centered: cx+bw/2-0.25, y=1.45"
│       SKPD           │
│  Diterbitkan         │
│  Bapenda             │
└──────────────────────┘

Rumus:
  circle_x     = cx + bw/2 - 0.25"    ← 0.25 = circle_size/2
  circle_y     = sy + 0.15"           ← 0.15" di bawah accent bar
  circle_size  = 0.5"                 ← diameter
  circle_color = step["color"]
  text_inside  = step num, 18pt bold white, CENTER

  ⚡ CIRCLED TEXT: supaya teks center di dalam lingkaran, 
     textbox posisi dan ukuran SAMA dengan circle, align=CENTER
     Jadi teks otomatis di tengah lingkaran.
     
  Arrow "›" antar step:
  arrow_x      = cx + bw              ← persis di kanan card
  arrow_y      = sy + 0.7"            ← vertikal tengah circle area
  arrow_w      = bgap = 0.3"          ← selebar gap
  font_size    = 24pt GOLD
  align        = CENTER               ← center di tengah gap
```

#### d) NSR Factor Circle

```
┌──────────────────┐
│ ◌ (0.55")  label │ ← circle di (x+0.15, y+0.25)
│                  │
└──────────────────┘

Rumus:
  circle_x     = x + 0.15"            ← 0.15" dari kiri card
  circle_y     = y + 0.25"            ← 0.25" dari atas card
  circle_size  = 0.55"                ← diameter (terbesar!)
  circle_color = factor["color"]
  text_inside  = factor num, 18pt bold white, CENTER
  
  label_x      = x + 0.80"            ← setelah circle + gap
  label_y      = y + 0.20"
  label_width  = fw - 1.0"            ← sisa width card
  label_height = 0.9"                 ← tengah vertikal
  vAlign       = MIDDLE
```

#### e) Header Gold Accent Bar

```
┌──────────────────────────────────┐
│══════ gold_bar 0.035" (y=0)     │
│                                  │
│ ████████████████████████████████ │ ← navy header 0.9" (y=0.035)
│ ██│                             │
│ ██│  Judul 20pt white           │
│ ██│  Subtitle 9pt TEXT_L        │
│ ██│                             │
│ ██│  GOLD BAR VERTIKAL          │
│ ██│  x=0.6, y=0.12              │
│ ██│  width=0.07, height=0.55    │
│ ██│                             │
│                                  │
│       gap 0.21"                  │
├──────────────────────────────────┤

Rumus:
  gold_bar_x      = 0.6"            (= MARGIN_H)
  gold_bar_y      = 0.12"           (header atas + padding)
  gold_bar_w      = 0.07"           (tipis vertikal)
  gold_bar_h      = 0.55"           (tinggi — tengah header)
```

#### f) Left Accent Bar pada Card (card_grid & two_col)

```
┌──────────────────────┐
│ ██                   │
│ ██  Card content     │ ← accent bar di (cx, cy)
│ ██                   │    width=0.05", height=ch/row_h
│ ██                   │    color = card["color"]
│ ██                   │
└──────────────────────┘

Rumus:
  bar_x  = cx              ← persis di border kiri card
  bar_y  = cy              ← persis di border atas card
  bar_w  = 0.05"           ← tipis
  bar_h  = row_h (card_grid)  atau ch (two_col)
  color  = card/column color
```

#### g) Top Accent Bar (callout & flow)

```
┌──────────────────────┐
│══════════════════════│ ← bar di (cx, sy), width=cw/bw, height=0.05"
│                      │
│  Callout/Flow card   │
│                      │
└──────────────────────┘
```

#### h) Footer Elements

```
┌─────────────────────────────────────┐
│══════ navy bar 0.03" (y=7.0)       │
│                                     │
│ Sumber: ...              [page #]   │
│ x=0.6"                   x=12.333"  │
│ width=4"                 width=0.8" │
│ font=7pt TEXT_L          font=8pt   │
│                                     │
└─────────────────────────────────────┘
  y = 7.06" (7.0 + 0.06)
```

#### i) Decorative Ovals (Cover / Section / Closing)

Oval dekoratif = **hiasan latar belakang** di slide gelap. Posisinya **sengaja sebagian di luar slide** untuk efek lembut.

| Slide | Oval | x | y | Size | Warna | Efek Posisi |
|-------|------|---|---|------|-------|-------------|
| Cover | 1 | -1.0 | -1.5 | 4.5 | NAVY_D | pojok kiri atas (66% visible) |
| Cover | 2 | 8.5 | -2.0 | 7.0 | NAVY_D | pojok kanan atas (71% visible) |
| Cover | 3 | 10.0 | 4.5 | 5.0 | NAVY_D | pojok kanan bawah (40% visible) |
| Cover | 4 | 0.5 | 5.5 | 2.5 | NAVY_L | pojok kiri bawah, lebih terang |
| Section | 1 | -1.5 | -1.5 | 5.0 | NAVY_L | kiri atas (lebih terang) |
| Section | 2 | -0.5 | -0.5 | 3.5 | NAVY_D | kiri atas nested |
| Section | 3 | 9.5 | 4.0 | 5.0 | NAVY_D | kanan bawah |
| Section | 4 | 10.5 | 3.0 | 3.0 | NAVY_L | kanan tengah (kecil) |
| Closing | 1 | -1.0 | -1.5 | 4.5 | NAVY_D | = cover oval 1 |
| Closing | 2 | 8.5 | -2.0 | 7.0 | NAVY_D | = cover oval 2 |
| Closing | 3 | 10.0 | 4.5 | 5.0 | NAVY_D | = cover oval 3 |

**Pattern Matematika Oval:**
```
Bagian oval yang terlihat dalam slide:
  visible_x_start = max(0, x)
  visible_y_start = max(0, y)
  visible_x_end   = min(SLIDE_W, x + size)
  visible_y_end   = min(SLIDE_H, y + size)
  
Cover Oval 1: terlihat dari (0,0) sampai (3.5, 3.0) = 66% area
Cover Oval 2: terlihat dari (8.5, 0) sampai (13.333, 5.0) = 71% area
Cover Oval 3: terlihat dari (10.0, 4.5) sampai (13.333, 7.5) = 40% area
```

#### j) Ringkasan Semua Ukuran Icon/Logo

| Jenis Icon | Bentuk | Ukuran | Posisi X | Posisi Y | Warna |
|------------|--------|--------|----------|----------|-------|
| **Card icon** | Lingkaran | 0.42" | `cx + 0.15"` | `cy + 0.15"` | Card color |
| **Flow circle** | Lingkaran | 0.50" | `cx + bw/2 - 0.25"` | `sy + 0.15"` | Step color |
| **NSR circle** | Lingkaran | 0.55" | `x + 0.15"` | `y + 0.25"` | Factor color |
| **TOC badge** | Rounded rect | 0.50" | `x + 0.10"` | `y + 0.06"` | Item color |
| **Header accent** | Rectangle | 0.07×0.55" | `0.60"` | `0.12"` | GOLD |
| **Card accent bar** | Rectangle | 0.05×row_h | `cx` | `cy` | Card color |
| **Top accent bar** | Rectangle | cw×0.05 | `cx` | `sy` | Step color |
| **Gold separator** | Rectangle | 2.5-4×0.04" | `MARGIN_H` | varies | GOLD |
| **Bottom bar** | Rectangle | full×0.22" | `0` | `7.28"` | GOLD |
| **Gold top bar** | Rectangle | full×0.035" | `0` | `0` | GOLD |
| **Navy footer bar** | Rectangle | full×0.03" | `0` | `7.00"` | NAVY |

#### k) Aturan Emoji dalam Lingkaran

```
Emoji di dalam lingkaran:
  • Posisi textbox = PERSIS sama dengan posisi lingkaran
  • Ukuran textbox = PERSIS sama dengan diameter lingkaran
  • Alignment = CENTER (horizontal & vertical)
  • Font size = 13pt (card icon), 18pt (flow/NSR circle)
  • Warna font = WHITE

  ⚡ RAHASIA: karena emoji tidak selalu ter-center sempurna,
     textbox yang sama persis dengan lingkaran memastikan 
     emoji muncul di tengah lingkaran.
```

### 10. Contoh: Cover Visual Layout

```
┌────────────────────────────────────────┐
│              ◉ (oval NAVY_D)           │
│  BERITA DAERAH (12pt gold bold)        │
│  KOTA BEKASI (14pt white bold)         │
│                                        │
│  PERATURAN [...] (20pt white bold)     │
│  NOMOR 51 TAHUN 2024 (15pt gold bold)  │
│  ════════ gold bar 0.04"              │
│                                        │
│  TENTANG                               │
│  PENGELOLAAN PAJAK REKLAME             │
│  40pt white bold                       │
│                                        │
│  ┌──────────────────────────────┐      │
│  │ Pemerintah Kota Bekasi · 2024│      │ ← badge NAVY_D
│  │ Berlaku sejak diundangkan    │      │
│  └──────────────────────────────┘      │
│                        ◉ (oval NAVY_D) │
│  ════════ gold bar 0.22"              │
└────────────────────────────────────────┘
```

---

## 🆕 Font Pairings — Header / Body Font

Engine punya 8 font style yang bisa dipilih via `font_style=` parameter.

| Style | Header | Body | Mono | Kesan |
|-------|--------|------|------|-------|
| `classic` (default) | Georgia | Calibri | Consolas | Profesional, formal |
| `modern` | Arial Black | Arial | Consolas | Berani, kontemporer |
| `clean` | Calibri | Calibri Light | Consolas | Minimalis, bersih |
| `formal` | Cambria | Calibri | Consolas | Resmi, akademik |
| `tech` | Consolas | Calibri | Consolas | Teknis, modern |
| `elegant` | Palatino | Garamond | Consolas | Elegan, klasik |
| `bold` | Impact | Arial | Consolas | Sangat berani, poster |
| `friendly` | Trebuchet MS | Calibri | Consolas | Ramah, approachable |

### Cara Pakai

```python
engine = Engine(font_style="modern")           # Arial Black + Arial
engine = Engine(font_style="formal")            # Cambria + Calibri
engine = Engine(primary_color="#E91E63", font_style="modern")
```

### Di Slide Dict

Font otomatis terpakai di semua elemen. Untuk override per-elemen, bisa via data dict (coming soon).

---

## 🆕 Icon System — Built-in Shapes + Emoji

Engine punya `_add_icon()` yang bikin lingkaran berwarna + icon di dalamnya.

### Shape Icons (built-in, tanpa dependency)

| Nama | Simbol | Contoh |
|------|--------|--------|
| `'check'` | ✓ | Centang hijau |
| `'x'` | ✗ | Silang merah |
| `'arrow'` | → | Panah |
| `'star'` | ★ | Bintang |
| `'circle'` | ● | Lingkaran |
| `'info'` | ℹ | Informasi |
| `'warning'` | ⚠ | Peringatan |
| `'question'` | ? | Pertanyaan |

### Cara Pakai di Card Grid

```python
# Shape icon (string name)
{"icon": "check", "title": "Selesai", "color": "TEAL", "items": [...]}

# Emoji (tetap support)
{"icon": "🔵", "title": "Info", "color": "BLUE", "items": [...]}
```

### Cara Pakai di Kode Engine

```python
# Shape icon in colored circle
engine._add_icon(slide, x, y, 0.42, icon="warning", fill=engine.C.WARM)

# Emoji
engine._add_icon(slide, x, y, 0.42, icon="⚠️", fill=engine.C.WARM)
```

---

## 🆕 Design Guidance — Bikin Slide yang Tidak Membosankan

Berdasarkan praktik dari Anthropic PPTX Skill + 60-30-10 rule.

### Sebelum Mulai

1. **Pilih palette yang spesifik untuk TOPIK** — bukan biru generik. Kalau warna bisa dipindah ke PPT topic lain dan masih cocok, warnanya belum spesifik.
2. **Dominance > equality** — 60-30-10. Satu warna dominan, 1-2 pendukung, 1 aksen tajam.
3. **Dark/light contrast** — Dark bg untuk cover+closing, light untuk konten ("sandwich").
4. **Commit ke 1 visual motif** — Pilih satu elemen khas dan ulangi: icon dalam lingkaran, border tebal di satu sisi, shape dekoratif. Bawa ke semua slide.

### Variasi Layout

Jangan ulang layout yang sama. Bergantian antar slide:

| Layout | Cocok Untuk | Archetype |
|--------|-------------|-----------|
| **Full card grid** | 4-8 item setara | `card_grid` |
| **Two column** | Perbandingan 2 sisi | `two_col` |
| **Number callout** | Angka-angka penting | `callout` |
| **Flow horizontal** | Proses / tahapan | `flow` |
| **Table** | Data terstruktur | `table` |
| **Content polos** | Intro / penutup bab | `content` |

### Visual Polish

- **Icons in circles** untuk setiap item (via `_add_icon`)
- **Gold bar** sebagai aksen header (otomatis oleh engine)
- **Bottom footer** konsisten (otomatis)
- **Action title** di navy header — kalimat kesimpulan, bukan topik

### Common Mistakes — Hindari!

| ❌ Salah | ✅ Benar |
|----------|----------|
| Action title = topik ("Bab I") | Action title = kesimpulan ("7 Definisi Kunci...") |
| Semua slide layout sama | Variasi card_grid, two_col, flow, callout |
| Body text di-center | Left-align body; center hanya judul |
| Terlalu rapat (< 0.3" gap) | 0.3" minimum antar elemen |
| Warna rendah kontras | Semua WCAG AA ≥4.5:1 (engine auto-verify) |
| Text-only slides | Tambah icon, shape, atau elemen visual |

---

## 🆕 QA Workflow — Verifikasi Visual

Gunakan `src/qa.py` untuk render PPTX ke gambar dan inspeksi.

### Quick Text Check

```bash
python3 src/qa.py --text-only output.pptx
```

Cek: placeholder text, urutan konten, typo.

### Render to Images

```bash
# Butuh LibreOffice + poppler-utils
sudo apt install -y libreoffice poppler-utils
pip install markitdown[pptx]

python3 src/qa.py output.pptx --render-only
# Output: output_dir/slide-01.jpg, slide-02.jpg, ...
```

### Full QA Loop

```bash
# 1. Generate PPT
python3 -c "from src.ppt_engine import *; ..."

# 2. Extract text + cek placeholder
python3 src/qa.py --text-only output.pptx

# 3. Render ke images
python3 src/qa.py --render-only output.pptx --dpi 150

# 4. Subagent inspect — via task() tool
python3 src/qa.py --inspect output.pptx
# → Output siap di-copy ke tool task()
```

### Cara Panggil Subagent (System Kita)

Gunakan tool **`task`** atau **`agentic_delegate`**:

```python
# Opsi 1: task() — untuk inspeksi gambar
task(
  description="Inspect PPTX slide images",
  prompt="""Visually inspect these slides...
Read and analyze these images:
1. /tmp/qa_output/slide-01.jpg
2. /tmp/qa_output/slide-02.jpg
...""",
  subagent_type="general"
)

# Opsi 2: agentic_delegate — untuk QA khusus
agentic_delegate(
  taskId="qa-pptx-001",
  description="Inspect slide images for visual issues",
  role="qa"
)
```

### Verification Loop (Wajib!)

1. Generate slides → Extract text → Render images
2. **Panggil subagent** via `task()` untuk inspeksi visual
3. Subagent laporin **semua masalah** (overlap, overflow, contrast, gaps)
4. Fix masalah
5. **Re-verify** slide yang berubah (render ulang → subagent lagi)
6. Ulang sampai subagent nemuin **no new issues**

> ⚠️ **WAJIB pakai subagent untuk visual QA** — matamu sudah capek lihat kode dan akan melihat apa yang kamu harapkan, bukan apa yang sebenarnya ada. Subagent punya *fresh eyes*.

---

## 🆕 Template Editing — Unpack / Edit / Pack

Gunakan `src/pptx_tools.py` untuk edit konten PPTX existing.

### Workflow

```bash
# 1. Unpack — extract ke direktori
python3 src/pptx_tools.py unpack template.pptx unpacked/

# 2. List slide content
python3 src/pptx_tools.py list template.pptx

# 3. Edit slide XML (gunakan Edit tool)
# unpacked/ppt/slides/slide1.xml
# Cari <a:t> untuk teks, ganti dengan konten baru

# 4. Pack — repack ke PPTX baru
python3 src/pptx_tools.py pack unpacked/ output.pptx
```

### Kapan Pakai Template Editing

| Situasi | Pakai |
|---------|-------|
| Buat PPT dari awal | `Engine(primary_color=...)` |
| Edit teks di PPT existing | `src/pptx_tools.py unpack → edit → pack` |
| Butuh template layout tertentu | Unpack template → edit konten → pack |
| Content extraction | `src/pptx_tools.py list` atau `src/qa.py --text-only` |

---

## ⚠️ Troubleshooting

| Masalah | Penyebab | Solusi |
|---------|----------|--------|
| PPTX tidak bisa dibuka | ZIP entry order | Pakai python-pptx (engine sudah aman) |
| Card overflow | Terlalu banyak items | Max 4 item per card |
| Teks terpotong | item_h kurang | Engine auto-calculate, tapi batasi teks |
| Warna salah | Format salah | Pakai `"#RRGGBB"` atau `"BLUE"` |
| Slide kurang/banyak | Hitung manual | Section divider = 1 slide |
| Action title tidak ada | Lupa | Setiap slide harus action title |

---

## 🔍 Contoh Hasil Analisis Dokumen → Slide

**Input**: Perwal Bekasi No 51/2024 tentang Pajak Reklame

**Output LLM**:
```
Cover: "PENGELOLAAN PAJAK REKLAME"
TOC: 11 bab
Bab I (Pasal 1): Section + card_grid 7 definisi
Bab II (Pasal 2-4): Section + card_grid 10 jenis + card_grid subjek
Bab III (Pasal 5): Section + callout masa pajak
Bab IV (Pasal 6-8): Section + two_col pendaftaran
Bab V (Pasal 9): Section + nsr_factors 7 faktor
Bab VI (Pasal 10): Section + callout rumus + 3 tabel + 2 card_grid
Bab VII (Pasal 11-14): Section + flow alur
Bab VIII (Pasal 15-20,29-33): Section + 2 × two_col
Bab IX (Pasal 22-26): Section + card_grid + flow
Bab X (Pasal 27-28,34-35): Section + card_grid
Bab XI (Pasal 36-37): Section + two_col
Closing: "TERIMA KASIH"
```

---

## ✅ Checklist untuk LLM

- [ ] Baca & pahami dokumen sumber
- [ ] Ekstrak bab, pasal, definisi, tabel, poin penting
- [ ] Buat action title tiap slide (kesimpulan, bukan topik)
- [ ] Pilih archetype sesuai konten
- [ ] Rotasi warna card agar variatif
- [ ] Hitung jumlah slide: cover + toc + (section+isi)×bab + closing
- [ ] Generate langsung via `python3 -c "..."` — tanpa file content
- [ ] Verifikasi file PPTX bisa dibuka
