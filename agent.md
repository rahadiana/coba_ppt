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
from ppt_engine import Engine

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
from ppt_engine import Engine
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

## 🎨 Color System (60-30-10 Rule)

| Nama | Hex | Peran | Penggunaan |
|------|-----|-------|------------|
| Navy | `#0A1628` | **60%** dominan | Background, header, section divider |
| Ice | `#F5F7FA` | **30%** sekunder | Background content slide |
| White | `#FFFFFF` | **30%** sekunder | Card background, teks header |
| Gold | `#C8962E` | **10%** aksen | Bar, highlight, arrow |
| Blue | `#2563EB` | semantic | Info, definisi, card 1 |
| Teal | `#0D9488` | semantic | Prosedur, data, card 2 |
| Warm | `#B8860B` | semantic | Peringatan, faktor, card 3 |
| Red | `#DC2626` | semantic | Sanksi, bahaya, card 4 |

> **Pakai di content**: `"#2563EB"` atau `"BLUE"` (case insensitive)

### Rotasi Warna untuk Card
```python
# Untuk card_grid dengan banyak card, rotasi warna:
colors = ["#2563EB", "#0D9488", "#B8860B", "#1B3A6B"]
for i, card in enumerate(cards):
    card["color"] = colors[i % len(colors)]
```

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
from ppt_engine import Engine
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
from ppt_engine import Engine
SLIDES = [...]  # konten
Engine().build(SLIDES, source_text='...', output_path='output.pptx')
```
```bash
python3 _gen.py && rm _gen.py
```

### Opsi 3 — Lewat buat_ppt_generik.py (dengan file content)
Hanya jika konten akan dipakai berulang:
```bash
CONTENT_MODULE=content_xxx python3 buat_ppt_generik.py
```

---

## ✅ Referensi File

| File | Fungsi | Wajib? |
|------|--------|--------|
| `ppt_engine.py` | Engine — LayoutFrame + 11 archetype builders | ✅ Ya |
| `buat_ppt_generik.py` | Entry point untuk content file | 🔧 Optional |
| `fix_pptx_zip.py` | Utility fix ZIP order PPTX corrupt | 🔧 Optional |

> **LLM tidak perlu** menyentuh `ppt_engine.py`. Cukup baca `agent.md` ini, extract konten dari sumber, lalu panggil engine langsung.

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
python3 -c "from ppt_engine import Engine; print('✅ Engine siap')"
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
from ppt_engine import Engine, LayoutFrame, Colors
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
| `ModuleNotFoundError: ppt_engine` | engine file tidak ada | Cek `ls ppt_engine.py` |
| `SyntaxError` | Python < 3.8 | Upgrade Python |

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
