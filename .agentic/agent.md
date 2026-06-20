# 🎯 PPT Generator — Agent Navigation

> **Skill**: Membuat presentasi PowerPoint profesional dengan layout presisi, action titles (McKinsey style), visual hierarchy, bebas overlap, kompatibel Microsoft PowerPoint.

---

## 📋 Kapan Pakai Skill Ini

- Butuh PPT dari data/regulasi/dokumen dalam waktu singkat
- Butuh **action titles** (judul = kesimpulan, bukan topik)
- Butuh **layout presisi** tanpa overlap
- Butuh **PowerPoint compatible** (bisa buka di PowerPoint)
- Butuh **content terpisah dari engine** — ganti konten tanpa edit kode

---

## 🏗️ Arsitektur

```
📁 project/
├── ppt_engine.py               ← ENGINE REUSABLE — never edit
├── content_*.py                ← CONTENT — edit untuk PPT baru
├── buat_ppt_generik.py         ← ENTRY POINT — panggil aja
├── fix_pptx_zip.py             ← UTILITY — fix ZIP order PptxGenJS
├── .agentic/
│   ├── agent.md                ← file ini
│   └── pengetahuan.md          ← dokumentasi lengkap
└── output.pptx                 ← hasil generate
```

### Aliran Data
```
content_xxx.py ──→ buat_ppt_generik.py ──→ ppt_engine.py ──→ output.pptx
                     (entry point)            (engine)
```

---

## 🚀 Cara Generate PPT

### Dengan content default (Perwal Bekasi):
```bash
python3 buat_ppt_generik.py
```

### Dengan content kustom:
```bash
CONTENT_MODULE=content_topik_baru python3 buat_ppt_generik.py
```

### Dari kode:
```python
from ppt_engine import Engine
from content_topik_baru import PRESENTATION, SLIDES

engine = Engine()
engine.build(SLIDES,
             source_text=PRESENTATION['source'],
             output_path=PRESENTATION['output'])
```

---

## 📝 Cara Buat Content Baru

1. **Copy** `content_perwal_51_2024.py` → `content_topik_baru.py`
2. **Ubah** `PRESENTATION` dict:
   ```python
   PRESENTATION = {
       "title": "Judul Presentasi",
       "source": "Sumber: ...",
       "output": "output.pptx",
   }
   ```
3. **Ubah** `SLIDES` — list of slide dicts:
   ```python
   SLIDES = [
       {"type": "cover", "data": {...}},
       {"type": "section", "data": {...}},
       {"type": "card_grid", "data": {...}},
       # ...
   ]
   ```
4. **Jalankan**: `CONTENT_MODULE=content_topik_baru python3 buat_ppt_generik.py`

---

## 🧩 10 Slide Archetype + Data Format

### 1. `cover` — Halaman Sampul
```python
{"type": "cover", "data": {
    "pre_title": "BERITA DAERAH",        # atas, gold, 12pt
    "city": "KOTA BEKASI",               # 14pt white
    "main_title": "PERATURAN...",         # 20pt white bold
    "main_subtitle": "NOMOR ...",        # 15pt gold
    "display_title": "TENTANG\n...",     # 40pt white bold — judul utama
    "badge_text": "Kota Bekasi · 2024",  # badge navy_d, 12pt ice
    "badge_subtext": "Berlaku sejak...", # 10pt text_l
}}
```

### 2. `toc` — Daftar Isi
```python
{"type": "toc", "data": {
    "title": "Action Title Daftar Isi",
    "subtitle": "Subtitle",
    "cols": 2,                           # jumlah kolom
    "items": [
        {"num": "1", "label": "Bab 1", "color": "#2563EB"},
        {"num": "2", "label": "Bab 2", "color": "#0D9488"},
        # ...
    ],
}}
```

### 3. `section` — Section Divider (full-bleed navy)
```python
{"type": "section", "data": {
    "title": "BAB I\nKETENTUAN UMUM",   # 34pt bold white
    "subtitle": "Pasal 1",              # 11pt text_l
    "action_text": "7 Definisi Kunci",  # 14pt gold bold (action title)
}}
```

### 4. `content` — Standar Header + Footer
```python
{"type": "content", "data": {
    "title": "Action Title",
    "subtitle": "Subtitle",
}}
```

### 5. `card_grid` — Multi-row Cards
```python
{"type": "card_grid", "data": {
    "title": "Action Title",
    "subtitle": "Pasal 1",
    "cols": 0,                          # 0 = auto (2/3/4 cols)
    "cards": [
        {"icon": "🏛️", "title": "Judul Card",
         "color": "#2563EB",
         "items": ["Item 1", "Item 2"]},  # → bullet "• Item 1"
        {"icon": "📊", "title": "",        # tanpa icon: hapus field icon
         "color": "#0D9488",
         "items": ["Item A"]},
    ],
}}
```

### 6. `two_col` — 2 Column Cards
```python
{"type": "two_col", "data": {
    "title": "Action Title",
    "subtitle": "Pasal 6–8",
    "left": {
        "color": "#2563EB",
        "lines": [
            "$JUDUL SECTION",            # $ → highlight 14pt bold
            "",                           # baris kosong = spacer
            "• bullet otomatis",          # teks biasa → "• teks"
            "Line tanpa bullet",
        ],
    },
    "right": {
        "color": "#0D9488",
        "lines": ["$SECTION 2", "", "item 1", "item 2"],
    },
}}
```

### 7. `callout` — Big Number Cards
```python
{"type": "callout", "data": {
    "title": "Action Title",
    "subtitle": "Pasal 5",
    "callouts": [
        {"number": "12", "label": "Bulan\n(Permanen)", "color": "#2563EB"},
        {"number": "30", "label": "Hari\n(Insidentil)", "color": "#0D9488"},
    ],
    "note": "• Note line 1\n• Note line 2",   # optional
}}
```

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
    "note": "• Jatuh tempo: 1 bulan\n• Bunga 1%/bln",
}}
```

### 9. `table` — Native Table
```python
{"type": "table", "data": {
    "title": "Tabel NSR",
    "subtitle": "Pasal 10",
    "headers": ["Kelas Jalan", "Zona", "NSR"],      # navy header
    "rows": [
        ["Kelas Khusus", "Tol", "23.575"],            # zebra stripe
        ["Kelas I", "Ketat", "13.225"],
    ],
}}
```

### 10. `nsr_factors` — Faktor NSR (custom)
```python
{"type": "nsr_factors", "data": {
    "title": "7 Faktor NSR",
    "subtitle": "Pasal 9",
    "factors": [
        {"num": "1", "label": "Jenis Reklame", "color": "#2563EB"},
        {"num": "2", "label": "Bahan", "color": "#0D9488"},
        # ... max 7 item, 4+3 grid
    ],
    "classification_note": "🏛️ Kelas Jalan Khusus\n🚗 Kelas Jalan I",
}}
```

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

Gunakan **string hex** atau **nama warna** di content:

| Nama | Hex | Penggunaan |
|------|-----|------------|
| `#0A1628` | Navy | 60% dominan — background, header |
| `#F5F7FA` | Ice | 30% sekunder — background slide |
| `#FFFFFF` | White | card background |
| `#C8962E` | Gold | 10% aksen — bar, highlight |
| `#2563EB` | Blue | info, definisi |
| `#0D9488` | Teal | prosedur, data |
| `#B8860B` | Warm | peringatan, faktor |
| `#DC2626` | Red | sanksi, bahaya |

Bisa pakai: `"#2563EB"` atau `"BLUE"` (case insensitive).

---

## 📐 Layout Rules (jangan dilanggar)

### Zona Slide (16:9 = 13.333" × 7.5")
```
┌─ 0.035" gold_bar ────────────────────────────┐
│ 0.9" navy header                              │ ← HEADER (0.94")
│ 0.21" gap                                     │
├───────────────────────────────────────────────┤
│                                               │
│    CONTENT AREA (12.133" × 5.85")             │
│    margin_h = 0.6"                            │
│    cx=0.6, cy=1.15                            │
│                                               │
├───────────────────────────────────────────────┤
│ 0.03" navy bar + source + page #             │ ← FOOTER (0.50")
└───────────────────────────────────────────────┘
```

### Rumus Kunci
```
col_width(N)      = (12.133 - (N-1) × 0.3) / N
row_height(M)     = (5.85 - (M-1) × 0.3) / M
grid_pos(col,row) = 0.6 + col × (col_w + 0.3), 1.15 + row × (row_h + 0.3)
text_height()     = ceil(len / (cpi × box_w)) × pt × 1.2 / 72
safe_item_height  = max(text_height + 0.05, 0.25)
```

### CPI Calibri
| pt | cpi (chars/inch) |
|----|------------------|
| 9  | 14 |
| 11 | 12 |
| 13 | 10 |
| 20 | 6.5 |
| 32 | 4 |
| 40 | 3 |

---

## ⚠️ Troubleshooting

### PPTX corrupt / tidak bisa dibuka PowerPoint
- **Penyebab**: ZIP entry order salah (PptxGenJS)
- **Fix**: `python3 fix_pptx_zip.py file.pptx`
- **Prevent**: pakai `python-pptx`, bukan PptxGenJS

### Card text overflow / overlap
- **Penyebab**: item_h terlalu kecil untuk teks
- **Fix**: engine sudah auto-calculate `safe_item_height()`
- Jika masih overflow: kurangi jumlah items atau perpendek teks

### Warna tidak sesuai
- Pastikan pakai format hex `#RRGGBB` atau nama warna `BLUE`
- Case insensitive untuk nama warna

### Slide count tidak sesuai
- Hitung manual jumlah dict di `SLIDES`
- Section divider juga dihitung sebagai slide

---

## 📂 Referensi File

| File | Fungsi | Wajib Ada? |
|------|--------|------------|
| `ppt_engine.py` | Engine kelas LayoutFrame + archetype builders | ✅ Ya |
| `content_*.py` | Data konten spesifik | ✅ Ya (min 1) |
| `buat_ppt_generik.py` | Entry point | ✅ Ya |
| `fix_pptx_zip.py` | Utility fix ZIP order | 🔧 Optional |
| `.agentic/pengetahuan.md` | Dokumentasi lengkap | 📖 Optional |

---

## ✅ Checklist Sebelum Push

- [ ] PPT bisa dibuka di PowerPoint
- [ ] Tidak ada overlap (cek visual)
- [ ] Action title di setiap slide (judul = kesimpulan)
- [ ] Page number berurutan
- [ ] Source citation di footer
- [ ] Warna konsisten 60-30-10
- [ ] Content terpisah dari engine (content_*.py)
