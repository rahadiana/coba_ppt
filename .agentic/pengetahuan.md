# Pengetahuan — Proyek PPT Generator

## Arsitektur 3 File

```
📁 coba_ppt/
├── ppt_engine.py               ← Engine reusable (tidak perlu diubah)
├── content_*.py                ← Content data (ganti untuk PPT beda topik)
├── buat_ppt_generik.py         ← Entry point (panggil aja)
└── fix_pptx_zip.py             ← Utility fix ZIP PptxGenJS
```

### Cara Pakai
```bash
python3 buat_ppt_generik.py                          # content default
CONTENT_MODULE=content_my_topic python3 buat_ppt_generik.py  # content custom
```

---

## Root Cause: PptxGenJS Corrupt di PowerPoint

**Bukan** masalah XML/shape. Masalahnya **ZIP entry order**.

- ECMA-376 OPC §9.1.2.1: `[Content_Types].xml` WAJIB sebagai stream pertama dalam ZIP.
- PptxGenJS menaruh `[Content_Types].xml` di entry #19 → PowerPoint tolak.
- **Fix**: `python3 fix_pptx_zip.py <file.pptx>` — tulis ulang ZIP dengan entry order benar.
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

### CPI (Calibri characters-per-inch)
| Font Size | CPI |
|-----------|-----|
| 9pt | 14 |
| 11pt | 12 |
| 13pt | 10 |
| 20pt | 6.5 |
| 32pt | 4 |
| 40pt | 3 |

---

## 10 Slide Archetype

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

## Color Palette 60-30-10 — WCAG AA Verified ✅

| Peran | Warna | Hex | RGB | Kontras Putih |
|-------|-------|-----|-----|:-------------:|
| **60%** Dominan | Navy | `#0A1628` | `10,22,40` | 18.1:1 🏆 |
| **30%** Sekunder | Ice/White | `#F5F7FA` / `#FFFFFF` | `245,247,250` | — |
| **10%** Aksen | Gold | `#C8962E` | `200,150,46` | 2.7:1 (dark bg only) |

### Semantic Colors (WCAG AA ≥ 4.5:1 ✅)
| Nama | Hex | Kontras Putih | Penggunaan |
|------|-----|:------------:|------------|
| BLUE | `#2563EB` | 5.2:1 ✅ | Informasi, definisi |
| TEAL | **`#0B7C72`** | **5.1:1 ✅** | **Prosedur, data** |
| WARM | **`#A0522D`** | **5.6:1 ✅** | **Peringatan, faktor (sienna)** |
| RED | `#DC2626` | 4.8:1 ✅ | Sanksi, bahaya |

> ⚡ **Perubahan penting**: TEAL digelapkan `#0D9488→#0B7C72`, WARM diganti hue `#B8860B→#A0522D` (sienna, tidak clash dengan gold). TEXT_M `#6B7288→#8899B0` (6.2:1 on navy ✅). TEXT_L `#9CA3AF→#64748B` (4.8:1 on white ✅).

---

## Cara Buat PPT Baru (Topik Lain)

1. Copy `content_perwal_51_2024.py` → `content_topik_baru.py`
2. Ubah `PRESENTATION` dict (title, source, output)
3. Ubah `SLIDES` list — pakai type yang sesuai dari 10 archetype di atas
4. Jalankan: `CONTENT_MODULE=content_topik_baru python3 buat_ppt_generik.py`

### Content Dict Structure
```python
PRESENTATION = {
    "title": "Judul Presentasi",
    "source": "Sumber: ...",
    "output": "output.pptx",
}

SLIDES = [
    {"type": "cover", "data": {
        "pre_title": "...", "city": "...",
        "main_title": "...", "display_title": "...",
        "badge_text": "...",
    }},
    {"type": "card_grid", "data": {
        "title": "Action Title",
        "subtitle": "...",
        "cards": [
            {"icon": "📊", "title": "Judul Card",
             "color": "#2563EB", "items": ["item1", "item2"]},
        ],
    }},
    # ... slide lainnya
]
```

---

## Perwal Bekasi No 51/2024 — Content Spesifik

- **33 slide**: Cover + TOC + 11 Bab + Closing
- **Rumus**: `Pajak Reklame = Tarif (50%) × NSR`
- **NSR** = Nilai Kelas Jalan × Ukuran (m²) × Jumlah × Jangka Waktu
- **Ketentuan Khusus**:
  - Reklame indoor: NSR 50%
  - Ketinggian > 15m: tambahan 20%
  - Produk tembakau & miras: tambahan 50%
- **Kelas Jalan**: Khusus (Tol, Premium 1, Premium 2), I (Kendali Ketat), II (Kendali Sedang)
- **10 jenis reklame**: Papan, Videotron, Kain, Stiker, Selebaran, Berjalan, Udara, Apung, Film/Slide, Peragaan
- **8 pengecualian**: Internet/TV/radio, label produk, nama usaha ≤1m², pemerintah, ibadah, sosial, politik, olahraga KONI
