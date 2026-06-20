#!/usr/bin/env python3
"""
content_trigonometri.py — Konten Lengkap PPT Trigonometri
========================================================
Materi lengkap: Pengertian, Rumus Dasar, Kebalikan, Sudut Istimewa,
Kuadran, Identitas, Aturan Sinus/Cosinus, Luas Segitiga, Aplikasi.
"""

PRESENTATION = {
    "title": "Kupas Tuntas Trigonometri",
    "source": "Materi Trigonometri Lengkap — Matematika SMA",
    "output": "Trigonometri_Lengkap.pptx"
}

SLIDES = [
    # ═══════════════════════════════════════════════════════════
    # SLIDE 1: COVER
    # ═══════════════════════════════════════════════════════════
    {
        "type": "cover",
        "data": {
            "pre_title": "MATEMATIKA SMA",
            "city": "MATERI LENGKAP",
            "main_title": "KUPAS TUNTAS TRIGONOMETRI",
            "main_subtitle": "Hubungan Sudut dan Segitiga — Dari Dasar hingga Aplikasi",
            "display_title": "TRIGONOMETRI",
            "badge_text": "Pengertian  •  Rumus Utama  •  Kuadran  •  Aturan Sinus/Cosinus  •  Aplikasi",
            "badge_subtext": "12 Slide  •  Materi Lengkap  •  Siap Presentasi"
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 2: DAFTAR ISI
    # ═══════════════════════════════════════════════════════════
    {
        "type": "toc",
        "data": {
            "title": "Daftar Isi",
            "subtitle": "Struktur Materi Trigonometri",
            "items": [
                {"num": "1", "label": "Apa itu Trigonometri?", "color": "BLUE"},
                {"num": "2", "label": "Perbandingan Trigonometri Dasar", "color": "TEAL"},
                {"num": "3", "label": "Kebalikan Trigonometri", "color": "BLUE"},
                {"num": "4", "label": "Tabel Sudut Istimewa", "color": "TEAL"},
                {"num": "5", "label": "Pembagian Kuadran", "color": "BLUE"},
                {"num": "6", "label": "Identitas Trigonometri", "color": "TEAL"},
                {"num": "7", "label": "Aturan Sinus & Cosinus", "color": "BLUE"},
                {"num": "8", "label": "Luas Segitiga", "color": "TEAL"},
                {"num": "9", "label": "Aplikasi Dunia Nyata", "color": "BLUE"},
                {"num": "10", "label": "Kesimpulan & Penutup", "color": "TEAL"}
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 3: APA ITU TRIGONOMETRI
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "Apa itu Trigonometri?",
            "subtitle": "Asal kata, definisi, dan inti bahasan",
            "cards": [
                {
                    "icon": "📖",
                    "title": "Asal Kata",
                    "color": "BLUE",
                    "items": [
                        "Bahasa Yunani: trigonon (tiga sudut)",
                        "Bahasa Yunani: metron (mengukur)",
                        "Arti literal: mengukur tiga sudut"
                    ]
                },
                {
                    "icon": "📐",
                    "title": "Definisi",
                    "color": "TEAL",
                    "items": [
                        "Cabang matematika yang mempelajari",
                        "Hubungan antara panjang sisi",
                        "Dan besar sudut pada segitiga"
                    ]
                },
                {
                    "icon": "🔺",
                    "title": "Inti Bahasan",
                    "color": "BLUE",
                    "items": [
                        "Fokus pada segitiga siku-siku",
                        "Perbandingan panjang sisi",
                        "Hubungan dengan besar sudut"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 4: PERBANDINGAN TRIGONOMETRI DASAR
    # ═══════════════════════════════════════════════════════════
    {
        "type": "two_col",
        "data": {
            "title": "Perbandingan Trigonometri Dasar",
            "subtitle": "Rumus utama pada segitiga siku-siku dengan sudut θ (theta)",
            "left": {
                "color": "BLUE",
                "lines": [
                    "$Sisi pada Segitiga:",
                    "Depan (de) — sisi di depan sudut θ",
                    "Samping (sa) — sisi di samping sudut θ",
                    "Miring (mi) — sisi terpanjang, depan sudut siku-siku"
                ]
            },
            "right": {
                "color": "TEAL",
                "lines": [
                    "$Rumus Utama (Jembatan Keledai):",
                    "sin θ = Depan / Miring → Sin-De-Mi",
                    "cos θ = Samping / Miring → Cos-Sam-Mi",
                    "tan θ = Depan / Samping → Tan-De-Sam"
                ]
            }
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 5: KEBALIKAN TRIGONOMETRI
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "Kebalikan Trigonometri Dasar",
            "subtitle": "Tiga fungsi kebalikan dari sin, cos, tan",
            "cards": [
                {
                    "icon": "🔄",
                    "title": "Cosecan (csc)",
                    "color": "BLUE",
                    "items": [
                        "Kebalikan dari Sinus",
                        "csc θ = 1 / sin θ",
                        "csc θ = Miring / Depan"
                    ]
                },
                {
                    "icon": "🔄",
                    "title": "Secan (sec)",
                    "color": "TEAL",
                    "items": [
                        "Kebalikan dari Cosinus",
                        "sec θ = 1 / cos θ",
                        "sec θ = Miring / Samping"
                    ]
                },
                {
                    "icon": "🔄",
                    "title": "Cotangen (cot)",
                    "color": "BLUE",
                    "items": [
                        "Kebalikan dari Tangen",
                        "cot θ = 1 / tan θ",
                        "cot θ = Samping / Depan"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 6: TABEL SUDUT ISTIMEWA
    # ═══════════════════════════════════════════════════════════
    {
        "type": "table",
        "data": {
            "title": "Tabel Sudut Istimewa",
            "subtitle": "Wajib Hafal! Sudut 0°, 30°, 45°, 60°, 90°",
            "headers": ["Fungsi", "0°", "30°", "45°", "60°", "90°"],
            "rows": [
                ["sin", "0", "1/2", "½√2", "½√3", "1"],
                ["cos", "1", "½√3", "½√2", "1/2", "0"],
                ["tan", "0", "⅓√3", "1", "√3", "∞ (Tak hingga)"],
                ["csc", "∞", "2", "√2", "2/√3", "1"],
                ["sec", "1", "2/√3", "√2", "2", "∞"],
                ["cot", "∞", "√3", "1", "⅓√3", "0"]
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 7: PEMBAGIAN KUADRAN
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "Pembagian Kuadran",
            "subtitle": "Sudut lebih dari 90° — Nilai positif/negatif fungsi trigonometri",
            "cards": [
                {
                    "icon": "I",
                    "title": "Kuadran I (0° - 90°)",
                    "color": "BLUE",
                    "items": [
                        "Semua fungsi bernilai POSITIF (+)",
                        "sin(+), cos(+), tan(+)",
                        "csc(+), sec(+), cot(+)"
                    ]
                },
                {
                    "icon": "II",
                    "title": "Kuadran II (90° - 180°)",
                    "color": "TEAL",
                    "items": [
                        "Hanya Sinus & Cosecan POSITIF",
                        "sin(+), csc(+)",
                        "cos(-), tan(-), sec(-), cot(-)"
                    ]
                },
                {
                    "icon": "III",
                    "title": "Kuadran III (180° - 270°)",
                    "color": "BLUE",
                    "items": [
                        "Hanya Tangen & Cotangen POSITIF",
                        "tan(+), cot(+)",
                        "sin(-), cos(-), csc(-), sec(-)"
                    ]
                },
                {
                    "icon": "IV",
                    "title": "Kuadran IV (270° - 360°)",
                    "color": "TEAL",
                    "items": [
                        "Hanya Cosinus & Secan POSITIF",
                        "cos(+), sec(+)",
                        "sin(-), tan(-), csc(-), cot(-)"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 8: IDENTITAS TRIGONOMETRI
    # ═══════════════════════════════════════════════════════════
    {
        "type": "two_col",
        "data": {
            "title": "Identitas Trigonometri Dasar",
            "subtitle": "Rumus hubungan antar-fungsi untuk menyederhanakan persamaan",
            "left": {
                "color": "BLUE",
                "lines": [
                    "$Hubungan Tangen:",
                    "tan θ = sin θ / cos θ",
                    "",
                    "$Identitas Pythagoras:",
                    "sin²θ + cos²θ = 1",
                    "1 + tan²θ = sec²θ"
                ]
            },
            "right": {
                "color": "TEAL",
                "lines": [
                    "$Identitas Pythagoras (lanjutan):",
                    "1 + cot²θ = csc²θ",
                    "",
                    "$Hubungan Lain:",
                    "sin(90° - θ) = cos θ",
                    "cos(90° - θ) = sin θ"
                ]
            }
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 9: ATURAN SINUS & COSINUS
    # ═══════════════════════════════════════════════════════════
    {
        "type": "two_col",
        "data": {
            "title": "Aturan Sinus & Cosinus",
            "subtitle": "Untuk menyelesaikan segitiga yang bukan siku-siku",
            "left": {
                "color": "BLUE",
                "lines": [
                    "$Aturan Sinus:",
                    "a/sin A = b/sin B = c/sin C",
                    "",
                    "$Kegunaan:",
                    "Mencari sisi jika sudut diketahui",
                    "Mencari sudut jika sisi diketahui",
                    "Cocok untuk kasus AAS atau ASA"
                ]
            },
            "right": {
                "color": "TEAL",
                "lines": [
                    "$Aturan Cosinus:",
                    "c² = a² + b² - 2ab·cos C",
                    "",
                    "$Kegunaan:",
                    "Mencari sisi ketiga jika 2 sisi + sudut diketahui",
                    "Mencari sudut jika 3 sisi diketahui",
                    "Cocok untuk kasus SAS atau SSS"
                ]
            }
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 10: LUAS SEGITIGA
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "Luas Segitiga Menggunakan Trigonometri",
            "subtitle": "Rumus alternatif untuk menghitung luas segitiga",
            "cards": [
                {
                    "icon": "📐",
                    "title": "Rumus Utama",
                    "color": "BLUE",
                    "items": [
                        "L = ½ × a × b × sin C",
                        "L = ½ × a × c × sin B",
                        "L = ½ × b × c × sin A"
                    ]
                },
                {
                    "icon": "💡",
                    "title": "Kapan Digunakan?",
                    "color": "TEAL",
                    "items": [
                        "Ketika tinggi segitiga tidak diketahui",
                        "Ketika diketahui 2 sisi + sudut di antaranya",
                        "Alternatif dari L = ½ × alas × tinggi"
                    ]
                },
                {
                    "icon": "📝",
                    "title": "Contoh Soal",
                    "color": "BLUE",
                    "items": [
                        "Diketahui: a=8 cm, b=6 cm, C=60°",
                        "L = ½ × 8 × 6 × sin 60°",
                        "L = 24 × ½√3 = 12√3 cm²"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 11: APLIKASI DUNIA NYATA
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "Penerapan di Dunia Nyata",
            "subtitle": "Trigonometri bukan cuma teori, tapi dipakai di berbagai bidang",
            "cards": [
                {
                    "icon": "🏗️",
                    "title": "Arsitektur & Sipil",
                    "color": "BLUE",
                    "items": [
                        "Menghitung tinggi gedung/jembatan",
                        "Menggunakan klinometer & sudut elevasi",
                        "Tanpa harus memanjat langsung"
                    ]
                },
                {
                    "icon": "🧭",
                    "title": "Navigasi",
                    "color": "TEAL",
                    "items": [
                        "Menentukan posisi kapal/pesawat",
                        "Berdasarkan koordinat satelit/radar",
                        "Penggunaan dalam GPS modern"
                    ]
                },
                {
                    "icon": "🎮",
                    "title": "Pembuatan Game",
                    "color": "BLUE",
                    "items": [
                        "Grafis 3D dan pergerakan karakter",
                        "Sudut pandang kamera (field of view)",
                        "Fisika game dan animasi"
                    ]
                },
                {
                    "icon": "📡",
                    "title": "Sains Gelombang",
                    "color": "TEAL",
                    "items": [
                        "Analisis gelombang suara & cahaya",
                        "Sinyal radio dan telekomunikasi",
                        "Grafik fungsi trigonometri"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 12: KESIMPULAN & PENUTUP
    # ═══════════════════════════════════════════════════════════
    {
        "type": "closing",
        "data": {
            "pre_title": "KESIMPULAN",
            "main_title": "Terima Kasih!",
            "subtitle": "Trigonometri adalah alat matematika yang kuat untuk mengukur jarak dan sudut yang sulit dijangkau secara langsung.\n\nKunci menguasai trigonometri:\n• Paham konsep segitiga siku-siku\n• Hafal sudut istimewa\n• Tahu aturan kuadran\n• Kuasai aturan sinus & cosinus",
            "source": "Ada Pertanyaan? Silakan diskusi!"
        }
    }
]


if __name__ == "__main__":
    from ppt_engine import Engine
    
    print("=" * 60)
    print("📐 PPT GENERATOR — TRIGONOMETRI LENGKAP")
    print("=" * 60)
    print()
    
    pres = PRESENTATION
    print(f"📋 Judul: {pres['title']}")
    print(f"📄 Slide: {len(SLIDES)} definisi")
    print(f"📁 Output: {pres['output']}")
    print()
    
    engine = Engine()
    prs = engine.build(SLIDES, source_text=pres['source'], output_path=pres['output'])
    
    import os
    print(f"✅ BERHASIL: {pres['output']}")
    print(f"   {len(prs.slides)} slide, {os.path.getsize(pres['output']):,} bytes")
    
    # Tampilkan komposisi slide
    types = {}
    for s in SLIDES:
        t = s.get("type", "?")
        types[t] = types.get(t, 0) + 1
    print(f"\n📊 Komposisi slide:")
    for t, n in sorted(types.items()):
        print(f"   • {t}: {n}")
    
    abs_path = os.path.abspath(pres['output'])
    print(f"\n📍 Path: {abs_path}")
