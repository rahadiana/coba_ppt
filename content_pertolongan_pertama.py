#!/usr/bin/env python3
"""
content_pertolongan_pertama.py — PPT Panduan Pertolongan Pertama
===============================================================
Materi lengkap: Prinsip 3A, DRSABCD, Do's and Don'ts, Golden Hour.
Warna medis (merah-putih-biru) dengan emoji dan clipart banyak.
"""

from ppt_engine import Engine, Colors, PaletteGenerator
from pptx.dml.color import RGBColor

# ═══════════════════════════════════════════════════════════
# CUSTOM PALETTE — MEDICAL / RED CROSS THEME
# ═══════════════════════════════════════════════════════════

class MedicalColors:
    """
    Warna tema medis: Merah, Putih, Biru muda.
    Terinspirasi dari Palang Merah / emblem kesehatan.
    """
    # ── 60% Dominant ──
    NAVY     = RGBColor(0x1A, 0x23, 0x7E)   # Biru tua (trust, professional)
    NAVY_L   = RGBColor(0x28, 0x35, 0x93)   # Biru lebih terang
    NAVY_D   = RGBColor(0x0D, 0x14, 0x5C)   # Biru lebih gelap
    NAVY_M   = RGBColor(0x1E, 0x2A, 0x8A)   # Biru medium
    
    # ── 30% Secondary ──
    WHITE    = RGBColor(0xFF, 0xFF, 0xFF)   # Putih bersih
    OFF_W    = RGBColor(0xF8, 0xF9, 0xFC)   # Off-white
    ICE      = RGBColor(0xE8, 0xEA, 0xF6)   # Biru muda sangat terang
    ICE_D    = RGBColor(0xD0, 0xD4, 0xE8)   # Biru muda lebih gelap
    
    # ── 10% Accent ──
    RED      = RGBColor(0xDC, 0x26, 0x26)   # Merah Palang Merah
    RED_L    = RGBColor(0xEF, 0x44, 0x44)   # Merah lebih terang
    RED_D    = RGBColor(0xB9, 0x1C, 0x1C)   # Merah lebih gelap
    
    # ── Accent (Gold for headers) ──
    GOLD     = RGBColor(0xDC, 0x26, 0x26)   # Merah untuk aksen (ganti gold)
    GOLD_L   = RGBColor(0xEF, 0x44, 0x44)   # Merah terang
    
    # ── Semantic ──
    TEXT_D   = RGBColor(0x1A, 0x1A, 0x2E)   # Teks utama
    TEXT_M   = RGBColor(0x6B, 0x72, 0x80)   # Teks sekunder
    TEXT_L   = RGBColor(0x9C, 0xA3, 0xAF)   # Teks ringan
    GREEN    = RGBColor(0x16, 0xA3, 0x4A)   # Hijau (Boleh/Do's)
    GREEN_L  = RGBColor(0x22, 0xC5, 0x5E)   # Hijau terang
    AMBER    = RGBColor(0xF5, 0x9E, 0x0B)   # Kuning (Peringatan)
    AMBER_D  = RGBColor(0xD9, 0x77, 0x06)   # Kuning gelap
    TEAL     = RGBColor(0x08, 0x91, 0xB2)   # Teal (info)
    BLUE     = RGBColor(0x25, 0x63, 0xEB)   # Biru info


PRESENTATION = {
    "title": "Panduan Lengkap Pertolongan Pertama",
    "source": "Panduan Pertolongan Pertama pada Korban Kecelakaan",
    "output": "Pertolongan_Pertama_Lengkap.pptx"
}


SLIDES = [
    # ═══════════════════════════════════════════════════════════
    # SLIDE 1: COVER
    # ═══════════════════════════════════════════════════════════
    {
        "type": "cover",
        "data": {
            "pre_title": "🚨 PANDUAN KESEHATAN & KESELAMATAN",
            "city": "EDUKASI MASYARAKAT",
            "main_title": "PANDUAN LENGKAP\nPERTOLONGAN PERTAMA",
            "main_subtitle": "🏥 Penanganan Korban Kecelakaan — Berdasarkan Pedoman BTLS (Basic Trauma Life Support)",
            "display_title": "FIRST\nAID",
            "badge_text": "🚑 Prinsip 3A  •  🩺 DRSABCD  •  ❤️ Golden Hour  •  🩹 Do's & Don'ts",
            "badge_subtext": "14 Slide  •  Materi Lengkap  •  Siap Edukasi"
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 2: DAFTAR ISI
    # ═══════════════════════════════════════════════════════════
    {
        "type": "toc",
        "data": {
            "title": "📋 Daftar Isi",
            "subtitle": "Struktur Materi Pertolongan Pertama",
            "items": [
                {"num": "1", "label": "🚨 Mengapa Pertolongan Pertama Penting?", "color": "RED"},
                {"num": "2", "label": "🛡️ Prinsip 3A: Amankan Diri, Lingkungan, Korban", "color": "BLUE"},
                {"num": "3", "label": "📞 Hubungi Bantuan Medis (119/112)", "color": "RED"},
                {"num": "4", "label": "👀 Periksa Kesadaran Korban", "color": "BLUE"},
                {"num": "5", "label": "⚠️ Jangan Pindahkan Korban Sembarangan!", "color": "RED"},
                {"num": "6", "label": "🩸 Tangani Perdarahan Hebat", "color": "BLUE"},
                {"num": "7", "label": "✅❌ Do's and Don'ts", "color": "RED"},
                {"num": "8", "label": "⏰ The Golden Hour", "color": "BLUE"}
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 3: MENGAPA PERTOLONGAN PERTAMA PENTING
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "🚨 Mengapa Pertolongan Pertama Penting?",
            "subtitle": "Kecelakaan dapat terjadi kapan saja — penanganan cepat menyelamatkan nyawa",
            "cards": [
                {
                    "icon": "⏱️",
                    "title": "Waktu Sangat Kritis",
                    "color": "RED",
                    "items": [
                        "Setiap detik sangat berharga",
                        "\"The Golden Hour\" = 60 menit pertama",
                        "Penanganan cepat = peluang hidup lebih besar"
                    ]
                },
                {
                    "icon": "🩺",
                    "title": "Mencegah Cedera Lebih Parah",
                    "color": "BLUE",
                    "items": [
                        "Cedera tulang belakang permanen",
                        "Perdarahan tak terkendali",
                        "Gangguan pernapasan"
                    ]
                },
                {
                    "icon": "🚑",
                    "title": "Menjembatani ke Profesional",
                    "color": "RED",
                    "items": [
                        "Pertolongan awam = pertolongan pertama",
                        "Sambil menunggu ambulans tiba",
                        "Anda bisa jadi pahlawan nyawa!"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 4: PRINSIP 3A — AMAN DIRI
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "🛡️ Prinsip 3A — Amankan Keadaan",
            "subtitle": "Sebelum menolong, PASTIKAN Anda aman terlebih dahulu!",
            "cards": [
                {
                    "icon": "🧑‍⚕️",
                    "title": "A: AMAN DIRI",
                    "color": "RED",
                    "items": [
                        "⚠️ Jangan menjadi korban selanjutnya!",
                        "🧤 Pakai APD jika ada (sarung tangan medis)",
                        "🚗 Pastikan posisi aman dari lalu lintas",
                        "👁️ Waspada terhadap bahaya sekitar"
                    ]
                },
                {
                    "icon": "🛣️",
                    "title": "A: AMAN LINGKUNGAN",
                    "color": "BLUE",
                    "items": [
                        "🔑 Matikan mesin kendaraan yang terlibat",
                        "🔺 Pasang segitiga pengaman",
                        "🚦 Minta orang lain atur lalu lintas",
                        "🔥 Jauhkan api atau sumber percikan"
                    ]
                },
                {
                    "icon": "🤕",
                    "title": "A: AMAN KORBAN",
                    "color": "RED",
                    "items": [
                        "🚷 Jauhkan korban dari bahaya langsung",
                        "🔥 Risiko kebakaran atau ledakan",
                        "🚗 Risiko terlindas kendaraan lain",
                        "⚡ Risiko korsleting listrik"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 5: HUBUNGI BANTUAN MEDIS
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "📞 Segera Hubungi Bantuan Medis Darurat",
            "subtitle": "Langkah terpenting: Panggil tenaga ahli segera!",
            "cards": [
                {
                    "icon": "1️⃣",
                    "title": "Nomor 119",
                    "color": "RED",
                    "items": [
                        "🚑 Ambulans / Gawat Darurat Medis",
                        "🏥 Untuk keadaan darurat medis",
                        "👨‍⚕️ Ditangani tenaga medis profesional"
                    ]
                },
                {
                    "icon": "2️⃣",
                    "title": "Nomor 112",
                    "color": "BLUE",
                    "items": [
                        "📞 Layanan Panggilan Darurat Terpadu",
                        "🚒 Gabungan: Polisi, PMI, Damkar",
                        "🌐 Akses di seluruh Indonesia"
                    ]
                },
                {
                    "icon": "📋",
                    "title": "Informasi yang Harus Disampaikan",
                    "color": "RED",
                    "items": [
                        "📍 Lokasi kejadian (alamat jelas)",
                        "👥 Jumlah korban",
                        "🩺 Kondisi korban (sadarkah/tidak)",
                        "🚗 Jenis kecelakaan (laka lintas, dll)"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 6: PERIKSA KESADARAN KORBAN
    # ═══════════════════════════════════════════════════════════
    {
        "type": "two_col",
        "data": {
            "title": "👀 Periksa Kesadaran Korban (Cek Respon)",
            "subtitle": "Dekati korban dan periksa respons dengan cara yang benar",
            "left": {
                "color": "BLUE",
                "lines": [
                    "$Cara Memeriksa:",
                    "👋 Dekati korban secara perlahan",
                    "🫱 Tepuk bahu secara pelan",
                    "🗣️ Panggil: \"Pak/Bu, bisa dengar saya?\"",
                    "",
                    "$Perhatikan Reaksi:",
                    "👀 Apakah mata terbuka?",
                    "🗣️ Apakah ada respons verbal?",
                    "🖐️ Apakah ada gerakan?"
                ]
            },
            "right": {
                "color": "RED",
                "lines": [
                    "$Jika SADAR:",
                    "😌 Tenangkan korban",
                    "🚫 Larang banyak bergerak",
                    "💬 Ajak mengobrol agar panik berkurang",
                    "",
                    "$Jika TIDAK SADAR:",
                    "👀 Periksa pernapasan (gerakan dada)",
                    "🫁 Jika tidak bernapas → RJP/CPR",
                    "⚡ Kompresi dada 30x : napas buatan 2x"
                ]
            }
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 7: JANGAN PINDAHKAN KORBAN
    # ═══════════════════════════════════════════════════════════
    {
        "type": "callout",
        "data": {
            "title": "⚠️ JANGAN Pindahkan Korban Secara Sembarangan!",
            "subtitle": "Kesalahan paling fatal yang sering terjadi di masyarakat",
            "callouts": [
                {
                    "number": "❌",
                    "label": "JANGAN pindahkan korban kecuali ada ancaman nyawa langsung (kebakaran, terlindas, ledakan)",
                    "color": "RED"
                },
                {
                    "number": "⚠️",
                    "label": "Memindahkan tanpa teknik benar = risiko kelumpuhan permanen atau kematian!",
                    "color": "RED"
                }
            ],
            "note": "🦺 ATURAN EMAS:\n\n• Cedera leher/tulang belakang = JANGAN DIGERAKKAN\n• Tunggu paramedis dengan alat penyangga (neck brace, spinal board)\n• Kecuali: korban berada di posisi yang sangat berbahaya\n\n⚖️ Risiko memindahkan vs Risiko tetap di tempat — pertimbangkan dengan bijak!"
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 8: TANGANI PERDARAHAN
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "🩸 Tangani Perdarahan Hebat",
            "subtitle": "Perdarahan hebat harus segera ditangani — setiap detik berharga!",
            "cards": [
                {
                    "icon": "🔍",
                    "title": "Kenali Tanda Perdarahan Hebat",
                    "color": "RED",
                    "items": [
                        "🩸 Darah menyemprot atau mengalir deras",
                        "😰 Korban pucat, dingin, berkeringat",
                        "💓 Detak jantung cepat dan lemah",
                        "😵 Kesadaran menurun"
                    ]
                },
                {
                    "icon": "🩹",
                    "title": "Cara Menghentikan Perdarahan",
                    "color": "BLUE",
                    "items": [
                        "🧻 Gunakan kain bersih (baju, handuk, kassa)",
                        "🖐️ Tekan langsung pada titik luka",
                        "💪 Pertahankan tekanan secara konsisten",
                        "⬆️ Angkat bagian yang berdarah jika mungkin"
                    ]
                },
                {
                    "icon": "🚨",
                    "title": "Yang TIDAK Boleh Dilakukan",
                    "color": "RED",
                    "items": [
                        "❌ Jangan cabut benda asing dari luka",
                        "❌ Jangan gunakan tourniquet sembarangan",
                        "❌ Jangan buka tekanan yang sudah diberikan",
                        "❌ Jangan gunakan obat-obatan tanpa anjuran"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 9: DO'S — YANG BOLEH DILAKUKAN
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "✅ DO's — Yang BOLEH Dilakukan",
            "subtitle": "Langkah-langkah yang aman dan dianjurkan",
            "cards": [
                {
                    "icon": "🧣",
                    "title": "Jaga Suhu Tubuh Korban",
                    "color": "GREEN",
                    "items": [
                        "🧥 Tutupi dengan jaket atau selimut",
                        "🌡️ Mencegah hipotermia (suhu tubuh turun)",
                        "🛏️ Jaga korban tetap hangat"
                    ]
                },
                {
                    "icon": "⛑️",
                    "title": "Buka Kaca Helm (Jika Full-Face)",
                    "color": "GREEN",
                    "items": [
                        "🪖 Buka kaca helm pelindung saja",
                        "💨 Agar udara lancar ke saluran napas",
                        "⚠️ JANGAN lepas helm dari kepala!"
                    ]
                },
                {
                    "icon": "💬",
                    "title": "Ajak Korban Mengobrol",
                    "color": "GREEN",
                    "items": [
                        "🗣️ Jika korban sadar, ajak bicara",
                        "😌 Mencegah kepanikan",
                        "😴 Mencegah korban tertidur"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 10: DON'TS — YANG DILARANG
    # ═══════════════════════════════════════════════════════════
    {
        "type": "card_grid",
        "data": {
            "title": "❌ DON'Ts — Yang DILARANG KERAS",
            "subtitle": "Kesalahan yang dapat berakibat fatal!",
            "cards": [
                {
                    "icon": "🚫",
                    "title": "Jangan Beri Makan/Minum",
                    "color": "RED",
                    "items": [
                        "🍔 Jangan beri makan atau minum",
                        "😮 Risiko tersedak (aspirasi)",
                        "🫁 Menutup jalan napas",
                        "💀 Dapat menyebabkan kematian"
                    ]
                },
                {
                    "icon": "🪖",
                    "title": "Jangan Lepas Helm Paksa",
                    "color": "RED",
                    "items": [
                        "❌ JANGAN lepas helm secara paksa!",
                        "🦴 Risiko merusak tulang leher",
                        "🦽 Dapat menyebabkan kelumpuhan",
                        "✅ Buka kaca helm saja"
                    ]
                },
                {
                    "icon": "👥",
                    "title": "Jangan Mengerumuni Korban",
                    "color": "RED",
                    "items": [
                        "🚫 Jangan berkerumun terlalu rapat",
                        "💨 Menghalangi sirkulasi udara segar",
                        "🚑 Mengganggu akses paramedis",
                        "📸 Jangan sibuk foto/video"
                    ]
                }
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 11: TABEL DO'S AND DON'TS
    # ═══════════════════════════════════════════════════════════
    {
        "type": "table",
        "data": {
            "title": "📋 Ringkasan: Do's and Don'ts",
            "subtitle": "Perbandingan tegas antara yang boleh dan yang dilarang",
            "headers": ["✅ BOLEH (Do's)", "❌ DILARANG (Don'ts)", "⚠️ Alasan"],
            "rows": [
                ["🧣 Jaga suhu tubuh korban tetap hangat", "🍔 Beri makan atau minum", "Risiko tersedak & sumbatan jalan napas"],
                ["⛑️ Buka kaca helm (jika full-face)", "🪖 Lepas helm secara paksa", "Risiko cedera tulang leher"],
                ["💬 Ajak korban mengobrol jika sadar", "👥 Mengerumuni korban terlalu rapat", "Menghalangi udara & akses medis"],
                ["📞 Segera hubungi 119/112", "📱 Sibuk foto/video kecelakaan", "Waktu sangat kritis!"],
                ["🩹 Tekan luka dengan kain bersih", "🩸 Biarkan perdarahan terus mengalir", "Perdarahan hebat = kematian"]
            ]
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 12: THE GOLDEN HOUR
    # ═══════════════════════════════════════════════════════════
    {
        "type": "callout",
        "data": {
            "title": "⏰ The Golden Hour — Jam Emas",
            "subtitle": "60 menit pertama setelah kecelakaan = penentu keselamatan!",
            "callouts": [
                {
                    "number": "60'",
                    "label": "Menit pertama setelah trauma sangat menentukan tingkat keselamatan dan pemulihan korban",
                    "color": "RED"
                },
                {
                    "number": "🚑",
                    "label": "Kecepatan menghubungi medis profesional (119) JAUH LEBIH BERHARGA daripada menolong dengan cara salah",
                    "color": "BLUE"
                }
            ],
            "note": "💡 PENTING DIINGAT:\n\n• The Golden Hour = Jam Emas pertama setelah kecelakaan\n• Penanganan TEPAT dalam waktu ini = peluang hidup LEBIH BESAR\n• JANGAN tunda — hubungi 119/112 SEGERA!\n• Anda tidak harus jadi pahlawan — yang penting BERTINDAK CEPAT & BENAR\n\n🙏 SETIAP ORANG BISA MENYELAMATKAN NYAWA — dengan pengetahuan yang benar!"
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 13: RINGKASAN LANGKAH
    # ═══════════════════════════════════════════════════════════
    {
        "type": "flow",
        "data": {
            "title": "🔄 Ringkasan: Urutan Penanganan",
            "subtitle": "Ingat urutan ini — bisa menyelamatkan nyawa!",
            "steps": [
                {
                    "num": "1",
                    "title": "🛡️ AMAN",
                    "desc": "Amankan diri, lingkungan, dan korban dari bahaya",
                    "color": "RED"
                },
                {
                    "num": "2",
                    "title": "📞 HUBUNGI",
                    "desc": "Panggil 119/112, sampaikan lokasi & kondisi korban",
                    "color": "BLUE"
                },
                {
                    "num": "3",
                    "title": "👀 CEK",
                    "desc": "Periksa kesadaran & pernapasan korban",
                    "color": "RED"
                },
                {
                    "num": "4",
                    "title": "🩹 TANGANI",
                    "desc": "Hentikan perdarahan, jaga suhu, JANGAN pindahkan!",
                    "color": "BLUE"
                }
            ],
            "note": "⚡ INGAT: AMAN → HUBUNGI → CEK → TANGANI\nJangan terburu-buru menolong tanpa memastikan keamanan!"
        }
    },

    # ═══════════════════════════════════════════════════════════
    # SLIDE 14: PENUTUP
    # ═══════════════════════════════════════════════════════════
    {
        "type": "closing",
        "data": {
            "pre_title": "🙏 TERIMA KASIH",
            "main_title": "SELAMATKAN NYAWA\nDENGAN PENGETAHUAN",
            "subtitle": "Ingat selalu:\n• AMAN → HUBUNGI → CEK → TANGANI\n• The Golden Hour = 60 menit pertama\n• Hubungi 119 (Ambulans) atau 112 (Darurat)\n• Jangan pindahkan korban sembarangan!\n\n🏥 Pedoman: Basic Trauma Life Support (BTLS)\n🩺 Palang Merah Indonesia / PMI",
            "source": "Ada Pertanyaan? Silakan Diskusi! 📞"
        }
    }
]


if __name__ == "__main__":
    import os
    
    print("=" * 60)
    print("🚨 PPT GENERATOR — PERTOLONGAN PERTAMA")
    print("=" * 60)
    print()
    
    pres = PRESENTATION
    print(f"📋 Judul: {pres['title']}")
    print(f"📄 Slide: {len(SLIDES)} definisi")
    print(f"📁 Output: {pres['output']}")
    print()
    
    # Gunakan palette medis custom
    engine = Engine(colors=MedicalColors())
    prs = engine.build(SLIDES, source_text=pres['source'], output_path=pres['output'])
    
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
