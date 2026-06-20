#!/usr/bin/env python3
"""
content_perwal_51_2024.py — Content data untuk Perwal Bekasi No 51/2024.
==============================================================
Struktur: list of {type, data} dicts.
Tinggal ganti file ini untuk PPT dengan topik berbeda.
Engine-nya tetap di ppt_engine.py — TIDAK PERLU DIUBAH.
"""

PRESENTATION = {
    "title": "Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024",
    "subtitle": "Tentang Pengelolaan Pajak Reklame",
    "source": "Sumber: Perwal Bekasi No 51/2024",
    "output": "Perwal_Bekasi_51_2024_Pajak_Reklame.pptx",
}

SLIDES = [
    # ═══════════════════════════════════════════
    # COVER (1)
    # ═══════════════════════════════════════════
    {
        "type": "cover",
        "data": {
            "pre_title": "BERITA DAERAH",
            "city": "KOTA BEKASI",
            "main_title": "PERATURAN WALI KOTA BEKASI",
            "main_subtitle": "NOMOR 51 TAHUN 2024",
            "display_title": "TENTANG\nPENGELOLAAN PAJAK REKLAME",
            "badge_text": "Pemerintah Kota Bekasi  ·  20 Desember 2024",
            "badge_subtext": "Berlaku sejak diundangkan",
        },
    },

    # ═══════════════════════════════════════════
    # DAFTAR ISI (2)
    # ═══════════════════════════════════════════
    {
        "type": "toc",
        "data": {
            "title": "11 Bab Kunci Mengatur Seluruh Aspek Pajak Reklame",
            "subtitle": "Perwal Bekasi No 51/2024",
            "cols": 2,
            "items": [
                {"num": "1",  "label": "Ketentuan Umum & Definisi",                  "color": "#2563EB"},
                {"num": "2",  "label": "Objek, Subjek & Wajib Pajak",                "color": "#0D9488"},
                {"num": "3",  "label": "Masa Pajak & Tahun Pajak",                   "color": "#B8860B"},
                {"num": "4",  "label": "Pendaftaran & Pendataan WP",                 "color": "#1B3A6B"},
                {"num": "5",  "label": "Nilai Sewa Reklame (NSR)",                   "color": "#DC2626"},
                {"num": "6",  "label": "Perhitungan & Tarif Pajak",                  "color": "#2563EB"},
                {"num": "7",  "label": "Penetapan, Tagihan & Pembayaran",            "color": "#0D9488"},
                {"num": "8",  "label": "Pembetulan, Keberatan & Banding",            "color": "#B8860B"},
                {"num": "9",  "label": "Pemeriksaan, Penagihan & Penghapusan",       "color": "#1B3A6B"},
                {"num": "10", "label": "Keringanan, Kemudahan & Penghargaan",        "color": "#2563EB"},
                {"num": "11", "label": "Ketentuan Penutup",                          "color": "#0D9488"},
            ],
        },
    },

    # ═══════════════════════════════════════════
    # BAB I — KETENTUAN UMUM (3-4)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB I\nKETENTUAN UMUM",
            "subtitle": "Pasal 1",
            "action_text": "7 Definisi Kunci Jadi Landasan Hukum Pajak Reklame",
        },
    },
    {
        "type": "card_grid",
        "data": {
            "title": "7 Definisi Kunci Menjadi Landasan Pengelolaan Pajak Reklame",
            "subtitle": "Pasal 1",
            "cards": [
                {"icon": "🏛️", "title": "Daerah",              "color": "#2563EB", "items": ["Kota Bekasi"]},
                {"icon": "📊",  "title": "Bapenda",             "color": "#0D9488", "items": ["Badan Pendapatan Daerah Kota Bekasi"]},
                {"icon": "📢",  "title": "Reklame",             "color": "#B8860B", "items": ["Media untuk promosi & pengenalan komersial"]},
                {"icon": "💰",  "title": "Pajak Reklame",       "color": "#1B3A6B", "items": ["Pajak atas penyelenggaraan reklame"]},
                {"icon": "📐",  "title": "NSR",                 "color": "#2563EB", "items": ["Nilai Sewa Reklame — dasar pengenaan pajak"]},
                {"icon": "🆔",  "title": "NPWPD",               "color": "#0D9488", "items": ["Nomor Pokok Wajib Pajak Daerah"]},
                {"icon": "👤",  "title": "Wajib Pajak",         "color": "#B8860B", "items": ["Orang pribadi/badan dg hak & kewajiban pajak"]},
            ],
        },
    },

    # ═══════════════════════════════════════════
    # BAB II — OBJEK, SUBJEK & WAJIB PAJAK (5-7)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB II\nOBJEK, SUBJEK & WAJIB PAJAK",
            "subtitle": "Pasal 2–4",
            "action_text": "10 Jenis Reklame Wajib Pajak + 8 Pengecualian",
        },
    },
    {
        "type": "card_grid",
        "data": {
            "title": "10 Jenis Reklame Wajib Pajak + 8 Jenis Dikecualikan",
            "subtitle": "Pasal 2–4",
            "cards": [
                {"icon": "📋", "title": "10 Jenis Reklame",     "color": "#2563EB", "items": [
                    "Papan / Billboard", "Videotron / Megatron",
                    "Kain (Spanduk, Umbul, Baliho)", "Melekat / Stiker",
                    "Selebaran", "Berjalan (Kendaraan)",
                    "Udara (Balon Gas)", "Apung", "Film / Slide", "Peragaan",
                ]},
                {"icon": "🚫", "title": "Dikecualikan",         "color": "#0D9488", "items": [
                    "Internet, TV, radio, media cetak", "Label / merek produk",
                    "Nama usaha ≤ 1 m² di tempat", "Reklame Pemerintah/Pemda",
                    "Tempat ibadah & panti asuhan", "Sosial & keagamaan ≤ 30 hari",
                    "Kegiatan politik (masa kampanye)", "Olahraga KONI ≤ 30 hari",
                ]},
            ],
        },
    },
    {
        "type": "card_grid",
        "data": {
            "title": "Subjek & Wajib Pajak: Siapa yang Terkena Kewajiban Pajak",
            "subtitle": "Pasal 3–4",
            "cards": [
                {"icon": "👤", "title": "Subjek Pajak (Pasal 3)",  "color": "#2563EB", "items": [
                    "Orang pribadi atau Badan", "yang menggunakan Reklame",
                ]},
                {"icon": "✋", "title": "Wajib Pajak (Pasal 4)",   "color": "#0D9488", "items": [
                    "Orang pribadi atau Badan", "yang menyelenggarakan Reklame",
                    "Jika pihak ketiga → menjadi WP", "Mendaftarkan diri & objek pajak",
                ]},
            ],
        },
    },

    # ═══════════════════════════════════════════
    # BAB III — MASA PAJAK (8-9)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB III\nMASA PAJAK & TAHUN PAJAK",
            "subtitle": "Pasal 5",
            "action_text": "12 Bulan atau 30 Hari — Tergantung Jenis Reklame",
        },
    },
    {
        "type": "callout",
        "data": {
            "title": "Masa Pajak: 12 Bulan Permanen, 30 Hari Insidentil",
            "subtitle": "Pasal 5",
            "callouts": [
                {"number": "12", "label": "Bulan\n(Permanen)",     "color": "#2563EB"},
                {"number": "30", "label": "Hari\n(Insidentil)",    "color": "#0D9488"},
                {"number": "1",  "label": "Tahun Pajak\n(Kalender)","color": "#B8860B"},
                {"number": "1",  "label": "Bulan\n(Bagian Tahun)", "color": "#1B3A6B"},
            ],
            "note": (
                "• Masa Pajak Permanen: 12 bulan atau sesuai jangka waktu penayangan reklame\n"
                "• Masa Pajak Insidentil: dihitung per hari, maksimal 30 hari\n"
                "• Tahun Pajak: 1 tahun kalender atau sesuai tahun buku wajib pajak\n"
                "• Bagian Tahun Pajak: 1 bulan penuh (jika tidak mencakup satu tahun penuh)"
            ),
        },
    },

    # ═══════════════════════════════════════════
    # BAB IV — PENDAFTARAN & PENDATAAN (10-11)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB IV\nPENDAFTARAN & PENDATAAN WP",
            "subtitle": "Pasal 6–8",
            "action_text": "Wajib Daftar atau NPWPD Jabatan — Bapenda Berwenang Mendata",
        },
    },
    {
        "type": "two_col",
        "data": {
            "title": "Pendaftaran Wajib, Pendataan Bapenda, dan Penonaktifan WP",
            "subtitle": "Pasal 6–8",
            "left": {
                "color": "#2563EB",
                "lines": [
                    "$PENDAFTARAN (Pasal 6)", "",
                    "WP wajib mendaftarkan diri & objek pajak",
                    "Formulir: ambil/online/dikirim petugas",
                    "Lampirkan: KTP, NPWP, Akta, NIB",
                    "Bapenda terbitkan NPWPD",
                    "Jika tidak mendaftar → NPWPD jabatan",
                    "Juga: NOPD & nomor registrasi",
                ],
            },
            "right": {
                "color": "#0D9488",
                "lines": [
                    "$PENDATAAN & NONAKTIF (Pasal 7–8)", "",
                    "Bapenda mendata WP & objek pajak",
                    "Termasuk data geografis",
                    "Dapat kerjasama dengan instansi lain",
                    "Penonaktifan: WP tak penuhi syarat",
                    "Keputusan maksimal 3 bulan",
                    "Syarat: tanpa tunggakan & keberatan",
                ],
            },
        },
    },

    # ═══════════════════════════════════════════
    # BAB V — NILAI SEWA REKLAME (12-13)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB V\nNILAI SEWA REKLAME",
            "subtitle": "Pasal 9",
            "action_text": "7 Faktor Penentu NSR — Jenis, Bahan, Lokasi, Waktu, dll",
        },
    },
    {
        "type": "nsr_factors",
        "data": {
            "title": "7 Faktor Penentu Nilai Sewa Reklame (NSR)",
            "subtitle": "Pasal 9",
            "factors": [
                {"num": "1", "label": "Jenis Reklame",      "color": "#2563EB"},
                {"num": "2", "label": "Bahan",               "color": "#0D9488"},
                {"num": "3", "label": "Lokasi (Kelas Jalan)", "color": "#B8860B"},
                {"num": "4", "label": "Waktu Tayang (detik)", "color": "#1B3A6B"},
                {"num": "5", "label": "Jangka Waktu (hari)",  "color": "#2563EB"},
                {"num": "6", "label": "Jumlah Media",        "color": "#0D9488"},
                {"num": "7", "label": "Ukuran (m²)",         "color": "#B8860B"},
            ],
            "classification_note": (
                "🏛️  Kelas Jalan Khusus — Tol | Premium 1 | Premium 2\n"
                "🚗  Kelas Jalan I (Kendali Ketat) — Lebar > 3 m, pusat pelayanan\n"
                "🏡  Kelas Jalan II (Kendali Sedang) — Lebar ≤ 3 m, jalan lingkungan"
            ),
        },
    },

    # ═══════════════════════════════════════════
    # BAB VI — PERHITUNGAN & TARIF (14-20)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB VI\nPERHITUNGAN & TARIF PAJAK",
            "subtitle": "Pasal 10",
            "action_text": "Rumus: Pajak = Tarif × NSR — Dengan Berbagai Ketentuan Khusus",
        },
    },
    {
        "type": "callout",
        "data": {
            "title": "Rumus: Pajak Reklame = Tarif (50%) × NSR",
            "subtitle": "Pasal 10",
            "callouts": [
                {"number": "×",    "label": "Pajak =\nTarif × NSR",        "color": "#2563EB"},
                {"number": "50%",  "label": "Indoor =\n50% NSR",            "color": "#0D9488"},
                {"number": "+20%", "label": "Tinggi > 15m\ntambahan 20%",   "color": "#B8860B"},
                {"number": "+50%", "label": "Tembakau &\nMiras +50%",       "color": "#DC2626"},
            ],
            "note": (
                "Rumus: Pajak Reklame = Tarif Pajak × NSR\n"
                "NSR = Nilai Kelas Jalan × Ukuran (m²) × Jumlah × Jangka Waktu\n\n"
                "Ketentuan Khusus:\n"
                "• Reklame indoor: NSR 50% dari NSR normal\n"
                "• Ketinggian > 15 meter: tambahan 20%\n"
                "• Produk tembakau & minuman beralkohol: tambahan 50%\n"
                "• Perubahan naskah/revisi isi reklame: dikecualikan"
            ),
        },
    },
    # Tabel Papan/Billboard
    {
        "type": "table",
        "data": {
            "title": "Tabel NSR — Papan / Billboard (Rp/m²/hari)",
            "subtitle": "Pasal 10",
            "headers": ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"],
            "rows": [
                ["Kelas Jalan Khusus", "Jalan Tol",     "23.575"],
                ["Kelas Jalan Khusus", "Premium 1",     "16.100"],
                ["Kelas Jalan Khusus", "Premium 2",     "14.950"],
                ["Kelas Jalan I",      "Kendali Ketat",  "13.225"],
                ["Kelas Jalan II",     "Kendali Sedang", "11.500"],
            ],
        },
    },
    # Tabel Megatron/Videotron
    {
        "type": "table",
        "data": {
            "title": "Tabel NSR — Megatron / Videotron",
            "subtitle": "Pasal 10",
            "headers": ["Kelas Jalan", "Zona", "NSR (/30 dtk)", "NSR (/m²/thn)"],
            "rows": [
                ["Kelas Jalan Khusus", "Jalan Tol",     "Rp 17,25",  "13.599.900"],
                ["Kelas Jalan Khusus", "Premium 1",     "Rp 13,80",  "10.879.920"],
                ["Kelas Jalan Khusus", "Premium 2",     "Rp 9,20",   "7.253.280"],
                ["Kelas Jalan I",      "Kendali Ketat",  "Rp 8,05",   "6.346.620"],
                ["Kelas Jalan II",     "Kendali Sedang", "Rp 5,75",   "4.533.300"],
            ],
        },
    },
    # Tabel Kain
    {
        "type": "table",
        "data": {
            "title": "Tabel NSR — Kain (Spanduk/Umbul/Baliho) (Rp/m²/hari)",
            "subtitle": "Pasal 10",
            "headers": ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"],
            "rows": [
                ["Kelas Jalan Khusus", "Jalan Tol",     "30.000"],
                ["Kelas Jalan Khusus", "Premium 1",     "30.000"],
                ["Kelas Jalan Khusus", "Premium 2",     "25.000"],
                ["Kelas Jalan I",      "Kendali Ketat",  "20.000"],
                ["Kelas Jalan II",     "Kendali Sedang", "19.000"],
            ],
        },
    },
    # NSR lainnya bagian 1
    {
        "type": "card_grid",
        "data": {
            "title": "NSR untuk 8 Jenis Reklame Lainnya — Bagian 1",
            "subtitle": "Pasal 10",
            "cards": [
                {"icon": "🏷️", "title": "Stiker",    "color": "#2563EB", "items": ["Rp 7,5/cm²", "Min. Rp 750.000/kali"]},
                {"icon": "🧱",  "title": "Melekat",   "color": "#0D9488", "items": ["Rp 750.000/m²/tahun"]},
                {"icon": "📄",  "title": "Selebaran", "color": "#B8860B", "items": ["Rp 600/lembar", "Min. Rp 6.000.000/kali"]},
                {"icon": "🚌",  "title": "Berjalan",  "color": "#DC2626", "items": ["Rp 6.000/m²/hari", "Termasuk kendaraan"]},
            ],
        },
    },
    # NSR lainnya bagian 2
    {
        "type": "card_grid",
        "data": {
            "title": "NSR untuk 8 Jenis Reklame Lainnya — Bagian 2",
            "subtitle": "Pasal 10",
            "cards": [
                {"icon": "🎈", "title": "Udara",        "color": "#2563EB", "items": ["Rp 2.400.000/sekali", "Maks. 1 bulan"]},
                {"icon": "🌊", "title": "Apung",        "color": "#0D9488", "items": ["Rp 600.000/sekali", "Maks. 1 bulan"]},
                {"icon": "🎬", "title": "Film / Slide", "color": "#B8860B", "items": ["Rp 12.000/15 detik"]},
                {"icon": "🎭", "title": "Peragaan",     "color": "#DC2626", "items": ["Rp 480.000/penyelenggaraan"]},
            ],
        },
    },

    # ═══════════════════════════════════════════
    # BAB VII — PENETAPAN, TAGIHAN & PEMBAYARAN (21-22)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB VII\nPENETAPAN, TAGIHAN & PEMBAYARAN",
            "subtitle": "Pasal 11–14",
            "action_text": "4 Langkah: SKPD → Bayar → Telat → STPD",
        },
    },
    {
        "type": "flow",
        "data": {
            "title": "Alur Penetapan hingga Pembayaran: 4 Langkah Wajib Dipahami WP",
            "subtitle": "Pasal 11–14",
            "steps": [
                {"num": "1", "title": "SKPD",   "desc": "Diterbitkan Bapenda\nMasa berlaku 5 tahun",   "color": "#2563EB"},
                {"num": "2", "title": "Pembayaran", "desc": "Lunas 1 bulan\nsejak SKPD diterima",        "color": "#0D9488"},
                {"num": "3", "title": "Keterlambatan", "desc": "Bunga 1%/bln\nDiterbitkan STPD",          "color": "#B8860B"},
                {"num": "4", "title": "STPD",  "desc": "Harus lunas\n≤ 30 hari",                       "color": "#DC2626"},
            ],
            "note": (
                "• Jatuh tempo: 1 bulan sejak tanggal pengiriman SKPD\n"
                "• Pembayaran: Kas Daerah / Bank Persepsi / tempat lain yang ditunjuk\n"
                "• Stiker sebagai tanda bukti pembayaran reklame\n"
                "• STPD dikenakan bunga 1%/bulan (maks. 24 bulan)"
            ),
        },
    },

    # ═══════════════════════════════════════════
    # BAB VIII — PEMBETULAN, KEBERATAN & BANDING (23-25)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB VIII\nPEMBETULAN, KEBERATAN & BANDING",
            "subtitle": "Pasal 15–20, 29–33",
            "action_text": "Hak WP: Koreksi, Keberatan, Banding — Dengan Batas Waktu Jelas",
        },
    },
    {
        "type": "two_col",
        "data": {
            "title": "Pembetulan: Koreksi Kesalahan Tulis, Hitung, dan Penerapan Aturan",
            "subtitle": "Pasal 15–20",
            "left": {
                "color": "#2563EB",
                "lines": [
                    "$PEMBETULAN (Pasal 15–20)", "",
                    "Kesalahan tulis: nama, alamat, NPWPD",
                    "Kesalahan hitung: jumlah, tarif",
                    "Kekeliruan penerapan aturan",
                    "1 permohonan = 1 ketetapan",
                    "Keputusan maksimal 6 bulan",
                    "> 6 bulan tanpa putusan → dikabulkan",
                    "Dapat dilakukan berulang (Ps 20)",
                    "Jenis keputusan: kabul / batal / tolak",
                ],
            },
            "right": {
                "color": "#B8860B",
                "lines": [
                    "$JANGKA WAKTU & PROSEDUR", "",
                    "Permohonan diajukan ke Bapenda",
                    "Keputusan: kabul (tambah/kurang/hapus)",
                    "Keputusan: batal | tolak",
                    "Pasal 19: pembetulan jabatan",
                    "Pasal 20: berulang jika masih salah",
                ],
            },
        },
    },
    {
        "type": "two_col",
        "data": {
            "title": "Keberatan & Banding: Upaya Hukum WP dalam Sengketa Pajak",
            "subtitle": "Pasal 29–33",
            "left": {
                "color": "#2563EB",
                "lines": [
                    "$KEBERATAN (Pasal 29–31)", "",
                    "Objek: SKPD, SKPDKB, SKPDKBT, dll",
                    "Diajukan maks. 3 bulan sejak SKPD",
                    "Sudah bayar min. yang disetujui",
                    "Keputusan maks. 12 bulan",
                    "Jika ditolak: denda 30%",
                    "Jika dikabulkan: + bunga 0,6%/bulan",
                ],
            },
            "right": {
                "color": "#0D9488",
                "lines": [
                    "$BANDING (Pasal 32–33)", "",
                    "Objek: Surat Keputusan Keberatan",
                    "Ke badan peradilan pajak",
                    "Maks. 3 bulan sejak keputusan",
                    "Menangguhkan kewajiban bayar",
                    "Jika ditolak: denda 60%",
                    "Jika dikabulkan: + bunga 0,6%/bulan",
                ],
            },
        },
    },

    # ═══════════════════════════════════════════
    # BAB IX — PEMERIKSAAN, PENAGIHAN & PENGHAPUSAN (26-28)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB IX\nPEMERIKSAAN, PENAGIHAN & PENGHAPUSAN",
            "subtitle": "Pasal 22–26",
            "action_text": "Bapenda Berwenang Periksa — Piutang Dihapus Lewat 4 Tahapan",
        },
    },
    {
        "type": "card_grid",
        "data": {
            "title": "Pemeriksaan & Penagihan: 3 Pilar Penegakan Kepatuhan WP",
            "subtitle": "Pasal 22–25",
            "cards": [
                {"icon": "🔍", "title": "Pemeriksaan (Ps 22–23)",   "color": "#2563EB", "items": [
                    "Kepala Bapenda berwenang periksa",
                    "Menguji kepatuhan WP",
                    "WP wajib: buka buku/dokumen",
                    "Beri akses tempat & keterangan",
                    "Jika tidak → pajak ditetapkan jabatan",
                ]},
                {"icon": "📬", "title": "Penagihan (Ps 24)",         "color": "#0D9488", "items": [
                    "Dasar: SKPD, SKPDKB, SKPDKBT",
                    "STPD, SK Pembetulan/Keberatan",
                    "Putusan Banding",
                ]},
                {"icon": "⏳", "title": "Kedaluwarsa (Ps 25)",       "color": "#B8860B", "items": [
                    "5 tahun sejak pajak terutang",
                    "Tertangguh jika ada: Surat Teguran / Paksa",
                    "Pengakuan utang dari WP",
                ]},
            ],
        },
    },
    {
        "type": "flow",
        "data": {
            "title": "Penghapusan Piutang Pajak: 4 Langkah dari Penelitian hingga SK",
            "subtitle": "Pasal 26",
            "steps": [
                {"num": "1", "title": "Penelitian",  "desc": "Dilakukan Bapenda",        "color": "#2563EB"},
                {"num": "2", "title": "Penetapan",   "desc": "Keputusan Wali Kota",      "color": "#0D9488"},
                {"num": "3", "title": "Koordinasi",  "desc": "Dengan Inspektorat",       "color": "#B8860B"},
                {"num": "4", "title": "SK Penghapusan", "desc": "Diterbitkan",              "color": "#1B3A6B"},
            ],
            "note": (
                "Syarat penghapusan piutang pajak:\n"
                "• Piutang tidak mungkin ditagih lagi karena kedaluwarsa\n"
                "• Ada koordinasi dengan Inspektorat Daerah\n"
                "• Dibuktikan dengan dokumen pelaksanaan penagihan"
            ),
        },
    },

    # ═══════════════════════════════════════════
    # BAB X — KERINGANAN, KEMUDAHAN & PENGHARGAAN (29-30)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB X\nKERINGANAN, KEMUDAHAN & PENGHARGAAN",
            "subtitle": "Pasal 27–28, 34–35",
            "action_text": "3 Fasilitas WP: Keringanan, Kemudahan Angsuran, dan Penghargaan",
        },
    },
    {
        "type": "card_grid",
        "data": {
            "title": "3 Fasilitas untuk WP: Keringanan, Kemudahan Angsuran, Penghargaan",
            "subtitle": "Pasal 27–28, 34–35",
            "cards": [
                {"icon": "🎯", "title": "Keringanan (Ps 27)",       "color": "#2563EB", "items": [
                    "Keringanan / Pengurangan", "Pembebasan / Penundaan",
                    "Atas pokok & sanksi pajak", "WP dengan likuiditas rendah",
                    "Objek terdampak bencana/kebakaran",
                ]},
                {"icon": "🤝", "title": "Kemudahan (Ps 28)",        "color": "#0D9488", "items": [
                    "Perpanjangan waktu bayar", "Angsuran maks. 24 bulan",
                    "Bunga 0,6%/bulan", "Keadaan kahar: bencana, wabah, kerusuhan",
                ]},
                {"icon": "🏆", "title": "Penghargaan (Ps 34–35)",   "color": "#B8860B", "items": [
                    "WP Taat Pajak", "Bayar tepat waktu ≥ 1 tahun",
                    "Tanpa tunggakan 3 tahun", "Kontribusi signifikan",
                    "Piagam / Hadiah (APBD)",
                ]},
            ],
        },
    },

    # ═══════════════════════════════════════════
    # BAB XI — KETENTUAN PENUTUP (31-32)
    # ═══════════════════════════════════════════
    {
        "type": "section",
        "data": {
            "title": "BAB XI\nKETENTUAN PENUTUP",
            "subtitle": "Pasal 36–37",
            "action_text": "Perwal No 48/2012 Dicabut — Berlaku Sejak 20 Desember 2024",
        },
    },
    {
        "type": "two_col",
        "data": {
            "title": "Perwal Lama Dicabut, Peraturan Baru Berlaku Mulai Diundangkan",
            "subtitle": "Pasal 36–37",
            "left": {
                "color": "#2563EB",
                "lines": [
                    "$PERATURAN YANG DICABUT (Pasal 36)", "",
                    "Perwal No. 48 Tahun 2012",
                    "Petunjuk Pelaksanaan Perda 14/2012",
                    "Perwal No. 52 Tahun 2013 (Perubahan)",
                ],
            },
            "right": {
                "color": "#1B3A6B",
                "lines": [
                    "$MULAI BERLAKU (Pasal 37)", "",
                    "Sejak diundangkan",
                    "20 Desember 2024", "",
                    "Pj. WALI KOTA BEKASI,",
                    "ttd.",
                    "R. GANI MUHAMAD",
                ],
            },
        },
    },

    # ═══════════════════════════════════════════
    # CLOSING (33)
    # ═══════════════════════════════════════════
    {
        "type": "closing",
        "data": {
            "pre_title": "BERITA DAERAH KOTA BEKASI",
            "main_title": "TERIMA KASIH",
            "subtitle": (
                "Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024\n"
                "Tentang Pengelolaan Pajak Reklame"
            ),
            "source": "Sumber: https://jdih.bekasikota.go.id",
        },
    },
]
