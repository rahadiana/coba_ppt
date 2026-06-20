#!/usr/bin/env node
/**
 * PPT RINGKAS — Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024
 * Tentang Pengelolaan Pajak Reklame
 * 
 * Versi ringkas ~15 slide, fokus pada poin-poin utama
 * Menggunakan PptxGenJS + best practice Anthropic
 */

const pptxgen = require("pptxgenjs");

const C = {
  NAVY:"0A1628", DARK_BLUE:"12294A", MID_BLUE:"1B3A6B",
  ACCENT_BLUE:"2D7DD2", GOLD:"D4A017", WARM_GOLD:"B8860B",
  WHITE:"FFFFFF", OFF_WHITE:"F0F2F5", ICE_BLUE:"E8EDF5",
  DARK_TEXT:"1A1A2E", GRAY_TEXT:"6B7288", MUTED:"9CA3AF",
  TEAL:"0D9488", RED_ACCENT:"DC2626",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Pemerintah Kota Bekasi";
pres.title = "Perwal Bekasi No 51/2024 - Pengelolaan Pajak Reklame";

let slideNum = 0;
const page = (s) => { slideNum++; s.addText(String(slideNum), { x:12.2, y:7.0, w:0.8, h:0.35, fontSize:9, color:C.GRAY_TEXT, align:"right", fontFace:"Calibri" }); };

// ─── COVER ───
(() => {
  const s = pres.addSlide(); s.background = { color: C.NAVY };
  s.addShape(pres.shapes.OVAL, { x:8.5, y:-2, w:7, h:7, fill:{ color:C.DARK_BLUE, transparency:60 } });
  s.addShape(pres.shapes.OVAL, { x:10, y:4, w:5, h:5, fill:{ color:C.DARK_BLUE, transparency:60 } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.8, y:2.4, w:3.5, h:0.06, fill:{ color:C.GOLD } });
  s.addText("BERITA DAERAH KOTA BEKASI", { x:0.8, y:0.4, w:8, h:0.4, fontSize:13, color:C.GOLD, bold:true, fontFace:"Calibri" });
  s.addText("NOMOR 51  ·  TAHUN 2024", { x:0.8, y:0.8, w:5, h:0.3, fontSize:11, color:C.MUTED, fontFace:"Calibri" });
  s.addText("PERATURAN WALI KOTA BEKASI\nNOMOR 51 TAHUN 2024", { x:0.8, y:1.5, w:11, h:1.0, fontSize:26, bold:true, color:C.WHITE, fontFace:"Calibri" });
  s.addText("TENTANG\nPENGELOLAAN PAJAK REKLAME", { x:0.8, y:3.0, w:11, h:1.6, fontSize:40, bold:true, color:C.GOLD, fontFace:"Calibri" });
  s.addText("Pemerintah Kota Bekasi  |  20 Desember 2024", { x:0.8, y:5.0, w:8, h:0.35, fontSize:13, color:C.ICE_BLUE, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.GOLD } });
})();

// ─── HELPER: Section ───
const section = (title, sub="") => {
  const s = pres.addSlide(); s.background = { color:C.NAVY };
  s.addShape(pres.shapes.OVAL, { x:9.5, y:-1.5, w:5, h:5, fill:{ color:C.DARK_BLUE, transparency:60 } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.8, y:2.0, w:2.0, h:0.06, fill:{ color:C.GOLD } });
  s.addText(title, { x:0.8, y:2.3, w:11, h:1.5, fontSize:36, bold:true, color:C.WHITE, fontFace:"Calibri" });
  if (sub) s.addText(sub, { x:0.8, y:4.0, w:10, h:0.5, fontSize:14, color:C.MUTED, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.GOLD } });
  page(s);
};

// ─── HELPER: Content ───
const slide = (title, items, opts={}) => {
  const s = pres.addSlide(); s.background = { color:C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:13.333, h:0.04, fill:{ color:C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.04, w:13.333, h:0.75, fill:{ color:C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.12, w:0.08, h:0.5, fill:{ color:C.ACCENT_BLUE } });
  s.addText(title, { x:0.75, y:0.12, w:11, h:0.5, fontSize:22, bold:true, color:C.DARK_TEXT, fontFace:"Calibri" });
  if (opts.notes) s.addText(opts.notes, { x:0.75, y:0.5, w:10, h:0.3, fontSize:11, color:C.GRAY_TEXT, fontFace:"Calibri" });

  const arr = items.map((t,i) => {
    if (t==="") return { text:"", options:{ breakLine:true, fontSize:6 } };
    const isH = t.startsWith("**"); const isS = t.startsWith("   "); const cl = t.replace(/\*\*/g,"").trim();
    return { text:cl, options:{ bullet:!isS && cl.length>0, indentLevel:isS?1:0, bold:isH,
      breakLine:i<items.length-1, fontSize:isH?15:14, color:C.DARK_TEXT, fontFace:"Calibri", paraSpaceAfter:2 } };
  });
  s.addText(arr, { x:0.6, y:1.0, w:11.5, h:5.5, valign:"top", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.3, w:13.333, h:0.2, fill:{ color:C.NAVY } });
  page(s);
};

// ─── HELPER: Two-col ───
const twoCol = (title, L, R, opts={}) => {
  const s = pres.addSlide(); s.background = { color:C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:13.333, h:0.04, fill:{ color:C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.04, w:13.333, h:0.75, fill:{ color:C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.12, w:0.08, h:0.5, fill:{ color:C.ACCENT_BLUE } });
  s.addText(title, { x:0.75, y:0.12, w:11, h:0.5, fontSize:22, bold:true, color:C.DARK_TEXT, fontFace:"Calibri" });
  if (opts.notes) s.addText(opts.notes, { x:0.75, y:0.5, w:10, h:0.3, fontSize:11, color:C.GRAY_TEXT, fontFace:"Calibri" });

  const mk = (items) => items.map((t,i) => {
    if(t==="") return { text:"", options:{ breakLine:true, fontSize:6 } };
    const isH=t.startsWith("**"); const isS=t.startsWith("   "); const cl=t.replace(/\*\*/g,"").trim();
    return { text:cl, options:{ bullet:!isS&&cl.length>0, indentLevel:isS?1:0, bold:isH,
      breakLine:i<items.length-1, fontSize:13, color:C.DARK_TEXT, fontFace:"Calibri", paraSpaceAfter:2 } };
  });
  
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.5, y:1.0, w:5.8, h:5.5, fill:{ color:C.ICE_BLUE, transparency:50 } });
  s.addText(mk(L), { x:0.65, y:1.1, w:5.5, h:5.3, valign:"top", margin:0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:6.8, y:1.0, w:5.8, h:5.5, fill:{ color:C.OFF_WHITE } });
  s.addText(mk(R), { x:6.95, y:1.1, w:5.5, h:5.3, valign:"top", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.3, w:13.333, h:0.2, fill:{ color:C.NAVY } });
  page(s);
};

// ─── HELPER: Simple Table ───
const tbl = (title, headers, rows, opts={}) => {
  const s = pres.addSlide(); s.background = { color:C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:13.333, h:0.04, fill:{ color:C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.04, w:13.333, h:0.75, fill:{ color:C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.12, w:0.08, h:0.5, fill:{ color:C.ACCENT_BLUE } });
  s.addText(title, { x:0.75, y:0.12, w:11, h:0.5, fontSize:22, bold:true, color:C.DARK_TEXT, fontFace:"Calibri" });
  if (opts.notes) s.addText(opts.notes, { x:0.75, y:0.5, w:10, h:0.3, fontSize:11, color:C.GRAY_TEXT, fontFace:"Calibri" });

  const hRow = headers.map(h => ({ text:h, options:{ bold:true, color:C.WHITE, fill:{ color:C.NAVY }, fontSize:11, fontFace:"Calibri", align:"center", valign:"middle" } }));
  const dRows = rows.map((r,i) => r.map((c,j) => ({ text:String(c), options:{ fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri", fill:{ color:i%2===0?C.ICE_BLUE:C.WHITE }, align:j===0?"left":"center", valign:"middle" } })));
  
  s.addTable([hRow, ...dRows], { x:0.8, y:1.1, w:11.7, colW:opts.colW||undefined, border:{ pt:0.5, color:C.MUTED }, rowH:0.45, margin:[2,6,2,6] });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.3, w:13.333, h:0.2, fill:{ color:C.NAVY } });
  page(s);
};

// ═══════════════════════════════════════════════
// RINGKASAN ISI
// ═══════════════════════════════════════════════

// ─── DAFTAR ISI ───
(() => {
  const s = pres.addSlide(); s.background = { color:C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:13.333, h:0.04, fill:{ color:C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.04, w:13.333, h:0.75, fill:{ color:C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.12, w:0.08, h:0.5, fill:{ color:C.ACCENT_BLUE } });
  s.addText("DAFTAR ISI RINGKAS", { x:0.75, y:0.12, w:11, h:0.5, fontSize:24, bold:true, color:C.DARK_TEXT, fontFace:"Calibri" });
  s.addText("Perwal Bekasi No 51/2024 — Pengelolaan Pajak Reklame", { x:0.75, y:0.5, w:10, h:0.3, fontSize:11, color:C.GRAY_TEXT, fontFace:"Calibri" });

  const toc = [
    ["I", "Ketentuan Umum & Definisi"],
    ["II", "Objek, Subjek & Wajib Pajak"],
    ["III", "Masa Pajak & Tahun Pajak"],
    ["IV", "Pendaftaran & Pendataan WP"],
    ["V", "Nilai Sewa Reklame (NSR)"],
    ["VI", "Perhitungan & Tarif Pajak"],
    ["VII", "Penetapan & Pembayaran"],
    ["VIII", "Pembetulan, Keberatan & Banding"],
    ["IX", "Pemeriksaan, Penagihan & Penghapusan"],
    ["X", "Keringanan, Kemudahan & Penghargaan"],
    ["XI", "Ketentuan Penutup"],
  ];
  toc.forEach(([n, ttl], i) => {
    const y = 1.1 + i * 0.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.6, y, w:0.6, h:0.35, fill:{ color:C.ACCENT_BLUE } });
    s.addText(n, { x:0.6, y:y+0.01, w:0.6, h:0.33, fontSize:10, bold:true, color:C.WHITE, align:"center", fontFace:"Calibri" });
    s.addText(ttl, { x:1.4, y:y+0.01, w:5, h:0.33, fontSize:13, color:C.DARK_TEXT, fontFace:"Calibri" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.3, w:13.333, h:0.2, fill:{ color:C.NAVY } });
  page(s);
})();

// ─── BAB I: KETENTUAN UMUM ───
section("BAB I: KETENTUAN UMUM", "Pasal 1 — Definisi Kunci");
slide("Definisi Penting", [
  "**Daerah** = Kota Bekasi",
  "**Bapenda** = Badan Pendapatan Daerah Kota Bekasi",
  "**Reklame** = Benda/alat/media untuk tujuan komersial (promosi, pengenalan)",
  "**Pajak Reklame** = Pajak atas penyelenggaraan reklame",
  "**NSR (Nilai Sewa Reklame)** = Dasar pengenaan Pajak Reklame",
  "**NPWPD** = Nomor Pokok Wajib Pajak Daerah",
  "**Wajib Pajak** = Orang pribadi/badan dengan hak & kewajiban perpajakan",
]);

// ─── BAB II: OBJEK, SUBJEK, WAJIB PAJAK ───
section("BAB II: OBJEK, SUBJEK & WAJIB PAJAK", "Pasal 2–4");
twoCol("Objek & Pengecualian Pajak Reklame",
  ["**Objek Pajak:**", "", "10 jenis reklame:", "Papan/Billboard", "Videotron/Megatron", "Kain (spanduk, umbul, baliho)", "Melekat/Stiker", "Selebaran", "Berjalan (kendaraan)", "Udara (balon gas)", "Apung", "Film/Slide", "Peragaan"],
  ["**Dikecualikan:**", "", "Internet, TV, radio, media cetak", "Label/merek produk", "Nama usaha ≤ 1 m² di tempat", "Reklame Pemerintah/Pemda", "Tempat ibadah & panti asuhan", "Kegiatan politik (masa kampanye)", "Sosial/keagamaan (≤ 30 hari)", "Olahraga KONI (≤ 30 hari)", "", "**Subjek:** Pengguna reklame", "**Wajib Pajak:** Penyelenggara reklame"]
);

// ─── BAB III: MASA PAJAK ───
section("BAB III: MASA & TAHUN PAJAK", "Pasal 5");
slide("Masa Pajak", [
  "**Masa Pajak Permanen:** 12 bulan / sesuai masa penayangan",
  "**Masa Pajak Insidentil:** per hari, maks. 30 hari",
  "**Tahun Pajak:** 1 tahun kalender (atau sesuai tahun buku WP)",
]);

// ─── BAB IV: PENDAFTARAN & PENDATAAN ───
section("BAB IV: PENDAFTARAN & PENDATAAN WP", "Pasal 6–8");
twoCol("Pendaftaran, Pendataan & Penonaktifan",
  ["**Pendaftaran (Pasal 6):**", "", "WP daftarkan diri + objek pajak", "Form: ambil/online/dikirim petugas", "Lampirkan: KTP, NPWP, Akta, NIB", "Bapenda terbitkan NPWPD", "Jika tak daftar → NPWPD jabatan"],
  ["**Pendataan (Pasal 7):**", "Bapenda data WP & objek pajak", "Termasuk data geografis", "", "**Penonaktifan (Pasal 8):**", "WP tak penuhi syarat → nonaktif", "Keputusan maks. 3 bulan", "Syarat: tanpa tunggakan & keberatan"]
);

// ─── BAB V: NILAI SEWA REKLAME ───
section("BAB V: NILAI SEWA REKLAME (NSR)", "Pasal 9 — Dasar Pengenaan Pajak");
slide("Faktor Penentu NSR", [
  "**NSR** = Nilai Jual Objek Pajak + Nilai Strategis Pemasangan",
  "",
  "**7 Faktor:** Jenis reklame | Bahan | Lokasi (Kelas Jalan) | Waktu (detik) | Jangka waktu (hari) | Jumlah media | Ukuran (m²)",
  "",
  "**Kelas Jalan:** Khusus (Tol, Premium I, Premium II) | Kelas I (Kendali Ketat) | Kelas II (Kendali Sedang)",
]);

// ─── BAB VI: PERHITUNGAN ───
section("BAB VI: PERHITUNGAN & TARIF PAJAK", "Pasal 10");
slide("Rumus Perhitungan", [
  "Pajak Reklame = Tarif × NSR",
  "NSR Papan/Billboard = Nilai Kelas Jalan × Ukuran × Jumlah × Jangka Waktu",
  "",
  "**Ketentuan Khusus:**",
  "Reklame indoor: NSR = 50%",
  "Ketinggian > 15 m: tambahan 20%",
  "Produk tembakau & miras: tambahan 50%",
]);

tbl("NSR — Papan/Billboard", ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"], [
  ["Kelas Jalan Khusus", "Jalan Tol", "Rp 23.575"],
  ["Kelas Jalan Khusus", "Premium 1", "Rp 16.100"],
  ["Kelas Jalan Khusus", "Premium 2", "Rp 14.950"],
  ["Kelas Jalan I", "Kendali Ketat", "Rp 13.225"],
  ["Kelas Jalan II", "Kendali Sedang", "Rp 11.500"],
], { notes:"Satuan: 1 m², 1 buah, 1 hari", colW:[3,6,2.7] });

tbl("NSR — Megatron/Videotron", ["Kelas Jalan", "Zona", "NSR (/30 detik)", "NSR (/m²/tahun)"], [
  ["Kelas Jalan Khusus", "Jalan Tol", "Rp 17,25", "Rp 13.599.900"],
  ["Kelas Jalan Khusus", "Premium 1", "Rp 13,80", "Rp 10.879.920"],
  ["Kelas Jalan Khusus", "Premium 2", "Rp 9,20", "Rp 7.253.280"],
  ["Kelas Jalan I", "Kendali Ketat", "Rp 8,05", "Rp 6.346.620"],
  ["Kelas Jalan II", "Kendali Sedang", "Rp 5,75", "Rp 4.533.300"],
], { colW:[3,4,2.7,2.7] });

tbl("NSR — Kain (Spanduk/Umbul/Baliho)", ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"], [
  ["Kelas Jalan Khusus", "Jalan Tol", "Rp 30.000"],
  ["Kelas Jalan Khusus", "Premium 1", "Rp 30.000"],
  ["Kelas Jalan Khusus", "Premium 2", "Rp 25.000"],
  ["Kelas Jalan I", "Kendali Ketat", "Rp 20.000"],
  ["Kelas Jalan II", "Kendali Sedang", "Rp 19.000"],
], { notes:"Satuan: 1 m², 1 buah, 1 hari", colW:[3,6,2.7] });

slide("NSR — Jenis Lainnya", [
  "**Stiker:** Rp 7,5/cm² (min. Rp 750.000/kali)",
  "**Melekat:** Rp 750.000/m²/tahun",
  "**Selebaran:** Rp 600/lembar (min. Rp 6.000.000/kali)",
  "**Berjalan:** Rp 6.000/m²/hari",
  "**Udara:** Rp 2.400.000/sekali (maks. 1 bln)",
  "**Apung:** Rp 600.000/sekali (maks. 1 bln)",
  "**Film/Slide:** Rp 12.000/15 detik",
  "**Peragaan:** Rp 480.000/penyelenggaraan",
]);

// ─── BAB VII: PENETAPAN & PEMBAYARAN ───
section("BAB VII: PENETAPAN, TAGIHAN & PEMBAYARAN", "Pasal 11–14");
slide("Ringkasan Penetapan hingga Pembayaran", [
  "**Penetapan (Pasal 11):** SKPD berdasarkan pendaftaran / jabatan / hasil periksa — maks. 5 tahun",
  "",
  "**STPD (Pasal 12):**",
  "Jika pajak tak/kurang dibayar → STPD + bunga 1%/bln (maks. 24 bln)",
  "Jika SK Keberatan/Banding tak dibayar → bunga 0,6%/bln",
  "",
  "**Pembayaran (Pasal 13–14):**",
  "Lunas via Kas Daerah/bank persepsi — maks. 1 bulan sejak SKPD",
  "Keterlambatan: bunga 1%/bln + STPD (lunas ≤ 30 hari)",
]);

// ─── BAB VIII: PEMBETULAN, KEBERATAN, BANDING ───
section("BAB VIII: PEMBETULAN, KEBERATAN & BANDING", "Pasal 15–20 (Pembetulan), 29–33 (Keberatan & Banding)");
twoCol("Pembetulan, Keberatan & Banding",
  ["**Pembetulan (Pasal 15–20):**", "", "Kesalahan tulis/hitung/penerapan", "Permohonan → keputusan 6 bulan", "Jika > 6 bln tanpa keputusan → dikabulkan", "Bisa jabatan (Pasal 19)", "Bisa berulang (Pasal 20)"],
  ["**Keberatan (Pasal 29–31):**", "Maks. 3 bulan sejak SKPD", "Keputusan maks. 12 bulan", "Ditolak: denda 30%", "", "**Banding (Pasal 32–33):**", "Maks. 3 bulan sejak SK Keberatan", "Ditolak: denda 60%", "Dikabulkan: + bunga 0,6%/bln"]
);

// ─── BAB IX: PEMERIKSAAN, PENAGIHAN, PENGHAPUSAN ───
section("BAB IX: PEMERIKSAAN, PENAGIHAN & PENGHAPUSAN", "Pasal 22–26");
slide("Pemeriksaan, Penagihan & Penghapusan", [
  "**Pemeriksaan (Pasal 22–23):** Kepala Bapenda berwenang periksa kepatuhan WP",
  "WP wajib: tunjukkan dokumen, beri akses tempat, beri keterangan",
  "Jika tidak → pajak ditetapkan jabatan",
  "",
  "**Penagihan (Pasal 24):** Dasar: SKPD, SKPDKB, SKPDKBT, STPD, dll.",
  "",
  "**Kedaluwarsa (Pasal 25):** 5 tahun sejak pajak terutang",
  "Tertangguh jika: ada Surat Teguran/Paksa atau ada pengakuan utang",
  "",
  "**Penghapusan Piutang (Pasal 26):** Untuk piutang tak tertagih → Keputusan Wali Kota",
]);

// ─── BAB X: KERINGANAN, KEMUDAHAN, PENGHARGAAN ───
section("BAB X: KERINGANAN, KEMUDAHAN & PENGHARGAAN", "Pasal 27–28, 34–35");
slide("Fasilitas & Penghargaan", [
  "**Keringanan/Pengurangan/Pembebasan (Pasal 27):**",
  "Untuk WP dengan likuiditas rendah atau objek terdampak bencana",
  "",
  "**Kemudahan Perpajakan (Pasal 28):**",
  "Perpanjangan waktu bayar & angsuran (maks. 24 bln, bunga 0,6%/bln)",
  "Untuk WP keadaan kahar: bencana, kebakaran, kerusuhan, wabah",
  "",
  "**Penghargaan (Pasal 34–35):**",
  "Untuk WP Taat Pajak: bayar tepat waktu ≥ 1 tahun, tanpa tunggakan 3 tahun",
]);

// ─── BAB XI: PENUTUP ───
section("BAB XI: KETENTUAN PENUTUP", "Pasal 36–37");
twoCol("Pencabutan & Mulai Berlaku",
  ["**Pasal 36 — Dicabut:**", "", "❌ Perwal No. 48/2012", "❌ Perwal No. 52/2013"],
  ["**Pasal 37 — Berlaku:**", "", "✅ Sejak diundangkan", "📅 20 Desember 2024", "", "**Pj. WALI KOTA BEKASI:**", "R. GANI MUHAMAD"]
);

// ─── CLOSING ───
(() => {
  const s = pres.addSlide(); s.background = { color:C.NAVY };
  s.addShape(pres.shapes.OVAL, { x:9, y:-2, w:6, h:6, fill:{ color:C.DARK_BLUE, transparency:60 } });
  s.addShape(pres.shapes.OVAL, { x:-2, y:4, w:5, h:5, fill:{ color:C.DARK_BLUE, transparency:60 } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.8, y:3.4, w:3.5, h:0.06, fill:{ color:C.GOLD } });
  s.addText("BERITA DAERAH KOTA BEKASI", { x:0.8, y:1.5, w:10, h:0.4, fontSize:14, color:C.GOLD, bold:true, fontFace:"Calibri" });
  s.addText("TERIMA KASIH", { x:0.8, y:2.3, w:11, h:1.5, fontSize:48, bold:true, color:C.WHITE, fontFace:"Calibri" });
  s.addText("Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024\nTentang Pengelolaan Pajak Reklame", { x:0.8, y:4.0, w:11, h:0.8, fontSize:15, color:C.ICE_BLUE, fontFace:"Calibri" });
  s.addText("Sumber: https://jdih.bekasikota.go.id", { x:0.8, y:5.0, w:11, h:0.4, fontSize:11, color:C.MUTED, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.GOLD } });
})();

// ═══════════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════════
const path = require("path");
const out = path.resolve(__dirname, "Perwal_Bekasi_51_2024_Pajak_Reklame_Ringkas.pptx");
pres.writeFile({ fileName: out })
  .then(() => { console.log(`✅ Ringkasan PPT: ${out} (${pres.slides.length} slide)`); })
  .catch(e => console.error("❌", e));
