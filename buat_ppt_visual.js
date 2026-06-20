#!/usr/bin/env node
/**
 * PPT RINGKAS — Visual Rich Edition
 * Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024
 * Tentang Pengelolaan Pajak Reklame
 *
 * Dengan: pola dekoratif, ikon shape, kartu warna, data callout, timeline
 */

const pptxgen = require("pptxgenjs");

// ═══════════════════════════════════════════════
// PALETTE
// ═══════════════════════════════════════════════
const C = {
  NAVY:"0A1628", DARK_BLUE:"12294A", MID_BLUE:"1B3A6B",
  ACCENT_BLUE:"2D7DD2", GOLD:"D4A017", WARM_GOLD:"B8860B",
  WHITE:"FFFFFF", OFF_WHITE:"F0F2F5", ICE_BLUE:"E8EDF5",
  DARK_TEXT:"1A1A2E", GRAY_TEXT:"6B7288", MUTED:"9CA3AF",
  TEAL:"0D9488", RED_ACCENT:"DC2626",
  SOFT_RED:"FEE2E2", SOFT_GREEN:"D1FAE5", SOFT_YELLOW:"FEF3C7",
  SOFT_PURPLE:"EDE9FE", SOFT_ORANGE:"FFEDD5",
  BG_CARD:"FAFBFC",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Pemerintah Kota Bekasi";
pres.title = "Perwal Bekasi No 51/2024 - Pajak Reklame";

let sn = 0;
const pn = (s) => { sn++; s.addText(String(sn),{x:12.2,y:7.0,w:0.8,h:0.35,fontSize:9,color:C.GRAY_TEXT,align:"right",fontFace:"Calibri"}); };

// ═══════════════════════════════════════════════
// DECORATIVE HELPERS
// ═══════════════════════════════════════════════

const bgPattern = (s, color=C.OFF_WHITE) => {
  s.background = { color };
  // Subtle corner decoration
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:13.333, h:0.04, fill:{ color:C.GOLD } });
};

const topBar = (s, title, subtitle="") => {
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.04, w:13.333, h:0.85, fill:{ color:C.NAVY } });
  // Gold accent dots pattern
  for (let i=0; i<8; i++) {
    s.addShape(pres.shapes.OVAL, { x:0.4+i*0.16, y:0.2, w:0.06, h:0.06, fill:{ color:C.GOLD } });
  }
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.4, w:0.08, h:0.4, fill:{ color:C.GOLD } });
  s.addText(title, { x:0.75, y:0.12, w:11, h:0.45, fontSize:20, bold:true, color:C.WHITE, fontFace:"Calibri" });
  if (subtitle) s.addText(subtitle, { x:0.75, y:0.52, w:10, h:0.3, fontSize:10, color:C.MUTED, fontFace:"Calibri" });
};

const footerBar = (s) => {
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.NAVY } });
  // Small gold dots in footer
  for (let i=0; i<5; i++) {
    s.addShape(pres.shapes.OVAL, { x:0.4+i*0.14, y:7.3, w:0.04, h:0.04, fill:{ color:C.GOLD } });
  }
};

const iconCircle = (s, x, y, size, color, symbol, symbolColor=C.WHITE) => {
  s.addShape(pres.shapes.OVAL, { x, y, w:size, h:size, fill:{ color } });
  if (symbol) s.addText(symbol, { x, y:y+size*0.05, w:size, h:size, fontSize:size*0.5, color:symbolColor, align:"center", valign:"middle", fontFace:"Segoe UI Symbol" });
};

const decoLine = (s, x, y, w, color=C.GOLD, h=0.04) => {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{ color } });
};

const cornerDeco = (s, x, y, color=C.MID_BLUE) => {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.6, h:0.04, fill:{ color } });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.04, h:0.6, fill:{ color } });
};

// ─── COVER ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.NAVY };
  // Big decorative circles
  s.addShape(pres.shapes.OVAL, { x:9, y:-2, w:7, h:7, fill:{ color:"0D1F3C", transparency:50 } });
  s.addShape(pres.shapes.OVAL, { x:10.5, y:4.5, w:4.5, h:4.5, fill:{ color:"0D1F3C", transparency:50 } });
  s.addShape(pres.shapes.OVAL, { x:-1.5, y:-1, w:3, h:3, fill:{ color:"0D1F3C", transparency:50 } });
  
  // Gold dot pattern top
  for (let i=0; i<12; i++) {
    s.addShape(pres.shapes.OVAL, { x:0.8+i*0.2, y:0.3, w:0.06, h:0.06, fill:{ color:C.GOLD } });
  }
  
  // Diamond pattern
  for (let i=0; i<3; i++) {
    s.addShape(pres.shapes.RECTANGLE, { x:0.8+i*0.4, y:0.5, w:0.2, h:0.02, fill:{ color:C.WARM_GOLD, transparency:60 } });
  }

  decoLine(s, 0.8, 2.5, 4);
  
  s.addText("BERITA DAERAH", { x:0.8, y:0.6, w:5, h:0.35, fontSize:12, color:C.GOLD, bold:true, fontFace:"Calibri" });
  s.addText("KOTA BEKASI", { x:0.8, y:0.9, w:5, h:0.35, fontSize:14, color:C.WHITE, bold:true, fontFace:"Calibri" });
  
  s.addText("PERATURAN WALI KOTA BEKASI", { x:0.8, y:1.8, w:11, h:0.5, fontSize:22, bold:true, color:C.WHITE, fontFace:"Calibri" });
  s.addText("NOMOR 51 TAHUN 2024", { x:0.8, y:2.2, w:11, h:0.4, fontSize:16, color:C.GOLD, bold:true, fontFace:"Calibri" });
  s.addText("TENTANG\nPENGELOLAAN PAJAK REKLAME", { x:0.8, y:3.2, w:11, h:1.8, fontSize:40, bold:true, color:C.WHITE, fontFace:"Calibri" });

  // Bottom info box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.8, y:5.3, w:7, h:1.0, fill:{ color:"0D1F3C" } });
  s.addText("Pemerintah Kota Bekasi  |  20 Desember 2024", { x:1.0, y:5.35, w:6.5, h:0.4, fontSize:12, color:C.ICE_BLUE, fontFace:"Calibri" });
  s.addText("Berlaku sejak diundangkan", { x:1.0, y:5.65, w:6, h:0.4, fontSize:11, color:C.MUTED, fontFace:"Calibri" });
  
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.GOLD } });
})();

// ─── DAFTAR ISI ───
(() => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, "DAFTAR ISI RINGKAS", "Perwal Bekasi No 51/2024 · Pengelolaan Pajak Reklame");
  
  const toc = [
    ["1", "Ketentuan Umum & Definisi", C.ACCENT_BLUE],
    ["2", "Objek, Subjek & Wajib Pajak", C.TEAL],
    ["3", "Masa Pajak & Tahun Pajak", C.WARM_GOLD],
    ["4", "Pendaftaran & Pendataan WP", C.MID_BLUE],
    ["5", "Nilai Sewa Reklame (NSR)", C.RED_ACCENT],
    ["6", "Perhitungan & Tarif Pajak", C.ACCENT_BLUE],
    ["7", "Penetapan & Pembayaran", C.TEAL],
    ["8", "Pembetulan, Keberatan & Banding", C.WARM_GOLD],
    ["9", "Pemeriksaan, Penagihan & Penghapusan", C.MID_BLUE],
    ["10", "Keringanan, Kemudahan & Penghargaan", C.ACCENT_BLUE],
    ["11", "Ketentuan Penutup", C.TEAL],
  ];
  
  toc.forEach(([num, ttl, clr], i) => {
    const col = i < 6 ? 0 : 1;
    const row = i < 6 ? i : i - 6;
    const x = col === 0 ? 0.5 : 6.8;
    const y = 1.2 + row * 0.9;
    
    // Card background
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:5.8, h:0.7, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
    
    // Number circle
    iconCircle(s, x+0.12, y+0.1, 0.5, clr, num);
    s.addText(ttl, { x:x+0.75, y:y+0.08, w:4.5, h:0.55, fontSize:14, color:C.DARK_TEXT, fontFace:"Calibri", valign:"middle" });
  });
  
  footerBar(s); pn(s);
})();

// ═══════════════════════════════════════════════
// SECTION SLIDE (decorative)
// ═══════════════════════════════════════════════
const section = (title, sub="") => {
  const s = pres.addSlide();
  s.background = { color:C.NAVY };
  
  // Corner decorations
  cornerDeco(s, 0.5, 0.5, C.GOLD);
  cornerDeco(s, 12.2, 0.5, C.GOLD);
  cornerDeco(s, 0.5, 6.5, C.GOLD);
  cornerDeco(s, 12.2, 6.5, C.GOLD);
  
  // Dot lines
  for (let i=0; i<20; i++) {
    s.addShape(pres.shapes.OVAL, { x:0.8+i*0.18, y:6.8, w:0.05, h:0.05, fill:{ color:C.GOLD, transparency:50 } });
  }
  
  // Large subtle circle
  s.addShape(pres.shapes.OVAL, { x:-1, y:1, w:5, h:5, fill:{ color:"0D1F3C", transparency:60 } });
  s.addShape(pres.shapes.OVAL, { x:9, y:-1, w:4.5, h:4.5, fill:{ color:"0D1F3C", transparency:60 } });
  
  decoLine(s, 0.8, 1.8, 2.5);
  s.addText(title, { x:0.8, y:2.0, w:11, h:1.8, fontSize:34, bold:true, color:C.WHITE, fontFace:"Calibri" });
  if (sub) s.addText(sub, { x:0.8, y:3.8, w:10, h:0.5, fontSize:13, color:C.MUTED, fontFace:"Calibri" });
  
  // Icon in circle
  const icons = ["📋","🎯","📅","📝","💰","🧮","💳","⚖️","🔍","🎁","🔚"];
  const idx = sn % icons.length;
  iconCircle(s, 11.5, 5.0, 0.7, C.GOLD, icons[sn % icons.length], C.NAVY);
  
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.GOLD } });
  pn(s);
};

// ─── HELPERS: Card (2 or 3 columns) ───
const cardSlide = (title, cards, opts={}) => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, title, opts.subtitle);
  const n = cards.length;
  const w = n === 3 ? 3.8 : (n === 2 ? 5.8 : 11.8);
  const gap = 0.3;
  const tw = n * w + (n-1) * gap;
  const sx = (13.333 - tw) / 2;
  const colors = [C.ACCENT_BLUE, C.TEAL, C.WARM_GOLD, C.MID_BLUE, C.RED_ACCENT];
  
  cards.forEach((c, i) => {
    const cx = sx + i*(w+gap);
    const clr = c.color || colors[i%colors.length];
    
    // Card bg
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:1.2, w, h:5.3, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:4, offset:1.5, angle:135, color:"000000", opacity:0.08 } });
    
    // Top accent
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:1.2, w, h:0.06, fill:{ color:clr } });
    
    // Icon
    if (c.icon) iconCircle(s, cx+0.2, y=1.4, 0.45, clr, c.icon);
    const ty = c.icon ? 1.4 : 1.3;
    
    s.addText(c.title, { x:cx+0.2, y:ty+0.55, w:w-0.4, h:0.35, fontSize:15, bold:true, color:clr, fontFace:"Calibri" });
    
    const items = c.items.map((t,j) => ({
      text:t, options:{ bullet:true, breakLine:j<c.items.length-1, fontSize:12, color:C.DARK_TEXT, fontFace:"Calibri", paraSpaceAfter:3 }
    }));
    s.addText(items, { x:cx+0.2, y:ty+0.95, w:w-0.4, h:4.0, valign:"top", margin:0 });
  });
  footerBar(s); pn(s);
};

// ─── HELPERS: Two Column ───
const twoColSlide = (title, L, R, opts={}) => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, title, opts.subtitle);
  
  // Left column card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.5, y:1.15, w:5.8, h:5.4, fill:{ color:C.WHITE },
    shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.15, w:5.8, h:0.06, fill:{ color:C.ACCENT_BLUE } });
  
  const mkL = L.items.map((t,j) => ({
    text:t, options:{ bullet:L.bullets !== false, breakLine:j<L.items.length-1, fontSize:13, bold:j===0, color:C.DARK_TEXT, fontFace:"Calibri", paraSpaceAfter:3 }
  }));
  s.addText(mkL, { x:0.65, y:1.35, w:5.5, h:5.0, valign:"top", margin:0 });
  
  // Right column card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:6.8, y:1.15, w:5.8, h:5.4, fill:{ color:C.WHITE },
    shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
  s.addShape(pres.shapes.RECTANGLE, { x:6.8, y:1.15, w:5.8, h:0.06, fill:{ color:C.TEAL } });
  
  const mkR = R.items.map((t,j) => ({
    text:t, options:{ bullet:R.bullets !== false, breakLine:j<R.items.length-1, fontSize:13, bold:j===0, color:C.DARK_TEXT, fontFace:"Calibri", paraSpaceAfter:3 }
  }));
  s.addText(mkR, { x:6.95, y:1.35, w:5.5, h:5.0, valign:"top", margin:0 });
  
  footerBar(s); pn(s);
};

// ─── HELPERS: Table ───
const tblSlide = (title, headers, rows, opts={}) => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, title, opts.subtitle);
  
  const hRow = headers.map(h => ({ text:h, options:{ bold:true, color:C.WHITE, fill:{ color:C.NAVY }, fontSize:11, fontFace:"Calibri", align:"center", valign:"middle" } }));
  const dRows = rows.map((r,i) => r.map((c,j) => ({ text:String(c), options:{ fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri", fill:{ color:i%2===0?C.ICE_BLUE:C.WHITE }, align:j===0?"left":"center", valign:"middle" } })));
  
  s.addTable([hRow, ...dRows], { x:0.8, y:1.15, w:11.7, colW:opts.colW, border:{ pt:0.5, color:C.MUTED }, rowH:0.45, margin:[2,6,2,6] });
  footerBar(s); pn(s);
};

// ─── HELPERS: Big Number Callout ───
const calloutSlide = (title, callouts, opts={}) => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, title, opts.subtitle);
  const n = callouts.length;
  const w = n === 4 ? 2.8 : (n === 3 ? 3.8 : 5.8);
  const gap = 0.3;
  const tw = n * w + (n-1) * gap;
  const sx = (13.333 - tw) / 2;
  
  callouts.forEach((c, i) => {
    const cx = sx + i*(w+gap);
    const clr = c.color || C.ACCENT_BLUE;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:1.3, w, h:2.3, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:1.3, w, h:0.06, fill:{ color:clr } });
    s.addText(c.num, { x:cx, y:1.6, w, h:0.8, fontSize:36, bold:true, color:clr, align:"center", fontFace:"Calibri" });
    s.addText(c.label, { x:cx+0.15, y:2.4, w:w-0.3, h:0.6, fontSize:12, color:C.GRAY_TEXT, align:"center", valign:"middle", fontFace:"Calibri" });
  });
  
  if (opts.extra) s.addText(opts.extra, { x:0.6, y:3.8, w:12, h:2.5, fontSize:13, color:C.DARK_TEXT, fontFace:"Calibri", valign:"top" });
  footerBar(s); pn(s);
};

// ═══════════════════════════════════════════════
// CONTENT SLIDES
// ═══════════════════════════════════════════════

// BAB I
section("BAB I: KETENTUAN UMUM", "Pasal 1 — Definisi Kunci");

(() => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, "Definisi Penting");
  
  const defs = [
    ["🏛️", "Daerah", "Kota Bekasi"],
    ["📊", "Bapenda", "Badan Pendapatan Daerah Kota Bekasi"],
    ["📢", "Reklame", "Media komersial untuk promosi & pengenalan"],
    ["💰", "Pajak Reklame", "Pajak atas penyelenggaraan reklame"],
    ["📐", "NSR", "Nilai Sewa Reklame (dasar pengenaan pajak)"],
    ["🆔", "NPWPD", "Nomor Pokok Wajib Pajak Daerah"],
    ["👤", "Wajib Pajak", "Orang pribadi/badan dgn hak & kewajiban pajak"],
  ];
  
  defs.forEach(([icon, term, def], i) => {
    const y = 1.15 + i * 0.78;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.6, y, w:12, h:0.65, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:2, offset:0.5, angle:135, color:"000000", opacity:0.04 } });
    iconCircle(s, 0.75, y+0.08, 0.5, C.ICE_BLUE, icon, C.NAVY);
    s.addText(term, { x:1.4, y:y+0.04, w:3, h:0.3, fontSize:13, bold:true, color:C.DARK_TEXT, fontFace:"Calibri" });
    s.addText(def, { x:1.4, y:y+0.32, w:9, h:0.3, fontSize:11, color:C.GRAY_TEXT, fontFace:"Calibri" });
  });
  footerBar(s); pn(s);
})();

// BAB II
section("BAB II: OBJEK, SUBJEK & WAJIB PAJAK", "Pasal 2–4");

cardSlide("Objek Pajak Reklame", [
  { icon:"📋", title:"10 Jenis Reklame", color:C.ACCENT_BLUE, items:[
    "Papan/Billboard", "Videotron/Megatron", "Kain (Spanduk/Umbul/Baliho)", "Melekat/Stiker", "Selebaran", "Berjalan (Kendaraan)", "Udara (Balon Gas)", "Apung", "Film/Slide", "Peragaan",
  ]},
  { icon:"🚫", title:"Dikecualikan", color:C.TEAL, items:[
    "Internet, TV, radio, media cetak", "Label/merek produk", "Nama usaha ≤ 1 m² di tempat", "Reklame Pemerintah/Pemda", "Tempat ibadah & panti asuhan", "Kegiatan sosial/keagamaan ≤ 30 hr", "Kegiatan politik (masa kampanye)", "Olahraga KONI ≤ 30 hari",
  ]},
]);

cardSlide("Subjek & Wajib Pajak", [
  { icon:"👤", title:"Subjek Pajak (Pasal 3)", color:C.ACCENT_BLUE, items:[
    "Orang pribadi atau Badan", "yang menggunakan Reklame",
  ]},
  { icon:"✋", title:"Wajib Pajak (Pasal 4)", color:C.TEAL, items:[
    "Orang pribadi atau Badan", "yang menyelenggarakan Reklame", "Jika pihak ketiga → pihak ketiga", "menjadi Wajib Pajak",
  ]},
]);

// BAB III
section("BAB III: MASA PAJAK & TAHUN PAJAK", "Pasal 5");

calloutSlide("Masa Pajak Reklame", [
  { num:"12", label:"Bulan (Permanen)", color:C.ACCENT_BLUE },
  { num:"30", label:"Hari (Insidentil)", color:C.TEAL },
  { num:"1", label:"Tahun Pajak", color:C.WARM_GOLD },
  { num:"1", label:"Bulan (Bagian Tahun)", color:C.MID_BLUE },
], { extra:"Masa Pajak Permanen: 12 bulan atau sesuai masa penayangan.  ·  Masa Pajak Insidentil: per hari, maks. 30 hari.  ·  Tahun Pajak: 1 tahun kalender (atau sesuai tahun buku WP)." });

// BAB IV
section("BAB IV: PENDAFTARAN & PENDATAAN WP", "Pasal 6–8");

twoColSlide("Pendaftaran & Pendataan WP", {
  items:["📝 PENDAFTARAN (Pasal 6)","","WP wajib daftarkan diri & objek pajak","Form: ambil/online/dikirim petugas","Lampirkan: KTP, NPWP, Akta, NIB","Bapenda terbitkan NPWPD","Jika tak daftar → NPWPD jabatan","Juga: NOPD & nomor registrasi"],
}, {
  items:["📋 PENDATAAN & NONAKTIF (Pasal 7–8)","","Pendataan: Bapenda data WP & objek","Termasuk data geografis","Bisa kerjasama dgn instansi lain","Penonaktifan: WP tak penuhi syarat","Keputusan maks. 3 bulan","Syarat: tanpa tunggakan & keberatan"],
});

// BAB V
section("BAB V: NILAI SEWA REKLAME", "Pasal 9");

(() => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, "Faktor Penentu NSR");
  
  const factors = [
    ["1", "Jenis Reklame", C.ACCENT_BLUE],
    ["2", "Bahan", C.TEAL],
    ["3", "Lokasi (Kelas Jalan)", C.WARM_GOLD],
    ["4", "Waktu (detik)", C.MID_BLUE],
    ["5", "Jangka Waktu (hari)", C.ACCENT_BLUE],
    ["6", "Jumlah Media", C.TEAL],
    ["7", "Ukuran (m²)", C.WARM_GOLD],
  ];
  
  factors.forEach(([n, label, clr], i) => {
    const col = i % 4;
    const row = Math.floor(i/4);
    const x = 0.6 + col * 3.1;
    const y = 1.2 + row * 1.5;
    
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:2.8, h:1.2, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
    iconCircle(s, x+0.15, y+0.2, 0.55, clr, n);
    s.addText(label, { x:x+0.8, y:y+0.2, w:1.8, h:0.8, fontSize:14, color:C.DARK_TEXT, fontFace:"Calibri", valign:"middle" });
  });
  
  // Bottom note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.6, y:4.2, w:12, h:1.8, fill:{ color:C.OFF_WHITE } });
  s.addText("Klasifikasi Kelas Jalan:", { x:0.8, y:4.3, w:5, h:0.35, fontSize:13, bold:true, color:C.NAVY, fontFace:"Calibri" });
  s.addText("🏛️ Kelas Jalan Khusus: Tol | Premium 1 (Jl. A. Yani, Juanda, Sudirman, dll) | Premium 2 (Jl. Narogong, Jatiwaringin, dll)", { x:0.8, y:4.65, w:11.5, h:0.35, fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri" });
  s.addText("🚗 Kelas Jalan I (Kendali Ketat): Lebar > 3 m, pusat pelayanan", { x:0.8, y:5.0, w:11, h:0.35, fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri" });
  s.addText("🏡 Kelas Jalan II (Kendali Sedang): Lebar ≤ 3 m, jalan lingkungan", { x:0.8, y:5.35, w:11, h:0.35, fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri" });
  
  footerBar(s); pn(s);
})();

// BAB VI
section("BAB VI: PERHITUNGAN & TARIF PAJAK", "Pasal 10");

calloutSlide("Rumus Dasar", [
  { num:"×", label:"Pajak = Tarif × NSR", color:C.ACCENT_BLUE },
  { num:"📐", label:"NSR = Nilai × Ukuran × Jml × Waktu", color:C.TEAL },
  { num:"50%", label:"Indoor = 50% NSR", color:C.WARM_GOLD },
  { num:"+20%", label:"Tinggi >15m = +20%", color:C.RED_ACCENT },
], { subtitle:"Pasal 10 · Ketentuan: Tembakau/Miras +50% | Perubahan naskah dikecualikan" });

tblSlide("NSR — Papan/Billboard", ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"], [
  ["Kelas Jalan Khusus", "Jalan Tol", "23.575"],
  ["Kelas Jalan Khusus", "Premium 1", "16.100"],
  ["Kelas Jalan Khusus", "Premium 2", "14.950"],
  ["Kelas Jalan I", "Kendali Ketat", "13.225"],
  ["Kelas Jalan II", "Kendali Sedang", "11.500"],
], { colW:[3,6,2.7] });

tblSlide("NSR — Megatron/Videotron", ["Kelas Jalan", "Zona", "NSR (/30 dtk)", "NSR (/m²/thn)"], [
  ["Kelas Jalan Khusus", "Jalan Tol", "Rp 17,25", "13.599.900"],
  ["Kelas Jalan Khusus", "Premium 1", "Rp 13,80", "10.879.920"],
  ["Kelas Jalan Khusus", "Premium 2", "Rp 9,20", "7.253.280"],
  ["Kelas Jalan I", "Kendali Ketat", "Rp 8,05", "6.346.620"],
  ["Kelas Jalan II", "Kendali Sedang", "Rp 5,75", "4.533.300"],
], { colW:[3,4,2.7,2.7] });

tblSlide("NSR — Kain (Spanduk/Umbul/Baliho)", ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"], [
  ["Kelas Jalan Khusus", "Jalan Tol", "30.000"],
  ["Kelas Jalan Khusus", "Premium 1", "30.000"],
  ["Kelas Jalan Khusus", "Premium 2", "25.000"],
  ["Kelas Jalan I", "Kendali Ketat", "20.000"],
  ["Kelas Jalan II", "Kendali Sedang", "19.000"],
], { colW:[3,6,2.7] });

cardSlide("NSR — Jenis Lainnya", [
  { icon:"🏷️", title:"Stiker", color:C.ACCENT_BLUE, items:["Rp 7,5/cm²", "Min. Rp 750.000/kali"] },
  { icon:"🧱", title:"Melekat", color:C.TEAL, items:["Rp 750.000/m²/tahun"] },
  { icon:"📄", title:"Selebaran", color:C.WARM_GOLD, items:["Rp 600/lembar", "Min. Rp 6.000.000/kali"] },
  { icon:"🚌", title:"Berjalan", color:C.RED_ACCENT, items:["Rp 6.000/m²/hari", "Termasuk kendaraan"] },
]);

cardSlide("NSR — Jenis Lainnya (lanjutan)", [
  { icon:"🎈", title:"Udara", color:C.ACCENT_BLUE, items:["Rp 2.400.000/sekali", "Maks. 1 bulan"] },
  { icon:"🌊", title:"Apung", color:C.TEAL, items:["Rp 600.000/sekali", "Maks. 1 bulan"] },
  { icon:"🎬", title:"Film/Slide", color:C.WARM_GOLD, items:["Rp 12.000/15 detik"] },
  { icon:"🎭", title:"Peragaan", color:C.RED_ACCENT, items:["Rp 480.000/penyelenggaraan"] },
]);

// BAB VII
section("BAB VII: PENETAPAN, TAGIHAN & PEMBAYARAN", "Pasal 11–14");

calloutSlide("Penetapan & Pembayaran", [
  { num:"5", label:"Tahun (batas penetapan)", color:C.ACCENT_BLUE },
  { num:"1%", label:"Bunga/bln (keterlambatan)", color:C.RED_ACCENT },
  { num:"24", label:"Maks. bulan bunga", color:C.WARM_GOLD },
  { num:"30", label:"Hari (lunas STPD)", color:C.TEAL },
], { subtitle:"Pasal 11–14" });

(() => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, "Alur Penetapan → Pembayaran");
  
  // Timeline flow
  const steps = [
    ["1", "SKPD", "Diterbitkan\n5 tahun", C.ACCENT_BLUE],
    ["2", "Bayar", "1 bulan\nsejak SKPD", C.TEAL],
    ["3", "Telat?", "Bunga 1%/bln\n+ STPD", C.WARM_GOLD],
    ["4", "STPD", "Lunas\n≤ 30 hari", C.RED_ACCENT],
  ];
  
  steps.forEach(([n, title, desc, clr], i) => {
    const x = 0.8 + i * 3.1;
    // Card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.4, w:2.8, h:2.5, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
    // Top bar
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.4, w:2.8, h:0.06, fill:{ color:clr } });
    // Circle number
    iconCircle(s, x+1.0, y=1.6, 0.6, clr, n, C.WHITE);
    s.addText(title, { x, y:2.4, w:2.8, h:0.4, fontSize:16, bold:true, color:clr, align:"center", fontFace:"Calibri" });
    s.addText(desc, { x:x+0.1, y:2.8, w:2.6, h:0.8, fontSize:11, color:C.GRAY_TEXT, align:"center", fontFace:"Calibri" });
    
    // Arrow between steps
    if (i < 3) {
      s.addText("→", { x:x+2.8, y:2.1, w:0.3, h:0.5, fontSize:20, color:C.GOLD, align:"center", fontFace:"Calibri" });
    }
  });
  
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.8, y:4.3, w:11.7, h:1.5, fill:{ color:C.ICE_BLUE } });
  s.addText("💡 Jatuh tempo pembayaran: 1 bulan sejak tanggal pengiriman SKPD. Pembayaran via Bank Persepsi atau tempat lain yang ditunjuk. Stiker sebagai tanda bukti pembayaran.", { x:1.0, y:4.4, w:11.2, h:1.2, fontSize:12, color:C.DARK_TEXT, fontFace:"Calibri", valign:"middle" });
  
  footerBar(s); pn(s);
})();

// BAB VIII
section("BAB VIII: PEMBETULAN, KEBERATAN & BANDING", "Pasal 15–20, 29–33");

twoColSlide("Pembetulan (Pasal 15–20)", {
  items:["✏️ PEMBETULAN KETETAPAN","","Kesalahan tulis (nama, alamat, NPWPD)","Kesalahan hitung (jumlah, tarif)","Kekeliruan penerapan aturan","1 permohonan = 1 ketetapan","Keputusan maks. 6 bulan","> 6 bulan tanpa putusan → dikabulkan"],
}, {
  items:["🔄 DAPAT DILAKUKAN BERULANG","","Pasal 19: jabatan","Pasal 20: berulang jika masih salah","","Jenis keputusan:","kabul (tambah/kurang/hapus)","batal | tolak"],
});

twoColSlide("Keberatan & Banding (Pasal 29–33)", {
  items:["⚖️ KEBERATAN","","Objek: SKPD, SKPDKB, SKPDKBT, dll","Maks. 3 bulan sejak SKPD","Sudah bayar min. yg disetujui","Keputusan maks. 12 bulan","Jika ditolak: denda 30%","Jika dikabulkan: + bunga 0,6%/bln"],
}, {
  items:["🏛️ BANDING","","Objek: SK Keberatan","Ke badan peradilan pajak","Maks. 3 bulan sejak keputusan","Menangguhkan kewajiban bayar","Jika ditolak: denda 60%","Jika dikabulkan: + bunga 0,6%/bln"],
});

// BAB IX
section("BAB IX: PEMERIKSAAN, PENAGIHAN & PENGHAPUSAN", "Pasal 22–26");

cardSlide("Pemeriksaan & Penagihan", [
  { icon:"🔍", title:"Pemeriksaan (Psl 22–23)", color:C.ACCENT_BLUE, items:[
    "Kepala Bapenda berwenang periksa", "Uji kepatuhan WP", "WP wajib: buka buku/dokumen", "Beri akses tempat & keterangan", "Jika tidak → pajak ditetapkan jabatan",
  ]},
  { icon:"📬", title:"Penagihan (Psl 24)", color:C.TEAL, items:[
    "Dasar: SKPD, SKPDKB, SKPDKBT", "STPD, SK Pembetulan/Keberatan", "Putusan Banding",
  ]},
  { icon:"⏳", title:"Kedaluwarsa (Psl 25)", color:C.WARM_GOLD, items:[
    "5 tahun sejak pajak terutang", "Tertangguh jika ada:", "Surat Teguran/Paksa", "Pengakuan utang dari WP",
  ]},
]);

(() => {
  const s = pres.addSlide(); bgPattern(s); topBar(s, "Penghapusan Piutang Pajak (Pasal 26)");
  
  // Flow chart boxes
  const steps = [
    ["1", "Penelitian", "Bapenda", C.ACCENT_BLUE],
    ["2", "Tim Peneliti", "Keputusan Wali Kota", C.TEAL],
    ["3", "Daftar Usulan", "Kepala Bapenda", C.WARM_GOLD],
    ["4", "Penetapan", "Keputusan Wali Kota", C.MID_BLUE],
  ];
  
  steps.forEach(([n, title, desc, clr], i) => {
    const x = 0.6 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.5, w:2.8, h:2.0, fill:{ color:C.WHITE },
      shadow:{ type:"outer", blur:3, offset:1, angle:135, color:"000000", opacity:0.06 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.5, w:2.8, h:0.06, fill:{ color:clr } });
    iconCircle(s, x+1.0, 1.7, 0.5, clr, n, C.WHITE);
    s.addText(title, { x, y:2.3, w:2.8, h:0.35, fontSize:13, bold:true, color:clr, align:"center", fontFace:"Calibri" });
    s.addText(desc, { x:x+0.1, y:2.65, w:2.6, h:0.5, fontSize:10, color:C.GRAY_TEXT, align:"center", fontFace:"Calibri" });
    if (i<3) s.addText("→", { x:x+2.8, y:2.0, w:0.35, h:0.5, fontSize:18, color:C.GOLD, align:"center", fontFace:"Calibri" });
  });
  
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.6, y:3.8, w:12, h:1.8, fill:{ color:C.OFF_WHITE } });
  s.addText("📋 Syarat Penghapusan:", { x:0.8, y:3.9, w:5, h:0.3, fontSize:12, bold:true, color:C.NAVY, fontFace:"Calibri" });
  s.addText("• Piutang tidak mungkin ditagih lagi karena kedaluwarsa", { x:0.8, y:4.2, w:11, h:0.3, fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri" });
  s.addText("• Ada koordinasi dengan Inspektorat Daerah", { x:0.8, y:4.5, w:11, h:0.3, fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri" });
  s.addText("• Dibuktikan dengan dokumen pelaksanaan penagihan", { x:0.8, y:4.8, w:11, h:0.3, fontSize:11, color:C.DARK_TEXT, fontFace:"Calibri" });
  
  footerBar(s); pn(s);
})();

// BAB X
section("BAB X: KERINGANAN, KEMUDAHAN & PENGHARGAAN", "Pasal 27–28, 34–35");

cardSlide("Fasilitas & Penghargaan", [
  { icon:"🎯", title:"Keringanan (Psl 27)", color:C.ACCENT_BLUE, items:[
    "Keringanan/Pengurangan", "Pembebasan/Penundaan", "Atas pokok &/ sanksi pajak", "WP: likuiditas rendah", "Objek: bencana/kebakaran",
  ]},
  { icon:"🤝", title:"Kemudahan (Psl 28)", color:C.TEAL, items:[
    "Perpanjangan waktu bayar", "Angsuran (maks. 24 bln)", "Bunga 0,6%/bln", "Keadaan kahar:", "bencana, kebakaran, wabah", "kerusuhan, dll",
  ]},
  { icon:"🏆", title:"Penghargaan (Psl 34–35)", color:C.WARM_GOLD, items:[
    "WP Taat Pajak", "Bayar tepat waktu ≥ 1 thn", "Tanpa tunggakan 3 thn", "Kontribusi signifikan", "Piagam/Hadiah (APBD)",
  ]},
]);

// BAB XI
section("BAB XI: KETENTUAN PENUTUP", "Pasal 36–37");

twoColSlide("Pencabutan & Mulai Berlaku", {
  items:["❌ PERATURAN DICABUT (Pasal 36)","","Perwal No. 48/2012","Petunjuk Pelaksanaan","Perda No. 14/2012","","Perwal No. 52/2013","Perubahan Perwal 48/2012"],
}, {
  items:["✅ MULAI BERLAKU (Pasal 37)","","Sejak diundangkan","20 Desember 2024","","━━━━━━━━━━━━━━━","Pj. WALI KOTA BEKASI,","ttd","R. GANI MUHAMAD"],
  bullets:false,
});

// ─── CLOSING ───
(() => {
  const s = pres.addSlide();
  s.background = { color:C.NAVY };
  s.addShape(pres.shapes.OVAL, { x:9, y:-2, w:6, h:6, fill:{ color:"0D1F3C", transparency:50 } });
  s.addShape(pres.shapes.OVAL, { x:-2, y:4, w:5, h:5, fill:{ color:"0D1F3C", transparency:50 } });
  
  // Gold dots
  for (let i=0; i<10; i++) s.addShape(pres.shapes.OVAL, { x:0.8+i*0.2, y:0.3, w:0.06, h:0.06, fill:{ color:C.GOLD } });
  cornerDeco(s, 0.5, 0.5, C.GOLD);
  cornerDeco(s, 12.2, 0.5, C.GOLD);
  
  decoLine(s, 0.8, 3.5, 3.5);
  
  s.addText("BERITA DAERAH KOTA BEKASI", { x:0.8, y:1.5, w:10, h:0.4, fontSize:14, color:C.GOLD, bold:true, fontFace:"Calibri" });
  s.addText("TERIMA KASIH", { x:0.8, y:2.3, w:11, h:1.5, fontSize:48, bold:true, color:C.WHITE, fontFace:"Calibri" });
  s.addText("Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024", { x:0.8, y:4.0, w:11, h:0.4, fontSize:14, color:C.ICE_BLUE, fontFace:"Calibri" });
  s.addText("Tentang Pengelolaan Pajak Reklame", { x:0.8, y:4.35, w:11, h:0.4, fontSize:14, color:C.ICE_BLUE, fontFace:"Calibri" });
  s.addText("Sumber: https://jdih.bekasikota.go.id", { x:0.8, y:5.1, w:11, h:0.4, fontSize:11, color:C.MUTED, fontFace:"Calibri" });
  
  // Bottom dots
  for (let i=0; i<20; i++) s.addShape(pres.shapes.OVAL, { x:0.8+i*0.18, y:6.7, w:0.05, h:0.05, fill:{ color:C.GOLD, transparency:50 } });
  
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:7.25, w:13.333, h:0.25, fill:{ color:C.GOLD } });
})();

// ═══════════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════════
const path = require("path");
const out = path.resolve(__dirname, "Perwal_Bekasi_51_2024_Pajak_Reklame.pptx");
pres.writeFile({ fileName: out })
  .then(() => { console.log(`✅ PPT VISUAL RICH: ${out} (${pres.slides.length} slide)`); })
  .catch(e => console.error("❌", e));
