#!/usr/bin/env node
/**
 * PPT V3 — Premium Presentation with PptxGenJS
 * Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024
 * Tentang Pengelolaan Pajak Reklame
 *
 * Mengikuti best practice dari Anthropic SKILL.md:
 * - ✅ NO accent lines under titles (hallmark AI-generated slides)
 * - ✅ Varied layouts across slides
 * - ✅ Proper font sizing (title 36-44pt, body 14-16pt)
 * - ✅ Left-align body text
 * - ✅ Dark/light sandwich (dark cover+closing, light content)
 * - ✅ Each slide needs visual elements
 * - ✅ Topic-informed color palette
 * - ✅ Dominance: one color 60-70%
 * - ✅ No "#" in hex colors
 * - ✅ bullet: true instead of unicode bullets
 * - ✅ breakLine: true for multi-line
 * - ✅ paraSpaceAfter instead of lineSpacing with bullets
 * - ✅ Factory functions for options (no reuse)
 * - ✅ margin: 0 for text-shape alignment
 */

const pptxgen = require("pptxgenjs");

// ═══════════════════════════════════════════════
// PALETTE — Government Authority
// ═══════════════════════════════════════════════
const C = {
  NAVY:       "0A1628",
  DARK_BLUE:  "12294A",
  MID_BLUE:   "1B3A6B",
  ACCENT_BLUE:"2D7DD2",
  GOLD:       "D4A017",
  WARM_GOLD:  "B8860B",
  WHITE:      "FFFFFF",
  OFF_WHITE:  "F0F2F5",
  ICE_BLUE:   "E8EDF5",
  DARK_TEXT:  "1A1A2E",
  GRAY_TEXT:  "6B7288",
  MUTED:      "9CA3AF",
  TEAL:       "0D9488",
  GREEN_BG:   "E6F7E6",
  RED_ACCENT: "DC2626",
};

// ═══════════════════════════════════════════════
// HELPERS — Factory functions (no reuse!)
// ═══════════════════════════════════════════════
const makeOpts = () => ({});
const slideTitle = (t) => ({ text: t, options: { fontSize: 36, bold: true, color: C.WHITE, fontFace: "Calibri", breakLine: true } });
const slideSubtitle = (t) => ({ text: t, options: { fontSize: 16, color: C.GOLD, fontFace: "Calibri", breakLine: true } });

// ═══════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
pres.author = "Pemerintah Kota Bekasi";
pres.title = "Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024 — Pengelolaan Pajak Reklame";
pres.subject = "Peraturan Daerah";

let slideNum = 0;

const addPageNum = (slide) => {
  slideNum++;
  slide.addText(String(slideNum), {
    x: 12.2, y: 7.0, w: 0.8, h: 0.35,
    fontSize: 9, color: C.GRAY_TEXT, align: "right", fontFace: "Calibri",
  });
};

// ─── SLIDE 1: COVER ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.NAVY };
  
  // Decorative circles (large, subtle)
  s.addShape(pres.shapes.OVAL, { x: 8.5, y: -2, w: 7, h: 7, fill: { color: C.DARK_BLUE, transparency: 60 } });
  s.addShape(pres.shapes.OVAL, { x: 10, y: 4, w: 5, h: 5, fill: { color: C.DARK_BLUE, transparency: 60 } });
  
  // Gold bars — not accent lines under titles, but decorative shapes
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.4, w: 3.5, h: 0.06, fill: { color: C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.8, w: 2.5, h: 0.04, fill: { color: C.WARM_GOLD } });
  
  // Top metadata
  s.addText("BERITA DAERAH KOTA BEKASI", { x: 0.8, y: 0.4, w: 8, h: 0.4, fontSize: 13, color: C.GOLD, bold: true, fontFace: "Calibri" });
  s.addText("NOMOR 51  ·  TAHUN 2024", { x: 0.8, y: 0.8, w: 5, h: 0.3, fontSize: 11, color: C.MUTED, fontFace: "Calibri" });
  
  // Main title
  s.addText("PERATURAN WALI KOTA BEKASI", { x: 0.8, y: 1.5, w: 11, h: 0.6, fontSize: 28, bold: true, color: C.WHITE, fontFace: "Calibri" });
  s.addText("NOMOR 51 TAHUN 2024", { x: 0.8, y: 2.0, w: 11, h: 0.5, fontSize: 20, color: C.GOLD, bold: true, fontFace: "Calibri" });
  
  // Core subject
  s.addText("TENTANG\nPENGELOLAAN PAJAK REKLAME", { x: 0.8, y: 3.0, w: 11, h: 1.6, fontSize: 40, bold: true, color: C.WHITE, fontFace: "Calibri" });
  
  // Footer info
  s.addText("Pemerintah Kota Bekasi", { x: 0.8, y: 4.9, w: 8, h: 0.35, fontSize: 14, color: C.ICE_BLUE, fontFace: "Calibri" });
  s.addText("Ditetapkan: 20 Desember 2024  ·  Berlaku sejak diundangkan", { x: 0.8, y: 5.3, w: 8, h: 0.35, fontSize: 12, color: C.MUTED, fontFace: "Calibri" });
  
  // Bottom gold bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.25, w: 13.333, h: 0.25, fill: { color: C.GOLD } });
})();

// ─── HELPER: Section Divider ───
const addSection = (title, subtitle = "") => {
  const s = pres.addSlide();
  s.background = { color: C.NAVY };
  s.addShape(pres.shapes.OVAL, { x: 9.5, y: -1.5, w: 5, h: 5, fill: { color: C.DARK_BLUE, transparency: 60 } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.5, w: 4, h: 4, fill: { color: C.DARK_BLUE, transparency: 60 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.0, w: 2.0, h: 0.06, fill: { color: C.GOLD } });
  s.addText(title, { x: 0.8, y: 2.3, w: 11, h: 2.0, fontSize: 38, bold: true, color: C.WHITE, fontFace: "Calibri" });
  if (subtitle) s.addText(subtitle, { x: 0.8, y: 4.3, w: 10, h: 0.5, fontSize: 14, color: C.MUTED, fontFace: "Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.25, w: 13.333, h: 0.25, fill: { color: C.GOLD } });
  addPageNum(s);
  return s;
};

// ─── HELPER: Content Slide ───
const addContent = (title, items, opts = {}) => {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };
  
  // Top thin bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 0.04, fill: { color: C.GOLD } });
  
  // Title area with subtle bg
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.04, w: 13.333, h: 0.9, fill: { color: C.OFF_WHITE } });
  
  // A small visual element — colored square instead of accent line
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.2, w: 0.08, h: 0.55, fill: { color: C.ACCENT_BLUE } });
  
  s.addText(title, { x: 0.75, y: 0.2, w: 10, h: 0.6, fontSize: 24, bold: true, color: C.DARK_TEXT, fontFace: "Calibri" });
  
  if (opts.subtitle) {
    s.addText(opts.subtitle, { x: 0.75, y: 0.6, w: 10, h: 0.3, fontSize: 11, color: C.GRAY_TEXT, fontFace: "Calibri" });
  }
  
  // Content area
  const contentY = opts.subtitle ? 1.1 : 1.0;
  
  // Build bullet items
  const textArr = items.map((item, i) => {
    if (item === "") return { text: "", options: { breakLine: true, fontSize: 8, color: C.WHITE } };
    const isHeader = item.startsWith("**");
    const isSub = item.startsWith("   ");
    const clean = item.replace(/\*\*/g, "").trim();
    const last = i === items.length - 1;
    return {
      text: isHeader ? clean : (isSub ? clean : clean),
      options: {
        bullet: !isSub && clean.length > 0 ? true : undefined,
        indentLevel: isSub ? 1 : 0,
        bold: isHeader ? true : false,
        breakLine: !last,
        fontSize: isHeader ? 15 : 14,
        color: isHeader ? C.DARK_TEXT : C.DARK_TEXT,
        fontFace: "Calibri",
        paraSpaceAfter: 2,
      }
    };
  }).filter(Boolean);
  
  s.addText(textArr, { x: 0.6, y: contentY, w: opts.twoCol ? 5.5 : 11.5, h: 5.5, valign: "top", margin: 0 });
  
  // If two-col, add second column for second half of items
  if (opts.twoCol && items.length > 0) {
    // Let's just put visual placeholder - a colored shape
    s.addShape(pres.shapes.RECTANGLE, { 
      x: 7.0, y: contentY, w: 5.5, h: 5.0,
      fill: { color: C.ICE_BLUE, transparency: 40 },
    });
  }
  
  if (opts.notes) {
    s.addText(opts.notes, { x: 0.6, y: 6.6, w: 11, h: 0.4, fontSize: 10, italic: true, color: C.GRAY_TEXT, fontFace: "Calibri" });
  }
  
  // Bottom
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.3, w: 13.333, h: 0.2, fill: { color: C.NAVY } });
  addPageNum(s);
  return s;
};

// ─── HELPER: Card Slide ───
const addCards = (title, cards, opts = {}) => {
  const s = pres.addSlide();
  s.background = { color: C.OFF_WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 0.04, fill: { color: C.GOLD } });
  
  // Title area
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.04, w: 13.333, h: 0.8, fill: { color: C.WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.15, w: 0.08, h: 0.5, fill: { color: C.ACCENT_BLUE } });
  s.addText(title, { x: 0.75, y: 0.15, w: 11, h: 0.5, fontSize: 22, bold: true, color: C.DARK_TEXT, fontFace: "Calibri" });
  if (opts.subtitle) s.addText(opts.subtitle, { x: 0.75, y: 0.5, w: 10, h: 0.3, fontSize: 11, color: C.GRAY_TEXT, fontFace: "Calibri" });
  
  const numCards = cards.length;
  const cardW = numCards === 2 ? 5.8 : (numCards === 3 ? 3.8 : (numCards === 4 ? 2.85 : 5.8));
  const gap = 0.3;
  const totalW = numCards * cardW + (numCards - 1) * gap;
  const startX = (13.333 - totalW) / 2;
  
  const cardColors = [C.ACCENT_BLUE, C.TEAL, C.WARM_GOLD, C.MID_BLUE];
  
  cards.forEach((card, i) => {
    const cx = startX + i * (cardW + gap);
    const cy = 1.1;
    
    // Card background
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cy, w: cardW, h: 5.5,
      fill: { color: C.WHITE },
      shadow: { type: "outer", blur: 4, offset: 1, angle: 135, color: "000000", opacity: 0.08 },
    });
    
    // Card top accent strip
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: 0.06,
      fill: { color: cardColors[i % cardColors.length] },
    });
    
    // Card title
    s.addText(card.title, {
      x: cx + 0.15, y: cy + 0.15, w: cardW - 0.3, h: 0.4,
      fontSize: 15, bold: true, color: cardColors[i % cardColors.length], fontFace: "Calibri",
    });
    
    // Card items
    const itemArr = card.items.map((item, j) => ({
      text: item,
      options: {
        bullet: true,
        breakLine: j < card.items.length - 1,
        fontSize: 12,
        color: C.DARK_TEXT,
        fontFace: "Calibri",
        paraSpaceAfter: 2,
      }
    }));
    
    s.addText(itemArr, {
      x: cx + 0.15, y: cy + 0.6, w: cardW - 0.3, h: 4.7,
      valign: "top", margin: 0,
    });
  });
  
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.3, w: 13.333, h: 0.2, fill: { color: C.NAVY } });
  addPageNum(s);
  return s;
};

// ─── HELPER: Table Slide ───
const addTableSlide = (title, headers, rows, opts = {}) => {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 0.04, fill: { color: C.GOLD } });
  
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.04, w: 13.333, h: 0.8, fill: { color: C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.15, w: 0.08, h: 0.5, fill: { color: C.ACCENT_BLUE } });
  s.addText(title, { x: 0.75, y: 0.15, w: 11, h: 0.5, fontSize: 22, bold: true, color: C.DARK_TEXT, fontFace: "Calibri" });
  if (opts.subtitle) s.addText(opts.subtitle, { x: 0.75, y: 0.5, w: 10, h: 0.3, fontSize: 11, color: C.GRAY_TEXT, fontFace: "Calibri" });
  
  const headerRow = headers.map(h => ({
    text: h,
    options: { bold: true, color: C.WHITE, fill: { color: C.NAVY }, fontSize: 12, fontFace: "Calibri", align: "center", valign: "middle" }
  }));
  
  const dataRows = rows.map((row, i) => {
    const bgColor = i % 2 === 0 ? C.ICE_BLUE : C.WHITE;
    return row.map((cell, j) => ({
      text: String(cell),
      options: { 
        fontSize: 11, color: C.DARK_TEXT, fontFace: "Calibri",
        fill: { color: bgColor },
        align: j === 0 ? "left" : "center",
        valign: "middle",
      }
    }));
  });
  
  const tableData = [headerRow, ...dataRows];
  
  s.addTable(tableData, {
    x: 0.8, y: 1.1, w: 11.7,
    colW: opts.colW || undefined,
    border: { pt: 0.5, color: C.MUTED },
    rowH: 0.45,
    margin: [2, 6, 2, 6],
  });
  
  if (opts.notes) {
    s.addText(opts.notes, { x: 0.8, y: 6.7, w: 11, h: 0.35, fontSize: 10, italic: true, color: C.GRAY_TEXT, fontFace: "Calibri" });
  }
  
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.3, w: 13.333, h: 0.2, fill: { color: C.NAVY } });
  addPageNum(s);
  return s;
};

// ─── HELPER: Two-Column Content ───
const addTwoCol = (title, leftItems, rightItems, opts = {}) => {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 0.04, fill: { color: C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.04, w: 13.333, h: 0.8, fill: { color: C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.15, w: 0.08, h: 0.5, fill: { color: C.ACCENT_BLUE } });
  s.addText(title, { x: 0.75, y: 0.15, w: 11, h: 0.5, fontSize: 22, bold: true, color: C.DARK_TEXT, fontFace: "Calibri" });
  if (opts.subtitle) s.addText(opts.subtitle, { x: 0.75, y: 0.5, w: 10, h: 0.3, fontSize: 11, color: C.GRAY_TEXT, fontFace: "Calibri" });

  const cy = opts.subtitle ? 1.1 : 1.0;
  const w = 5.5;
  
  const makeBullets = (items) => items.map((item, i) => {
    if (item === "") return { text: "", options: { breakLine: true, fontSize: 6 } };
    const isH = item.startsWith("**");
    const isS = item.startsWith("   ");
    const clean = item.replace(/\*\*/g, "").trim();
    const last = i === items.length - 1;
    return {
      text: clean,
      options: {
        bullet: !isS && clean.length > 0 ? true : undefined,
        indentLevel: isS ? 1 : 0,
        bold: isH,
        breakLine: !last,
        fontSize: 14,
        color: C.DARK_TEXT,
        fontFace: "Calibri",
        paraSpaceAfter: 2,
      }
    };
  });
  
  // Left column with colored bg
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: cy, w: w, h: 5.5,
    fill: { color: C.ICE_BLUE, transparency: 50 },
  });
  s.addText(makeBullets(leftItems), { x: 0.65, y: cy + 0.1, w: w - 0.3, h: 5.3, valign: "top", margin: 0 });
  
  // Right column with different bg
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.5, y: cy, w: w, h: 5.5,
    fill: { color: C.GREEN_BG, transparency: 50 },
  });
  s.addText(makeBullets(rightItems), { x: 6.65, y: cy + 0.1, w: w - 0.3, h: 5.3, valign: "top", margin: 0 });
  
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.3, w: 13.333, h: 0.2, fill: { color: C.NAVY } });
  addPageNum(s);
  return s;
};

// ═══════════════════════════════════════════════
// SLIDES
// ═══════════════════════════════════════════════

// ─── DAFTAR ISI ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.WHITE };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 0.04, fill: { color: C.GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.04, w: 13.333, h: 0.8, fill: { color: C.OFF_WHITE } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.15, w: 0.08, h: 0.5, fill: { color: C.ACCENT_BLUE } });
  s.addText("DAFTAR ISI", { x: 0.75, y: 0.15, w: 11, h: 0.5, fontSize: 24, bold: true, color: C.DARK_TEXT, fontFace: "Calibri" });
  s.addText("Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024", { x: 0.75, y: 0.5, w: 10, h: 0.3, fontSize: 11, color: C.GRAY_TEXT, fontFace: "Calibri" });

  const toc = [
    ["BAB I", "Ketentuan Umum"],
    ["BAB II", "Objek Pajak, Subjek Pajak, Wajib Pajak"],
    ["BAB III", "Masa Pajak, Tahun Pajak"],
    ["BAB IV", "Pendaftaran & Pendataan WP"],
    ["BAB V", "Nilai Sewa Reklame (NSR)"],
    ["BAB VI", "Perhitungan Pajak Reklame"],
    ["BAB VII", "Penetapan Pajak Terutang"],
    ["BAB VIII", "Surat Tagihan Pajak"],
    ["BAB IX", "Pembayaran & Penyetoran"],
    ["BAB X", "Pembetulan & Pembatalan Ketetapan"],
  ];
  const toc2 = [
    ["BAB XI", "Pengembalian Kelebihan Bayar"],
    ["BAB XII", "Pemeriksaan Pajak"],
    ["BAB XIII", "Penagihan Pajak"],
    ["BAB XIV", "Kedaluwarsa Penagihan"],
    ["BAB XV", "Penghapusan Piutang Pajak"],
    ["BAB XVI", "Keringanan, Pengurangan, Pembebasan"],
    ["BAB XVII", "Kemudahan Perpajakan"],
    ["BAB XVIII", "Keberatan & Banding"],
    ["BAB XIX", "Penghargaan"],
    ["BAB XX", "Ketentuan Penutup"],
  ];

  const makeTocCol = (data, x) => {
    data.forEach(([bab, title], i) => {
      const y = 1.1 + i * 0.52;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y, w: 1.1, h: 0.38,
        fill: { color: i < 10 ? C.ACCENT_BLUE : C.TEAL },
      });
      s.addText(bab, { x, y: y + 0.02, w: 1.1, h: 0.35, fontSize: 10, bold: true, color: C.WHITE, align: "center", fontFace: "Calibri" });
      s.addText(title, { x: x + 1.25, y: y + 0.02, w: 4.8, h: 0.35, fontSize: 13, color: C.DARK_TEXT, fontFace: "Calibri" });
    });
  };
  makeTocCol(toc, 0.6);
  makeTocCol(toc2, 6.8);

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.3, w: 13.333, h: 0.2, fill: { color: C.NAVY } });
  addPageNum(s);
})();

// ═══════════════════════════════════════════════
// BAB I: KETENTUAN UMUM
// ═══════════════════════════════════════════════
addSection("BAB I: KETENTUAN UMUM", "Pasal 1 — 45 Definisi dan Istilah");

addContent("Definisi Kunci", [
  "**Daerah** — Daerah Kota Bekasi",
  "**Pemerintah Daerah** — Wali Kota sebagai unsur penyelenggara pemerintahan",
  "**Bapenda** — Badan Pendapatan Daerah Kota Bekasi",
  "**Reklame** — Benda/alat/media untuk tujuan komersial memperkenalkan, mempromosikan sesuatu",
  "**Pajak Reklame** — Pajak atas penyelenggaraan reklame",
  "**Nilai Sewa Reklame (NSR)** — Dasar pengenaan Pajak Reklame (nilai jual + nilai strategis)",
  "**NPWPD** — Nomor Pokok Wajib Pajak Daerah",
  "**NOPD** — Nomor Objek Pajak Daerah",
  "**SKPD** — Surat Ketetapan Pajak Daerah",
  "**Wajib Pajak** — Orang pribadi/badan dengan hak dan kewajiban perpajakan",
  "",
  "**Penanggung Pajak** — Pihak bertanggung jawab atas pembayaran pajak",
], { notes: "Pasal 1 — hanya definisi kunci yang ditampilkan" });

// ═══════════════════════════════════════════════
// BAB II: OBJEK PAJAK, SUBJEK PAJAK, WAJIB PAJAK
// ═══════════════════════════════════════════════
addSection("BAB II: OBJEK PAJAK, SUBJEK PAJAK, DAN WAJIB PAJAK", "Pasal 2–4");

addTwoCol("Jenis & Pengecualian Objek Pajak",
  // Left: Jenis
  ["**Objek Pajak Reklame (Pasal 2 ayat 1-2):**", "", "Semua penyelenggaraan reklame:", "Reklame Papan/Billboard", "Reklame Videotron/Megatron", "Reklame Kain (spanduk, umbul, baliho)", "Reklame Melekat/Stiker", "Reklame Selebaran", "Reklame Berjalan (kendaraan)", "Reklame Udara (balon gas)", "Reklame Apung", "Reklame Film/Slide", "Reklame Peragaan"],
  // Right: Pengecualian
  ["**Dikecualikan dari Objek (Pasal 2 ayat 3):**", "", "Internet, TV, radio, media cetak", "Label/merek produk pada barang", "Nama usaha/profesi ≤ 1 m² di tempat", "Reklame Pemerintah/Pemerintah Daerah", "Nama tempat ibadah & panti asuhan", "Reklame tanah ≤ 1 m²", "Kegiatan politik (masa kampanye KPU)", "Sosial/keagamaan (≤ 30 hari)", "Olahraga KONI (≤ 30 hari)"]
);

addContent("Subjek & Wajib Pajak (Pasal 3–4)", [
  "**Subjek Pajak Reklame** (Pasal 3): Orang pribadi atau Badan yang menggunakan Reklame",
  "",
  "**Wajib Pajak Reklame** (Pasal 4):",
  "Orang pribadi atau Badan yang menyelenggarakan Reklame",
  "Jika reklame diselenggarakan pihak ketiga → pihak ketiga menjadi Wajib Pajak",
]);

// ═══════════════════════════════════════════════
// BAB III: MASA PAJAK
// ═══════════════════════════════════════════════
addSection("BAB III: MASA PAJAK, TAHUN PAJAK", "Pasal 5");

addCards("Masa Pajak & Tahun Pajak", [
  { title: "Masa Pajak", items: ["Permanen: 12 bulan", "Sesuai masa penayangan", "Insidentil: per hari", "Maks. 30 hari"] },
  { title: "Tahun Pajak", items: ["1 tahun kalender", "Atau sesuai tahun buku WP"] },
  { title: "Bagian Tahun Pajak", items: ["1 tahun pajak atas", "1 bulan kalender", "Beberapa bulan kalender"] },
]);

// ═══════════════════════════════════════════════
// BAB IV: PENDAFTARAN & PENDATAAN
// ═══════════════════════════════════════════════
addSection("BAB IV: PENDAFTARAN & PENDATAAN WAJIB PAJAK", "Pasal 6–8");

addCards("Pendaftaran, Pendataan & Penonaktifan", [
  { title: "Pendaftaran WP (Pasal 6)", items: [
    "WP wajib daftarkan diri & objek",
    "Form: ambil di Bapenda/online/dikirim",
    "Lampirkan: KTP, NPWP, Akta, NIB",
    "Surat kuasa (jika dikuasakan)",
    "Bapenda terbitkan NPWPD",
    "Jika tidak daftar → NPWPD jabatan",
  ]},
  { title: "Pendataan (Pasal 7)", items: [
    "Bapenda lakukan pendataan",
    "WP & objek Pajak Reklame",
    "Termasuk data geografis",
    "Bisa kerjasama dg instansi lain",
  ]},
  { title: "Penonaktifan (Pasal 8)", items: [
    "Jika WP tak penuhi syarat",
    "Jabatan atau permohonan WP",
    "Keputusan maks. 3 bulan",
    "Syarat: tanpa tunggakan",
    "Tidak sedang ajukan keberatan",
    "WP nonaktif wajib lapor",
  ]},
]);

// ═══════════════════════════════════════════════
// BAB V: NILAI SEWA REKLAME
// ═══════════════════════════════════════════════
addSection("BAB V: NILAI SEWA REKLAME (NSR)", "Pasal 9 — Dasar Pengenaan Pajak");

addContent("Faktor Penentu NSR", [
  "Dasar pengenaan Pajak Reklame adalah **Nilai Sewa Reklame (NSR)**",
  "NSR = Nilai Jual Objek Pajak Reklame + Nilai Strategis Pemasangan",
  "",
  "**7 Faktor Perhitungan NSR:**",
  "Jenis reklame",
  "Bahan yang digunakan",
  "Lokasi penempatan (Kelas Jalan)",
  "Waktu (satuan detik)",
  "Jangka waktu (hari kalender)",
  "Jumlah media reklame (lembar)",
  "Ukuran media reklame",
], { notes: "Pasal 9 ayat (1)–(2)" });

addContent("Klasifikasi Kelas Jalan", [
  "**Kelas Jalan Khusus:**",
  "Zona Tol | Zona Premium I | Zona Premium II",
  "Premium I: Jl. A. Yani, Cut Mutia, Juanda, Sudirman, Sultan Agung, Transyogi, KH. Noer Ali",
  "Premium II: Jl. Narogong Siliwangi, Jatiwaringin, Pekayon Jatiasih, Jatiasih Pondokgede, Jatimakmur, Joyo Martono, Chairil Anwar, Bintara",
  "",
  "**Kelas Jalan I (Kendali Ketat):**",
  "Lebar > 3 m, pusat pelayanan/permukiman",
  "",
  "**Kelas Jalan II (Kendali Sedang):**",
  "Lebar ≤ 3 m atau jalan lingkungan perumahan",
], { notes: "Pasal 9 ayat (3)–(8)" });

// ═══════════════════════════════════════════════
// BAB VI: PERHITUNGAN PAJAK REKLAME
// ═══════════════════════════════════════════════
addSection("BAB VI: PERHITUNGAN PAJAK REKLAME", "Pasal 10 — Tarif & Rumus");

addContent("Rumus Dasar Perhitungan", [
  "**Rumus Dasar:**",
  "Pajak Reklame = Tarif Pajak Reklame × NSR",
  "",
  "**Rumus NSR Papan/Billboard & Megatron/Videotron:**",
  "Nilai Kelas Jalan × Ukuran (m²) × Jumlah × Jangka Waktu",
  "",
  "**Ketentuan Khusus:**",
  "Reklame indoor: NSR = 50% dari hasil perhitungan",
  "Ketinggian > 15 m: tambahan 20% dari NSR",
  "Produk tembakau & minuman keras: tambahan 50% dari NSR",
  "Perubahan naskah dalam 1 badan usaha: dikecualikan",
  "Perubahan bentuk/ukuran: pajak atas selisih",
], { notes: "Pasal 10 ayat (1)–(13)" });

addTableSlide("NSR — Reklame Papan/Billboard",
  ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"],
  [
    ["Kelas Jalan Khusus", "Jalan Tol", "Rp 23.575"],
    ["Kelas Jalan Khusus", "Premium 1", "Rp 16.100"],
    ["Kelas Jalan Khusus", "Premium 2", "Rp 14.950"],
    ["Kelas Jalan I", "Kendali Ketat", "Rp 13.225"],
    ["Kelas Jalan II", "Kendali Sedang", "Rp 11.500"],
  ],
  { notes: "Pasal 10 ayat (5) huruf a · Satuan: 1 m², 1 buah, 1 hari", colW: [3, 6, 2.7] }
);

addTableSlide("NSR — Megatron/Videotron",
  ["Kelas Jalan", "Zona", "NSR (/30 detik)", "NSR (/m²/tahun)"],
  [
    ["Kelas Jalan Khusus", "Jalan Tol", "Rp 17,25", "Rp 13.599.900"],
    ["Kelas Jalan Khusus", "Premium 1", "Rp 13,80", "Rp 10.879.920"],
    ["Kelas Jalan Khusus", "Premium 2", "Rp 9,20", "Rp 7.253.280"],
    ["Kelas Jalan I", "Kendali Ketat", "Rp 8,05", "Rp 6.346.620"],
    ["Kelas Jalan II", "Kendali Sedang", "Rp 5,75", "Rp 4.533.300"],
  ],
  { notes: "Pasal 10 ayat (5) huruf b · 18 jam/hari = 2.160 tayangan/hari", colW: [3, 4, 2.7, 2.7] }
);

addTableSlide("NSR — Reklame Kain (Spanduk, Umbul, Baliho)",
  ["Kelas Jalan", "Zona", "NSR (Rp/m²/hari)"],
  [
    ["Kelas Jalan Khusus", "Jalan Tol", "Rp 30.000"],
    ["Kelas Jalan Khusus", "Premium 1", "Rp 30.000"],
    ["Kelas Jalan Khusus", "Premium 2", "Rp 25.000"],
    ["Kelas Jalan I", "Kendali Ketat", "Rp 20.000"],
    ["Kelas Jalan II", "Kendali Sedang", "Rp 19.000"],
  ],
  { notes: "Pasal 10 ayat (5) huruf c · Satuan: 1 m², 1 buah, 1 hari", colW: [3, 6, 2.7] }
);

addCards("NSR — Jenis Reklame Lainnya", [
  { title: "Stiker", items: ["Rp 7,5/cm²", "Min. Rp 750.000/kali"] },
  { title: "Melekat", items: ["Rp 750.000/m²/tahun"] },
  { title: "Selebaran", items: ["Rp 600/lembar", "Min. Rp 6.000.000/kali"] },
  { title: "Berjalan", items: ["Rp 6.000/m²/hari", "Termasuk kendaraan"] },
]);

addCards("NSR — Jenis Reklame Lainnya (lanjutan)", [
  { title: "Udara", items: ["Rp 2.400.000/sekali", "Maks. 1 bulan"] },
  { title: "Apung", items: ["Rp 600.000/sekali", "Maks. 1 bulan"] },
  { title: "Film/Slide", items: ["Rp 12.000/15 detik", "Kurang 15\" dibulatkan"] },
  { title: "Peragaan", items: ["Rp 480.000/penyelenggaraan"] },
]);

// ═══════════════════════════════════════════════
// BAB VII: PENETAPAN PAJAK TERUTANG
// ═══════════════════════════════════════════════
addSection("BAB VII: PENETAPAN PAJAK TERUTANG", "Pasal 11");

addCards("Skema Penetapan Pajak", [
  { title: "WP Mendaftar", items: ["SKPD berdasarkan", "data pendaftaran WP"] },
  { title: "WP Tidak Mendaftar", items: ["SKPD diterbitkan", "secara JABATAN"] },
  { title: "Hasil Pemeriksaan", items: ["Jika pajak > laporan", "SKPD sesuai temuan", "Tanpa sanksi admin"] },
  { title: "Batas Waktu", items: ["Maks. 5 tahun", "sejak terutangnya pajak"] },
]);

// ═══════════════════════════════════════════════
// BAB VIII: SURAT TAGIHAN PAJAK
// ═══════════════════════════════════════════════
addSection("BAB VIII: SURAT TAGIHAN PAJAK (STPD)", "Pasal 12");

addTwoCol("STPD — Surat Tagihan Pajak Daerah",
  ["**Dasar Penerbitan STPD:**", "", "Jangka waktu ≤ 5 tahun sejak terutang", "", "a. Pajak SKPD tidak/kurang dibayar", "   setelah jatuh tempo", "", "b. SK Pembetulan/Keberatan/Banding", "   tidak/kurang dibayar", "", "c. WP dikenakan sanksi admin", "   berupa bunga/denda"],
  ["**Sanksi Bunga:**", "", "Huruf a: Bunga 1% per bulan", "   dari pajak kurang bayar", "   Maks. 24 bulan", "", "Huruf b: Bunga 0,6% per bulan", "   dari pajak kurang bayar", "   Maks. 24 bulan", "", "Bagian bulan dihitung 1 bulan penuh"]
);

// ═══════════════════════════════════════════════
// BAB IX: PEMBAYARAN & PENYETORAN
// ═══════════════════════════════════════════════
addSection("BAB IX: PEMBAYARAN DAN PENYETORAN", "Pasal 13–14");

addContent("Tata Cara Pembayaran", [
  "**Metode Pembayaran:**",
  "Lunas melalui Kas Daerah / tempat ditunjuk Wali Kota",
  "Prioritas: sistem pembayaran elektronik",
  "Jika tidak tersedia: pembayaran tunai",
  "",
  "**Jatuh Tempo:** Paling lama 1 bulan sejak pengiriman SKPD",
  "",
  "**Keterlambatan:** Sanksi bunga 1%/bulan (maks. 24 bulan) + STPD",
  "STPD harus dilunasi ≤ 30 hari sejak pengiriman",
  "",
  "**Tanda Bukti:** Stiker reklame + bukti dari Bank Persepsi",
], { notes: "Pasal 13–14" });

// ═══════════════════════════════════════════════
// BAB X: PEMBETULAN & PEMBATALAN
// ═══════════════════════════════════════════════
addSection("BAB X: PEMBETULAN & PEMBATALAN KETETAPAN", "Pasal 15–20");

addTwoCol("Dasar & Prosedur Pembetulan",
  ["**Dasar Pembetulan (Pasal 15):**", "", "a. Kesalahan tulis", "   Nama, alamat, NPWPD, nomor surat,", "   masa pajak, SKPD ganda", "", "b. Kesalahan hitung", "   Penjumlahan, pengurangan,", "   perkalian, pembagian, tarif", "", "c. Kekeliruan penerapan", "   ketentuan perundang-undangan"],
  ["**Prosedur (Pasal 16–20):**", "", "1 permohonan = 1 ketetapan", "Diajukan tertulis (B. Indonesia)", "Lampirkan identitas + dokumen asli", "", "**Keputusan:** maks. 6 bulan", "Jika > 6 bulan → dikabulkan", "", "Keputusan: kabul/tambah/kurang/", "hapus/batal/tolak", "", "Pembetulan jabatan (Pasal 19)", "Bisa diulang jika masih salah (Pasal 20)"]
);

// ═══════════════════════════════════════════════
// BAB XI: PENGEMBALIAN KELEBIHAN BAYAR
// ═══════════════════════════════════════════════
addSection("BAB XI: PENGEMBALIAN KELEBIHAN PEMBAYARAN", "Pasal 21");

addContent("Prosedur Pengembalian", [
  "WP berhak mengajukan permohonan pengembalian kelebihan pembayaran pajak",
  "",
  "**Alur Pengembalian:**",
  "WP mengajukan → Keputusan maks. 12 bulan",
  "Jika > 12 bulan tanpa keputusan → dianggap dikabulkan",
  "SKPDLB diterbitkan dalam 1 bulan",
  "",
  "Jika WP punya utang pajak lain → kelebihan diperhitungkan",
  "",
  "**Pengembalian dana:** maks. 2 bulan sejak SKPDLB",
  "Jika terlambat → imbalan bunga 0,6%/bulan",
], { notes: "Pasal 21" });

// ═══════════════════════════════════════════════
// BAB XII: PEMERIKSAAN PAJAK
// ═══════════════════════════════════════════════
addSection("BAB XII: PEMERIKSAAN PAJAK", "Pasal 22–23");

addCards("Pemeriksaan Pajak", [
  { title: "Kewenangan", items: ["Kepala Bapenda", "berwenang periksa", "Uji kepatuhan WP", "Tujuan lain: NPWPD,", "keberatan, penagihan"] },
  { title: "Kewajiban WP", items: ["Tunjukkan buku/", "catatan/dokumen", "Beri akses tempat", "Beri keterangan", "Jika tidak → pajak", "ditetapkan jabatan"] },
  { title: "Hak WP", items: ["Minta identitas", "pemeriksa", "Minta penjelasan", "tujuan periksa", "Terima hasil periksa", "Beri tanggapan"] },
]);

// ═══════════════════════════════════════════════
// BAB XIII: PENAGIHAN PAJAK
// ═══════════════════════════════════════════════
addSection("BAB XIII: PENAGIHAN PAJAK", "Pasal 24");

addContent("Penagihan Pajak", [
  "**Dasar Penagihan Pajak:**",
  "SKPD | SKPDKB | SKPDKBT | STPD | SK Pembetulan | SK Keberatan | Putusan Banding",
  "",
  "**Prosedur:**",
  "Sebelum jatuh tempo → imbauan",
  "Setelah jatuh tempo & belum lunas → Penagihan Pajak sesuai peraturan",
], { notes: "Pasal 24" });

// ═══════════════════════════════════════════════
// BAB XIV: KEDALUWARSA PENAGIHAN
// ═══════════════════════════════════════════════
addSection("BAB XIV: KEDALUWARSA PENAGIHAN PAJAK", "Pasal 25");

addContent("Kedaluwarsa Penagihan", [
  "**Jangka Waktu:** 5 tahun sejak terutangnya pajak",
  "Kecuali WP melakukan tindak pidana perpajakan",
  "",
  "**Kedaluwarsa Tertangguh jika:**",
  "a. Diterbitkan Surat Teguran dan/atau Surat Paksa",
  "   — Dihitung sejak tanggal penyampaian",
  "b. Ada pengakuan utang pajak dari WP",
  "   — Langsung: WP menyatakan masih punya utang",
  "   — Tidak langsung: ajukan angsuran/penundaan/keberatan",
  "   — Dihitung sejak tanggal pengakuan",
], { notes: "Pasal 25" });

// ═══════════════════════════════════════════════
// BAB XV: PENGHAPUSAN PIUTANG
// ═══════════════════════════════════════════════
addSection("BAB XV: PENGHAPUSAN PIUTANG PAJAK", "Pasal 26");

addContent("Prosedur Penghapusan Piutang", [
  "Piutang yang tidak mungkin ditagih lagi karena kedaluwarsa → dapat dihapuskan",
  "",
  "**Prosedur:**",
  "Bapenda melakukan penelitian → Berita Acara",
  "Tim peneliti ditetapkan Keputusan Wali Kota",
  "Kepala Bapenda susun daftar usulan penghapusan",
  "Disampaikan ke Wali Kota",
  "Ditetapkan dengan Keputusan Wali Kota",
  "",
  "**Pertimbangan:** Penagihan sampai batas kedaluwarsa + koordinasi Inspektorat Daerah",
], { notes: "Pasal 26" });

// ═══════════════════════════════════════════════
// BAB XVI: KERINGANAN, PENGURANGAN, PEMBEBASAN
// ═══════════════════════════════════════════════
addSection("BAB XVI: KERINGANAN, PENGURANGAN, & PEMBEBASAN", "Pasal 27");

addCards("Fasilitas Pajak", [
  { title: "Bentuk Fasilitas", items: ["Keringanan", "Pengurangan", "Pembebasan", "Penundaan bayar", "Atas pokok &/ sanksi"] },
  { title: "Kondisi WP", items: ["Kemampuan membayar", "Tingkat likuiditas"] },
  { title: "Kondisi Objek", items: ["Bencana alam", "Kebakaran", "Huru-hara/kerusuhan"] },
]);

// ═══════════════════════════════════════════════
// BAB XVII: KEMUDAHAN PERPAJAKAN
// ═══════════════════════════════════════════════
addSection("BAB XVII: KEMUDAHAN PERPAJAKAN DAERAH", "Pasal 28");

addContent("Kemudahan Perpajakan", [
  "**Bentuk Kemudahan:**",
  "a. Perpanjangan batas waktu bayar — untuk WP keadaan kahar",
  "b. Fasilitas angsuran/penundaan — untuk WP kesulitan likuiditas",
  "",
  "**Keadaan Kahar:** Bencana alam | Kebakaran | Kerusuhan | Wabah | Keadaan lain (keputusan Wali Kota)",
  "",
  "**Ketentuan:**",
  "Jangka waktu: maks. 24 bulan",
  "Bunga: 0,6%/bulan dari jumlah pajak",
  "Wali Kota lihat kepatuhan WP 2 tahun terakhir",
], { notes: "Pasal 28" });

// ═══════════════════════════════════════════════
// BAB XVIII: KEBERATAN & BANDING
// ═══════════════════════════════════════════════
addSection("BAB XVIII: KEBERATAN DAN BANDING", "Pasal 29–33");

addTwoCol("Keberatan (Pasal 29–31)",
  ["**Objek Keberatan:**", "SKPD, SKPDKB, SKPDKBT", "SKPDLB, SKPDN", "Pemotongan pihak ke-3", "", "**Syarat:**", "Tertulis B. Indonesia", "Maks. 3 bulan sejak SKPD", "WP bayar min. yg disetujui", "", "**Keputusan:** maks. 12 bulan", "Jika tanpa keputusan → dikabulkan", "", "**Jika ditolak:** denda 30%"],
  ["**Banding (Pasal 32–33):**", "", "**Objek Banding:**", "SK Keberatan → badan peradilan pajak", "Maks. 3 bulan sejak keputusan", "", "**Efek:** Menangguhkan kewajiban bayar", "Sampai 1 bulan sejak Putusan Banding", "", "**Jika ditolak:** denda 60%", "**Jika dikabulkan:**", "kelebihan + bunga 0,6%/bln (maks. 24 bln)", "", "Sanksi 30% tidak dikenakan jika banding"]
);

// ═══════════════════════════════════════════════
// BAB XIX: PENGHARGAAN
// ═══════════════════════════════════════════════
addSection("BAB XIX: PENGHARGAAN", "Pasal 34–35");

addContent("Penghargaan WP Taat Pajak", [
  "**Bentuk:** Piagam | Hadiah | dan/atau sejenisnya (dibebankan ke APBD)",
  "",
  "**Kriteria WP Taat Pajak:**",
  "Bayar tepat waktu minimal 1 tahun",
  "Tanpa tunggakan 3 tahun terakhir",
  "Kontribusi signifikan pada program Pemda",
  "",
  "**Penetapan:** Keputusan Wali Kota + tim penilai",
], { notes: "Pasal 34–35" });

// ═══════════════════════════════════════════════
// BAB XX: KETENTUAN PENUTUP
// ═══════════════════════════════════════════════
addSection("BAB XX: KETENTUAN PENUTUP", "Pasal 36–37");

addTwoCol("Pencabutan & Mulai Berlaku",
  ["**Pasal 36 — Pencabutan:**", "", "Peraturan dicabut:", "❌ Perwal No. 48/2012", "   Petunjuk Pelaksanaan Perda 14/2012", "", "❌ Perwal No. 52/2013", "   Perubahan Perwal 48/2012"],
  ["**Pasal 37 — Mulai Berlaku:**", "", "✅ Berlaku sejak diundangkan", "", "━━━━━━━━━━━━━━━━━━━━", "📅 Ditetapkan: 20 Des 2024", "", "Pj. WALI KOTA BEKASI,", "ttd", "R. GANI MUHAMAD", "", "📅 Diundangkan: 20 Des 2024", "SEKDA KOTA BEKASI,", "ttd", "JUNAEDI"]
);

// ─── CLOSING SLIDE ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.NAVY };
  s.addShape(pres.shapes.OVAL, { x: 9, y: -2, w: 6, h: 6, fill: { color: C.DARK_BLUE, transparency: 60 } });
  s.addShape(pres.shapes.OVAL, { x: -2, y: 4, w: 5, h: 5, fill: { color: C.DARK_BLUE, transparency: 60 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.4, w: 3.5, h: 0.06, fill: { color: C.GOLD } });
  
  s.addText("BERITA DAERAH KOTA BEKASI", { x: 0.8, y: 1.5, w: 10, h: 0.4, fontSize: 14, color: C.GOLD, bold: true, fontFace: "Calibri" });
  s.addText("TERIMA KASIH", { x: 0.8, y: 2.3, w: 11, h: 1.5, fontSize: 48, bold: true, color: C.WHITE, fontFace: "Calibri" });
  s.addText("Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024\nTentang Pengelolaan Pajak Reklame", { x: 0.8, y: 4.0, w: 11, h: 0.8, fontSize: 16, color: C.ICE_BLUE, fontFace: "Calibri" });
  s.addText("Sumber: https://jdih.bekasikota.go.id", { x: 0.8, y: 5.0, w: 11, h: 0.4, fontSize: 12, color: C.MUTED, fontFace: "Calibri" });
  
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.25, w: 13.333, h: 0.25, fill: { color: C.GOLD } });
})();

// ═══════════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════════
const path = require("path");
const outputPath = path.resolve(__dirname, "Perwal_Bekasi_51_2024_Pajak_Reklame.pptx");

pres.writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`✅ PPT V3 saved to: ${outputPath}`);
    console.log(`   Total slides: ${pres.slides.length}`);
    const fs = require("fs");
    const size = fs.statSync(outputPath).size;
    console.log(`   File size: ${(size / 1024).toFixed(1)} KB`);
  })
  .catch(err => console.error("❌ Error:", err));
