#!/usr/bin/env node
/**
 * Perwal Bekasi No 51/2024 — Pajak Reklame
 * Professional Edition v4.0
 *
 * Desain: clean, premium, government-grade
 * Tools: PptxGenJS v4
 * Layout: Widescreen 13.33×7.5 in
 * Font: Calibri (heading) / Calibri (body)
 */

const PptxGenJS = require("pptxgenjs");
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Pemerintah Kota Bekasi";
pres.title = "Perwal Bekasi No 51/2024 - Pengelolaan Pajak Reklame";
pres.subject = "Peraturan Wali Kota Bekasi tentang Pajak Reklame";

// ─── COLOR PALETTE ───
const C = {
  navy:      "0A1628",
  navyLight: "12294A",
  navyMid:   "1B3A6B",
  gold:      "C8962E",
  goldLight: "D4A017",
  white:     "FFFFFF",
  offWhite:  "F5F7FA",
  ice:       "E8EDF5",
  textDark:  "1A1A2E",
  textGray:  "6B7288",
  textMuted: "9CA3AF",
  accent:    "2563EB",
  teal:      "0D9488",
  warm:      "B8860B",
  red:       "DC2626",
  green:     "16A34A",
  purple:    "7C3AED",
  orange:    "EA580C",
};

// ─── LAYOUT CONSTANTS ───
const W = 13.333;
const H = 7.5;
const M = 0.5;  // margin
const CW = W - 2 * M; // content width

// ─── GLOBAL COUNTER ───
let pgNum = 0;
const pageNum = (slide) => {
  pgNum++;
  slide.addText(String(pgNum), {
    x: W - 0.9, y: H - 0.35, w: 0.7, h: 0.25,
    fontSize: 8, color: C.textMuted, align: "right", fontFace: "Calibri",
  });
};

// ─── SHARED DECORATIONS ───
const goldBar = (slide, y = 0, w = W, h = 0.035) =>
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y, w, h, fill: { color: C.gold } });

const headerBar = (slide, title, subtitle) => {
  goldBar(slide, 0);
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.035, w: W, h: 0.85, fill: { color: C.navy },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 0.12, w: 0.07, h: 0.55, fill: { color: C.gold },
  });
  slide.addText(title, {
    x: M + 0.2, y: 0.13, w: CW - 1, h: 0.45,
    fontSize: 20, bold: true, color: C.white, fontFace: "Calibri",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: M + 0.2, y: 0.52, w: CW - 1, h: 0.3,
      fontSize: 10, color: C.textMuted, fontFace: "Calibri",
    });
  }
};

const footerLine = (slide) =>
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.25, w: W, h: 0.035, fill: { color: C.navy },
  });

const card = (slide, x, y, w, h) => {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 4, offset: 1.5, angle: 135, color: "000000", opacity: 0.08 },
  });
};

const cardLeftAccent = (slide, x, y, w, h, color) => {
  card(slide, x, y, w, h);
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.06, h,
    fill: { color },
  });
};

const iconCircle = (slide, x, y, size, color, symbol, symColor = C.white) => {
  slide.addShape(pres.shapes.OVAL, {
    x, y, w: size, h: size,
    fill: { color },
  });
  if (symbol) {
    slide.addText(symbol, {
      x, y: y + size * 0.05, w: size, h: size * 0.9,
      fontSize: size * 0.48, color: symColor,
      align: "center", valign: "middle", fontFace: "Segoe UI Symbol",
    });
  }
};

// ─── SLIDE BUILDERS ───

/**
 * Section slide (dark divider)
 */
const sectionSlide = (title, subtitle) => {
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Background circles
  s.addShape(pres.shapes.OVAL, {
    x: -1.5, y: -1.5, w: 5, h: 5,
    fill: { color: "0D1F3C", transparency: 60 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 9.5, y: 4, w: 5, h: 5,
    fill: { color: "0D1F3C", transparency: 60 },
  });

  // Corner decorations
  const cd = (x, y) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.5, h: 0.03, fill: { color: C.gold, transparency: 40 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.03, h: 0.5, fill: { color: C.gold, transparency: 40 },
    });
  };
  cd(M, M);
  cd(W - M - 0.5, M);
  cd(M, H - M - 0.5);
  cd(W - M - 0.5, H - M - 0.5);

  // Gold dot row
  for (let i = 0; i < 18; i++) {
    s.addShape(pres.shapes.OVAL, {
      x: M + 0.3 + i * 0.18, y: H - 0.55,
      w: 0.04, h: 0.04, fill: { color: C.gold, transparency: 50 },
    });
  }

  s.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 2.2, w: 2, h: 0.04, fill: { color: C.gold },
  });
  s.addText(title, {
    x: M, y: 2.5, w: CW, h: 1.6,
    fontSize: 34, bold: true, color: C.white, fontFace: "Calibri",
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: M, y: 4.1, w: CW - 2, h: 0.4,
      fontSize: 13, color: C.textMuted, fontFace: "Calibri",
    });
  }

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.25, w: W, h: 0.04, fill: { color: C.gold },
  });
  pageNum(s);
};

/**
 * Plain content slide (for definitions, single-column lists)
 */
const contentSlide = (title, textItems, opts) => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, title, opts?.subtitle);

  // Build text runs
  const runs = textItems.map((item, i) => {
    if (typeof item === "string") {
      if (item === "") {
        return { text: "", options: { breakLine: true, fontSize: 6 } };
      }
      const isHeader = item.startsWith("$");
      const content = item.replace(/^\$/, "");
      return {
        text: content,
        options: {
          bullet: !isHeader,
          fontSize: isHeader ? 14 : 12,
          bold: isHeader,
          color: isHeader ? C.accent : C.textDark,
          fontFace: "Calibri",
          breakLine: i < textItems.length - 1,
          paraSpaceAfter: isHeader ? 6 : 2,
        },
      };
    }
    // Custom run object
    return item;
  });

  s.addText(runs, {
    x: M + 0.15, y: 1.15, w: CW - 0.3, h: 5.6,
    valign: "top", margin: 0,
  });
  footerLine(s);
  pageNum(s);
};

/**
 * Two-column card slide
 */
const twoColSlide = (title, leftItems, rightItems, opts) => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, title, opts?.subtitle);

  const colW = 5.8;
  const colGap = 0.35;
  const startX = M;
  const startY = 1.15;

  [
    { items: leftItems, color: C.accent, x: startX },
    { items: rightItems, color: opts?.rightColor || C.teal, x: startX + colW + colGap },
  ].forEach(({ items, color, x }) => {
    cardLeftAccent(s, x, startY, colW, 5.45, color);

    const runs = items.map((item, i) => {
      if (typeof item === "string") {
        if (item === "") {
          return { text: "", options: { breakLine: true, fontSize: 6 } };
        }
        const isHeader = item.startsWith("$");
        const content = item.replace(/^\$/, "");
        return {
          text: content,
          options: {
            bullet: !isHeader,
            fontSize: isHeader ? 13 : 11,
            bold: isHeader,
            color: isHeader ? color : C.textDark,
            fontFace: "Calibri",
            breakLine: i < items.length - 1,
            paraSpaceAfter: isHeader ? 4 : 1,
          },
        };
      }
      return item;
    });

    s.addText(runs, {
      x: x + 0.25, y: startY + 0.2, w: colW - 0.45, h: 5.0,
      valign: "top", margin: 0,
    });
  });

  footerLine(s);
  pageNum(s);
};

/**
 * Card grid (2, 3, or 4 cards)
 */
const cardGridSlide = (title, cards, opts) => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, title, opts?.subtitle);

  const n = cards.length;
  const colW = n === 4 ? 2.85 : n === 3 ? 3.85 : 5.85;
  const gap = 0.3;
  const totalW = n * colW + (n - 1) * gap;
  const startX = (W - totalW) / 2;
  const startY = 1.15;
  const cardH = 5.3;

  cards.forEach((c, i) => {
    const cx = startX + i * (colW + gap);
    const clr = c.color || C.accent;

    cardLeftAccent(s, cx, startY, colW, cardH, clr);

    // Icon circle
    if (c.icon) {
      iconCircle(s, cx + 0.2, startY + 0.15, 0.45, clr, c.icon);
    }
    const ty = startY + (c.icon ? 0.7 : 0.2);

    // Title
    s.addText(c.title, {
      x: cx + 0.2, y: ty, w: colW - 0.4, h: 0.35,
      fontSize: 14, bold: true, color: clr, fontFace: "Calibri",
    });

    // Items as bullet list
    const runs = c.items.map((item, j) => ({
      text: item,
      options: {
        bullet: true,
        fontSize: 11,
        color: C.textDark,
        fontFace: "Calibri",
        breakLine: j < c.items.length - 1,
        paraSpaceAfter: 1,
      },
    }));
    s.addText(runs, {
      x: cx + 0.2, y: ty + 0.4, w: colW - 0.4, h: cardH - ty + startY - 0.6,
      valign: "top", margin: 0,
    });
  });

  footerLine(s);
  pageNum(s);
};

/**
 * Table slide
 */
const tableSlide = (title, headers, rows, opts) => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, title, opts?.subtitle);

  const hRow = headers.map((h) => ({
    text: h,
    options: {
      bold: true, color: C.white,
      fill: { color: C.navy },
      fontSize: 11, fontFace: "Calibri",
      align: "center", valign: "middle",
    },
  }));

  const dRows = rows.map((r, i) =>
    r.map((c, j) => ({
      text: String(c),
      options: {
        fontSize: 11,
        color: C.textDark,
        fontFace: "Calibri",
        fill: { color: i % 2 === 0 ? C.ice : C.white },
        align: j === 0 ? "left" : "center",
        valign: "middle",
      },
    }))
  );

  s.addTable([hRow, ...dRows], {
    x: M, y: 1.15,
    w: CW,
    colW: opts?.colW,
    border: { pt: 0.5, color: C.textMuted },
    rowH: 0.42,
    margin: [2, 6, 2, 6],
  });

  footerLine(s);
  pageNum(s);
};

/**
 * Big number callout slide
 */
const calloutSlide = (title, items, opts) => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, title, opts?.subtitle);

  const n = items.length;
  const cw = n === 4 ? 2.8 : n === 3 ? 3.8 : 5.8;
  const gap = 0.35;
  const tw = n * cw + (n - 1) * gap;
  const sx = (W - tw) / 2;
  const sy = 1.3;

  items.forEach((it, i) => {
    const cx = sx + i * (cw + gap);
    const clr = it.color || C.accent;

    card(s, cx, sy, cw, 2.1);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: sy, w: cw, h: 0.05, fill: { color: clr },
    });
    s.addText(it.num, {
      x: cx, y: sy + 0.3, w: cw, h: 0.7,
      fontSize: 32, bold: true, color: clr,
      align: "center", fontFace: "Calibri",
    });
    s.addText(it.label, {
      x: cx + 0.1, y: sy + 1.05, w: cw - 0.2, h: 0.6,
      fontSize: 11, color: C.textGray,
      align: "center", valign: "middle", fontFace: "Calibri",
    });
  });

  if (opts?.extra) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: M, y: 3.8, w: CW, h: 2.5,
      fill: { color: C.ice },
    });
    s.addText(opts.extra, {
      x: M + 0.25, y: 3.9, w: CW - 0.5, h: 2.3,
      fontSize: 12, color: C.textDark, fontFace: "Calibri",
      valign: "top",
    });
  }

  footerLine(s);
  pageNum(s);
};

/**
 * Process flow slide (horizontal steps with arrows)
 */
const flowSlide = (title, steps, opts) => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, title, opts?.subtitle);

  const n = steps.length;
  const boxW = 2.6;
  const gap = 0.5;
  const totalW = n * boxW + (n - 1) * gap;
  const sx = (W - totalW) / 2;
  const sy = 1.8;

  steps.forEach((st, i) => {
    const cx = sx + i * (boxW + gap);
    const clr = st.color || C.accent;

    // Step card
    card(s, cx, sy, boxW, 2.0);
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: sy, w: boxW, h: 0.05, fill: { color: clr },
    });

    // Number circle
    iconCircle(s, cx + boxW / 2 - 0.25, sy + 0.15, 0.5, clr, st.num, C.white);

    // Title
    s.addText(st.title, {
      x: cx + 0.1, y: sy + 0.75, w: boxW - 0.2, h: 0.35,
      fontSize: 13, bold: true, color: clr,
      align: "center", fontFace: "Calibri",
    });

    // Description
    s.addText(st.desc, {
      x: cx + 0.1, y: sy + 1.1, w: boxW - 0.2, h: 0.7,
      fontSize: 10, color: C.textGray,
      align: "center", valign: "top", fontFace: "Calibri",
    });

    // Arrow
    if (i < n - 1) {
      s.addText("›", {
        x: cx + boxW, y: sy + 0.6, w: gap, h: 0.5,
        fontSize: 24, color: C.gold,
        align: "center", valign: "middle", fontFace: "Calibri",
      });
    }
  });

  // Bottom note box
  if (opts?.note) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: M, y: 4.3, w: CW, h: 1.7,
      fill: { color: C.ice },
    });
    s.addText(opts.note, {
      x: M + 0.25, y: 4.35, w: CW - 0.5, h: 1.5,
      fontSize: 11, color: C.textDark, fontFace: "Calibri",
      valign: "middle",
    });
  }

  footerLine(s);
  pageNum(s);
};

// ═══════════════════════════════════════════════════════════
// SLIDE CONTENT
// ═══════════════════════════════════════════════════════════

// ─── COVER ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Large decorative circles
  s.addShape(pres.shapes.OVAL, {
    x: -1, y: -1.5, w: 4.5, h: 4.5,
    fill: { color: "0D1F3C", transparency: 55 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 8.5, y: -2, w: 7, h: 7,
    fill: { color: "0D1F3C", transparency: 55 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 10, y: 4.5, w: 5, h: 5,
    fill: { color: "0D1F3C", transparency: 55 },
  });

  // Gold dot row
  for (let i = 0; i < 14; i++) {
    s.addShape(pres.shapes.OVAL, {
      x: M + 0.3 + i * 0.18, y: 0.3,
      w: 0.05, h: 0.05, fill: { color: C.gold },
    });
  }

  // Gold accent line
  s.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 2.6, w: 4, h: 0.04, fill: { color: C.gold },
  });

  s.addText("BERITA DAERAH", {
    x: M, y: 0.5, w: 5, h: 0.35,
    fontSize: 12, color: C.gold, bold: true, fontFace: "Calibri",
  });
  s.addText("KOTA BEKASI", {
    x: M, y: 0.85, w: 5, h: 0.35,
    fontSize: 14, color: C.white, bold: true, fontFace: "Calibri",
  });

  s.addText("PERATURAN WALI KOTA BEKASI", {
    x: M, y: 1.8, w: CW, h: 0.45,
    fontSize: 20, bold: true, color: C.white, fontFace: "Calibri",
  });
  s.addText("NOMOR 51 TAHUN 2024", {
    x: M, y: 2.15, w: CW, h: 0.4,
    fontSize: 15, color: C.gold, bold: true, fontFace: "Calibri",
  });
  s.addText("TENTANG\nPENGELOLAAN PAJAK REKLAME", {
    x: M, y: 3.1, w: CW, h: 2.0,
    fontSize: 40, bold: true, color: C.white, fontFace: "Calibri",
  });

  // Info box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 5.4, w: 7.5, h: 0.9,
    fill: { color: "0D1F3C" },
  });
  s.addText("Pemerintah Kota Bekasi  ·  20 Desember 2024", {
    x: M + 0.2, y: 5.45, w: 7, h: 0.4,
    fontSize: 12, color: C.ice, fontFace: "Calibri",
  });
  s.addText("Berlaku sejak diundangkan", {
    x: M + 0.2, y: 5.75, w: 7, h: 0.35,
    fontSize: 10, color: C.textMuted, fontFace: "Calibri",
  });

  // Gold bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.22, w: W, h: 0.22,
    fill: { color: C.gold },
  });
})();

// ─── DAFTAR ISI ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, "DAFTAR ISI", "Perwal Bekasi No 51/2024 · Pengelolaan Pajak Reklame");

  const items = [
    ["1", "Ketentuan Umum & Definisi", C.accent],
    ["2", "Objek, Subjek & Wajib Pajak", C.teal],
    ["3", "Masa Pajak & Tahun Pajak", C.warm],
    ["4", "Pendaftaran & Pendataan WP", C.navyMid],
    ["5", "Nilai Sewa Reklame (NSR)", C.red],
    ["6", "Perhitungan & Tarif Pajak", C.accent],
    ["7", "Penetapan, Tagihan & Pembayaran", C.teal],
    ["8", "Pembetulan, Keberatan & Banding", C.warm],
    ["9", "Pemeriksaan, Penagihan & Penghapusan", C.navyMid],
    ["10", "Keringanan, Kemudahan & Penghargaan", C.accent],
    ["11", "Ketentuan Penutup", C.teal],
  ];

  const colW = 5.5;
  const gap = 0.35;
  const startX = (W - 2 * colW - gap) / 2;
  const startY = 1.15;

  items.forEach(([num, label, clr], i) => {
    const col = i < 6 ? 0 : 1;
    const row = i < 6 ? i : i - 6;
    const cx = startX + col * (colW + gap);
    const cy = startY + row * 0.85;

    // Card
    card(s, cx, cy, colW, 0.65);

    // Number badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx + 0.1, y: cy + 0.08, w: 0.5, h: 0.5,
      fill: { color: clr },
    });
    s.addText(num, {
      x: cx + 0.1, y: cy + 0.1, w: 0.5, h: 0.45,
      fontSize: 14, bold: true, color: C.white,
      align: "center", valign: "middle", fontFace: "Calibri",
    });

    s.addText(label, {
      x: cx + 0.75, y: cy + 0.08, w: colW - 0.9, h: 0.5,
      fontSize: 13, color: C.textDark, fontFace: "Calibri", valign: "middle",
    });
  });

  footerLine(s);
  pageNum(s);
})();

// ═══════════════════════════════════════════
// BAB I
// ═══════════════════════════════════════════
sectionSlide("BAB I\nKETENTUAN UMUM", "Pasal 1 — Definisi Kunci");

(() => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, "Definisi Penting");

  const defs = [
    { icon: "🏛️", term: "Daerah", def: "Kota Bekasi" },
    { icon: "📊", term: "Bapenda", def: "Badan Pendapatan Daerah Kota Bekasi" },
    { icon: "📢", term: "Reklame", def: "Benda/alat/media untuk promosi & pengenalan secara komersial" },
    { icon: "💰", term: "Pajak Reklame", def: "Pajak atas penyelenggaraan reklame" },
    { icon: "📐", term: "NSR", def: "Nilai Sewa Reklame — dasar pengenaan pajak" },
    { icon: "🆔", term: "NPWPD", def: "Nomor Pokok Wajib Pajak Daerah" },
    { icon: "👤", term: "Wajib Pajak", def: "Orang pribadi/badan dengan hak & kewajiban perpajakan" },
  ];

  defs.forEach((d, i) => {
    const y = 1.15 + i * 0.78;

    // Card
    card(s, M, y, CW, 0.65);

    // Icon
    iconCircle(s, M + 0.12, y + 0.08, 0.5, C.ice, d.icon, C.navy);

    // Term
    s.addText(d.term, {
      x: M + 0.75, y: y + 0.04, w: 2.5, h: 0.3,
      fontSize: 13, bold: true, color: C.textDark, fontFace: "Calibri",
    });
    // Definition
    s.addText(d.def, {
      x: M + 0.75, y: y + 0.32, w: CW - 1, h: 0.3,
      fontSize: 11, color: C.textGray, fontFace: "Calibri",
    });
  });

  footerLine(s);
  pageNum(s);
})();

// ═══════════════════════════════════════════
// BAB II
// ═══════════════════════════════════════════
sectionSlide("BAB II\nOBJEK, SUBJEK & WAJIB PAJAK", "Pasal 2–4");

cardGridSlide("Objek Pajak Reklame", [
  {
    icon: "📋", title: "10 Jenis Reklame", color: C.accent,
    items: [
      "Papan / Billboard",
      "Videotron / Megatron",
      "Kain (Spanduk, Umbul, Baliho)",
      "Melekat / Stiker",
      "Selebaran",
      "Berjalan (Kendaraan)",
      "Udara (Balon Gas)",
      "Apung",
      "Film / Slide",
      "Peragaan",
    ],
  },
  {
    icon: "🚫", title: "Dikecualikan", color: C.teal,
    items: [
      "Internet, TV, radio, media cetak",
      "Label / merek produk",
      "Nama usaha ≤ 1 m² di tempat",
      "Reklame Pemerintah / Pemda",
      "Tempat ibadah & panti asuhan",
      "Sosial & keagamaan ≤ 30 hari",
      "Kegiatan politik (masa kampanye)",
      "Olahraga KONI ≤ 30 hari",
    ],
  },
]);

cardGridSlide("Subjek & Wajib Pajak", [
  {
    icon: "👤", title: "Subjek Pajak (Ps 3)", color: C.accent,
    items: [
      "Orang pribadi atau Badan",
      "yang menggunakan Reklame",
    ],
  },
  {
    icon: "✋", title: "Wajib Pajak (Ps 4)", color: C.teal,
    items: [
      "Orang pribadi atau Badan",
      "yang menyelenggarakan Reklame",
      "Jika pihak ketiga → menjadi WP",
    ],
  },
]);

// ═══════════════════════════════════════════
// BAB III
// ═══════════════════════════════════════════
sectionSlide("BAB III\nMASA PAJAK & TAHUN PAJAK", "Pasal 5");

calloutSlide("Masa & Tahun Pajak", [
  { num: "12", label: "Bulan masa permanen", color: C.accent },
  { num: "30", label: "Hari maks. insidentil", color: C.teal },
  { num: "1", label: "Tahun pajak", color: C.warm },
  { num: "1", label: "Bulan (bagian tahun)", color: C.navyMid },
], {
  subtitle: "Pasal 5",
  extra:
    "• Masa Pajak Permanen: 12 bulan atau sesuai jangka waktu penayangan reklame\n" +
    "• Masa Pajak Insidentil: dihitung per hari, maksimal 30 hari\n" +
    "• Tahun Pajak: 1 tahun kalender atau sesuai tahun buku wajib pajak\n" +
    "• Bagian Tahun Pajak: 1 bulan penuh (jika tidak mencakup satu tahun penuh)",
});

// ═══════════════════════════════════════════
// BAB IV
// ═══════════════════════════════════════════
sectionSlide("BAB IV\nPENDAFTARAN & PENDATAAN WP", "Pasal 6–8");

twoColSlide(
  "Pendaftaran & Pendataan",
  [
    "$PENDAFTARAN (Pasal 6)",
    "WP wajib mendaftarkan diri & objek pajak",
    "Formulir: ambil langsung / online / dikirim petugas",
    "Lampirkan: KTP, NPWP, Akta, NIB",
    "Bapenda terbitkan NPWPD",
    "Jika tidak mendaftar → NPWPD jabatan",
    "Diterbitkan juga NOPD & nomor registrasi",
  ],
  [
    "$PENDATAAN & NONAKTIF (Pasal 7–8)",
    "Bapenda mendata WP & objek pajak",
    "Termasuk data geografis",
    "Dapat kerjasama dengan instansi lain",
    "Penonaktifan: WP tak penuhi syarat",
    "Keputusan maksimal 3 bulan",
    "Syarat: tanpa tunggakan & keberatan",
  ]
);

// ═══════════════════════════════════════════
// BAB V
// ═══════════════════════════════════════════
sectionSlide("BAB V\nNILAI SEWA REKLAME", "Pasal 9 — Dasar Pengenaan Pajak");

(() => {
  const s = pres.addSlide();
  s.background = { color: C.offWhite };
  headerBar(s, "Faktor Penentu NSR");

  const factors = [
    ["1", "Jenis Reklame", C.accent],
    ["2", "Bahan", C.teal],
    ["3", "Lokasi (Kelas Jalan)", C.warm],
    ["4", "Waktu Tayang (detik)", C.navyMid],
    ["5", "Jangka Waktu (hari)", C.accent],
    ["6", "Jumlah Media", C.teal],
    ["7", "Ukuran (m²)", C.warm],
  ];

  const perRow = 4;
  const boxW = 2.8;
  const gap = 0.3;
  const totalRowW = perRow * boxW + (perRow - 1) * gap;
  const sx = (W - totalRowW) / 2;
  const sy = 1.2;

  factors.forEach(([num, label, clr], i) => {
    const col = i % perRow;
    const row = Math.floor(i / perRow);
    const cx = sx + col * (boxW + gap);
    const cy = sy + row * 1.6;

    card(s, cx, cy, boxW, 1.3);
    iconCircle(s, cx + 0.15, cy + 0.25, 0.55, clr, num, C.white);
    s.addText(label, {
      x: cx + 0.8, y: cy + 0.2, w: boxW - 1, h: 0.9,
      fontSize: 13, color: C.textDark, fontFace: "Calibri", valign: "middle",
    });
  });

  // Klasifikasi kelas jalan
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.5, w: CW, h: 1.6,
    fill: { color: C.ice },
  });

  s.addText("KLASIFIKASI KELAS JALAN", {
    x: M + 0.2, y: 4.55, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: C.navy, fontFace: "Calibri",
  });

  const cls = [
    "🏛️  Kelas Jalan Khusus — Tol | Premium 1 (A. Yani, Sudirman, dll) | Premium 2",
    "🚗  Kelas Jalan I (Kendali Ketat) — Lebar > 3 m, pusat pelayanan",
    "🏡  Kelas Jalan II (Kendali Sedang) — Lebar ≤ 3 m, jalan lingkungan",
  ];
  cls.forEach((text, i) => {
    s.addText(text, {
      x: M + 0.2, y: 4.9 + i * 0.35, w: CW - 0.4, h: 0.3,
      fontSize: 10, color: C.textDark, fontFace: "Calibri",
    });
  });

  footerLine(s);
  pageNum(s);
})();

// ═══════════════════════════════════════════
// BAB VI
// ═══════════════════════════════════════════
sectionSlide("BAB VI\nPERHITUNGAN & TARIF PAJAK", "Pasal 10");

calloutSlide("Rumus Dasar Perhitungan", [
  { num: "×", label: "Pajak = Tarif × NSR", color: C.accent },
  { num: "50%", label: "Indoor = 50% NSR", color: C.teal },
  { num: "+20%", label: "Tinggi > 15 m", color: C.warm },
  { num: "+50%", label: "Tembakau & Miras", color: C.red },
], {
  subtitle: "Pasal 10",
  extra:
    "Rumus: Pajak Reklame = Tarif Pajak × NSR\n" +
    "NSR = Nilai Kelas Jalan × Ukuran (m²) × Jumlah × Jangka Waktu\n\n" +
    "Ketentuan Khusus:\n" +
    "• Reklame indoor: NSR 50% dari NSR normal\n" +
    "• Ketinggian > 15 meter: tambahan 20%\n" +
    "• Produk tembakau & minuman beralkohol: tambahan 50%\n" +
    "• Perubahan naskah/revisi isi reklame: dikecualikan dari tambahan biaya",
});

tableSlide("NSR — Papan / Billboard", [
  "Kelas Jalan", "Zona", "NSR (Rp/m²/hari)",
], [
  ["Kelas Jalan Khusus", "Jalan Tol", "23.575"],
  ["Kelas Jalan Khusus", "Premium 1", "16.100"],
  ["Kelas Jalan Khusus", "Premium 2", "14.950"],
  ["Kelas Jalan I", "Kendali Ketat", "13.225"],
  ["Kelas Jalan II", "Kendali Sedang", "11.500"],
], { colW: [3.5, 5.5, 2.7] });

tableSlide("NSR — Megatron / Videotron", [
  "Kelas Jalan", "Zona", "NSR (/30 dtk)", "NSR (/m²/thn)",
], [
  ["Kelas Jalan Khusus", "Jalan Tol", "Rp 17,25", "13.599.900"],
  ["Kelas Jalan Khusus", "Premium 1", "Rp 13,80", "10.879.920"],
  ["Kelas Jalan Khusus", "Premium 2", "Rp 9,20", "7.253.280"],
  ["Kelas Jalan I", "Kendali Ketat", "Rp 8,05", "6.346.620"],
  ["Kelas Jalan II", "Kendali Sedang", "Rp 5,75", "4.533.300"],
], { colW: [3.5, 4, 2.5, 2.7] });

tableSlide("NSR — Kain (Spanduk / Umbul / Baliho)", [
  "Kelas Jalan", "Zona", "NSR (Rp/m²/hari)",
], [
  ["Kelas Jalan Khusus", "Jalan Tol", "30.000"],
  ["Kelas Jalan Khusus", "Premium 1", "30.000"],
  ["Kelas Jalan Khusus", "Premium 2", "25.000"],
  ["Kelas Jalan I", "Kendali Ketat", "20.000"],
  ["Kelas Jalan II", "Kendali Sedang", "19.000"],
], { colW: [3.5, 5.5, 2.7] });

cardGridSlide("NSR — Jenis Reklame Lainnya", [
  { icon: "🏷️", title: "Stiker", color: C.accent, items: ["Rp 7,5/cm²", "Min. Rp 750.000/kali"] },
  { icon: "🧱", title: "Melekat", color: C.teal, items: ["Rp 750.000/m²/tahun"] },
  { icon: "📄", title: "Selebaran", color: C.warm, items: ["Rp 600/lembar", "Min. Rp 6.000.000/kali"] },
  { icon: "🚌", title: "Berjalan", color: C.red, items: ["Rp 6.000/m²/hari", "Termasuk kendaraan"] },
]);

cardGridSlide("NSR — Jenis Lainnya (lanjutan)", [
  { icon: "🎈", title: "Udara", color: C.accent, items: ["Rp 2.400.000/sekali", "Maks. 1 bulan"] },
  { icon: "🌊", title: "Apung", color: C.teal, items: ["Rp 600.000/sekali", "Maks. 1 bulan"] },
  { icon: "🎬", title: "Film / Slide", color: C.warm, items: ["Rp 12.000/15 detik"] },
  { icon: "🎭", title: "Peragaan", color: C.red, items: ["Rp 480.000/penyelenggaraan"] },
]);

// ═══════════════════════════════════════════
// BAB VII
// ═══════════════════════════════════════════
sectionSlide("BAB VII\nPENETAPAN, TAGIHAN & PEMBAYARAN", "Pasal 11–14");

flowSlide("Alur Penetapan → Tagihan → Pembayaran", [
  { num: "1", title: "SKPD", desc: "Diterbitkan Bapenda\nMasa berlaku 5 tahun", color: C.accent },
  { num: "2", title: "Pembayaran", desc: "Lunas 1 bulan\nsejak SKPD diterima", color: C.teal },
  { num: "3", title: "Keterlambatan", desc: "Bunga 1%/bln\nDiterbitkan STPD", color: C.warm },
  { num: "4", title: "STPD", desc: "Harus lunas\n≤ 30 hari", color: C.red },
], {
  subtitle: "Pasal 11–14",
  note:
    "• Jatuh tempo: 1 bulan sejak tanggal pengiriman SKPD\n" +
    "• Pembayaran: Kas Daerah / Bank Persepsi / tempat lain yang ditunjuk\n" +
    "• Stiker sebagai tanda bukti pembayaran reklame\n" +
    "• STPD dikenakan bunga 1%/bulan (maks. 24 bulan)",
});

// ═══════════════════════════════════════════
// BAB VIII
// ═══════════════════════════════════════════
sectionSlide("BAB VIII\nPEMBETULAN, KEBERATAN & BANDING", "Pasal 15–20, 29–33");

twoColSlide(
  "Pembetulan Ketetapan",
  [
    "$PEMBETULAN (Pasal 15–20)",
    "Kesalahan tulis: nama, alamat, NPWPD",
    "Kesalahan hitung: jumlah, tarif",
    "Kekeliruan penerapan aturan",
    "1 permohonan = 1 ketetapan",
    "Keputusan maksimal 6 bulan",
    "> 6 bulan tanpa putusan → dikabulkan",
    "Dapat dilakukan berulang (Ps 20)",
    "Jenis keputusan: kabul / batal / tolak",
  ],
  [
    "$JANGKA WAKTU & SANKSI",
    "Permohonan diajukan ke Bapenda",
    "Keputusan: kabul (tambah/kurang/hapus)",
    "Keputusan: batal | tolak",
    "Pasal 19: pembetulan jabatan",
    "Pasal 20: berulang jika masih salah",
  ],
  { rightColor: C.warm }
);

twoColSlide(
  "Keberatan & Banding",
  [
    "$KEBERATAN (Pasal 29–31)",
    "Objek: SKPD, SKPDKB, SKPDKBT, dll",
    "Diajukan maks. 3 bulan sejak SKPD",
    "Sudah bayar min. yang disetujui",
    "Keputusan maks. 12 bulan",
    "Jika ditolak: denda 30%",
    "Jika dikabulkan: + bunga 0,6%/bulan",
  ],
  [
    "$BANDING (Pasal 32–33)",
    "Objek: Surat Keputusan Keberatan",
    "Ke badan peradilan pajak",
    "Maks. 3 bulan sejak keputusan",
    "Menangguhkan kewajiban bayar",
    "Jika ditolak: denda 60%",
    "Jika dikabulkan: + bunga 0,6%/bulan",
  ],
);

// ═══════════════════════════════════════════
// BAB IX
// ═══════════════════════════════════════════
sectionSlide("BAB IX\nPEMERIKSAAN, PENAGIHAN & PENGHAPUSAN", "Pasal 22–26");

cardGridSlide("Pemeriksaan & Penagihan", [
  {
    icon: "🔍", title: "Pemeriksaan (Ps 22–23)", color: C.accent,
    items: [
      "Kepala Bapenda berwenang periksa",
      "Menguji kepatuhan WP",
      "WP wajib: buka buku/dokumen",
      "Beri akses tempat & keterangan",
      "Jika tidak → pajak ditetapkan jabatan",
    ],
  },
  {
    icon: "📬", title: "Penagihan (Ps 24)", color: C.teal,
    items: [
      "Dasar: SKPD, SKPDKB, SKPDKBT",
      "STPD, SK Pembetulan/Keberatan",
      "Putusan Banding",
    ],
  },
  {
    icon: "⏳", title: "Kedaluwarsa (Ps 25)", color: C.warm,
    items: [
      "5 tahun sejak pajak terutang",
      "Tertangguh jika ada:",
      "Surat Teguran / Paksa",
      "Pengakuan utang dari WP",
    ],
  },
]);

flowSlide("Penghapusan Piutang Pajak (Pasal 26)", [
  { num: "1", title: "Penelitian", desc: "Dilakukan Bapenda", color: C.accent },
  { num: "2", title: "Penetapan", desc: "Keputusan Wali Kota", color: C.teal },
  { num: "3", title: "Koordinasi", desc: "Dengan Inspektorat", color: C.warm },
  { num: "4", title: "SK Penghapusan", desc: "Diterbitkan", color: C.navyMid },
], {
  note:
    "Syarat penghapusan piutang pajak:\n" +
    "• Piutang tidak mungkin ditagih lagi karena kedaluwarsa\n" +
    "• Ada koordinasi dengan Inspektorat Daerah\n" +
    "• Dibuktikan dengan dokumen pelaksanaan penagihan",
});

// ═══════════════════════════════════════════
// BAB X
// ═══════════════════════════════════════════
sectionSlide("BAB X\nKERINGANAN, KEMUDAHAN & PENGHARGAAN", "Pasal 27–28, 34–35");

cardGridSlide("Fasilitas & Penghargaan", [
  {
    icon: "🎯", title: "Keringanan (Ps 27)", color: C.accent,
    items: [
      "Keringanan / Pengurangan",
      "Pembebasan / Penundaan",
      "Atas pokok & sanksi pajak",
      "WP dengan likuiditas rendah",
      "Objek terdampak bencana/kebakaran",
    ],
  },
  {
    icon: "🤝", title: "Kemudahan (Ps 28)", color: C.teal,
    items: [
      "Perpanjangan waktu bayar",
      "Angsuran maks. 24 bulan",
      "Bunga 0,6%/bulan",
      "Keadaan kahar:",
      "bencana, kebakaran, wabah, kerusuhan",
    ],
  },
  {
    icon: "🏆", title: "Penghargaan (Ps 34–35)", color: C.warm,
    items: [
      "WP Taat Pajak",
      "Bayar tepat waktu ≥ 1 tahun",
      "Tanpa tunggakan 3 tahun",
      "Kontribusi signifikan",
      "Piagam / Hadiah (APBD)",
    ],
  },
]);

// ═══════════════════════════════════════════
// BAB XI
// ═══════════════════════════════════════════
sectionSlide("BAB XI\nKETENTUAN PENUTUP", "Pasal 36–37");

twoColSlide(
  "Pencabutan & Mulai Berlaku",
  [
    "$PERATURAN YANG DICABUT (Pasal 36)",
    "Perwal No. 48 Tahun 2012",
    "Petunjuk Pelaksanaan Perda No. 14/2012",
    "Perwal No. 52 Tahun 2013",
    "Perubahan atas Perwal 48/2012",
  ],
  [
    "$MULAI BERLAKU (Pasal 37)",
    "Sejak diundangkan",
    "20 Desember 2024",
    "",
    "Pj. WALI KOTA BEKASI,",
    "ttd.",
    "R. GANI MUHAMAD",
  ],
  { rightColor: C.navyMid }
);

// ─── CLOSING ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.OVAL, {
    x: -1, y: -1.5, w: 4.5, h: 4.5,
    fill: { color: "0D1F3C", transparency: 55 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 8.5, y: -2, w: 7, h: 7,
    fill: { color: "0D1F3C", transparency: 55 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 10, y: 4.5, w: 5, h: 5,
    fill: { color: "0D1F3C", transparency: 55 },
  });

  // Gold dots
  for (let i = 0; i < 12; i++) {
    s.addShape(pres.shapes.OVAL, {
      x: M + 0.3 + i * 0.18, y: 0.3,
      w: 0.05, h: 0.05, fill: { color: C.gold },
    });
  }

  // Corner decos
  const cd = (x, y) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.5, h: 0.03, fill: { color: C.gold, transparency: 40 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.03, h: 0.5, fill: { color: C.gold, transparency: 40 },
    });
  };
  cd(M, M);
  cd(W - M - 0.5, M);

  s.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 3.6, w: 3.5, h: 0.04, fill: { color: C.gold },
  });

  s.addText("BERITA DAERAH KOTA BEKASI", {
    x: M, y: 1.6, w: CW, h: 0.4,
    fontSize: 14, color: C.gold, bold: true, fontFace: "Calibri",
  });
  s.addText("TERIMA KASIH", {
    x: M, y: 2.4, w: CW, h: 1.5,
    fontSize: 48, bold: true, color: C.white, fontFace: "Calibri",
  });
  s.addText([
    { text: "Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024", options: { breakLine: true, fontSize: 14, color: C.ice } },
    { text: "Tentang Pengelolaan Pajak Reklame", options: { fontSize: 14, color: C.ice } },
  ], {
    x: M, y: 4.1, w: CW, h: 0.8,
    fontFace: "Calibri",
  });
  s.addText("Sumber: https://jdih.bekasikota.go.id", {
    x: M, y: 5.1, w: CW, h: 0.35,
    fontSize: 10, color: C.textMuted, fontFace: "Calibri",
  });

  // Dots footer
  for (let i = 0; i < 22; i++) {
    s.addShape(pres.shapes.OVAL, {
      x: M + 0.3 + i * 0.16, y: H - 0.55,
      w: 0.04, h: 0.04, fill: { color: C.gold, transparency: 50 },
    });
  }

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.22, w: W, h: 0.22,
    fill: { color: C.gold },
  });
})();

// ═══════════════════════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════════════════════
const path = require("path");
const out = path.resolve(__dirname, "Perwal_Bekasi_51_2024_Pajak_Reklame.pptx");
pres.writeFile({ fileName: out })
  .then(() => console.log(`✅ SUKSES: ${out} (${pres.slides.length} slide)`))
  .catch((e) => console.error("❌ GAGAL:", e));
