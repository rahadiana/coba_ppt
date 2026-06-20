#!/usr/bin/env node
/**
 * Perwal Bekasi No 51/2024 — Pajak Reklame
 * VERSI FINAL — Clean, reliable, professional
 *
 * - No shadow on shapes (avoid rendering issues)
 * - Proper table colW matching width
 * - Clean card design with simple accent bars
 * - Reliable PptxGenJS API usage (v4.0.1)
 */

const PptxGenJS = require("pptxgenjs");
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Pemerintah Kota Bekasi";
pres.title = "Perwal Bekasi No 51/2024 - Pengelolaan Pajak Reklame";

// ─── COLORS ───
const C = {
  navy:      "0A1628",
  navyL:     "12294A",
  navyM:     "1B3A6B",
  gold:      "C8962E",
  goldL:     "D4A017",
  white:     "FFFFFF",
  offW:      "F5F7FA",
  ice:       "E8EDF5",
  txDk:      "1A1A2E",
  txGy:      "6B7288",
  txMu:      "9CA3AF",
  bl:        "2563EB",
  teal:      "0D9488",
  warm:      "B8860B",
  red:       "DC2626",
};

const W = 13.333;
const H = 7.5;
const M = 0.5;
const CW = W - 2 * M; // 12.333

let pg = 0;
const pn = (s) => { pg++; s.addText(String(pg), { x:W-0.9, y:H-0.4, w:0.7, h:0.25, fontSize:8, color:C.txMu, align:"right", fontFace:"Calibri" }); };

// ─── HELPERS ───

const goldTop = (s) => s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:W, h:0.03, fill:{ color:C.gold } });
const navyBot = (s) => s.addShape(pres.shapes.RECTANGLE, { x:0, y:H-0.22, w:W, h:0.03, fill:{ color:C.navy } });
const goldBot = (s) => s.addShape(pres.shapes.RECTANGLE, { x:0, y:H-0.22, w:W, h:0.22, fill:{ color:C.gold } });

const hdr = (s, title, sub) => {
  goldTop(s);
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0.03, w:W, h:0.8, fill:{ color:C.navy } });
  s.addShape(pres.shapes.RECTANGLE, { x:M, y:0.12, w:0.07, h:0.5, fill:{ color:C.gold } });
  s.addText(title, { x:M+0.2, y:0.12, w:CW-1, h:0.42, fontSize:20, bold:true, color:C.white, fontFace:"Calibri" });
  if (sub) s.addText(sub, { x:M+0.2, y:0.52, w:CW-1, h:0.25, fontSize:10, color:C.txMu, fontFace:"Calibri" });
};

const card = (s, x, y, w, h) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill:{ color:C.white }, line:{ color:C.ice, width:0.5 } });
};

const accentL = (s, x, y, w, h, color) => {
  card(s, x, y, w, h);
  s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.05, h, fill:{ color } });
};

const icn = (s, x, y, sz, color, sym) => {
  s.addShape(pres.shapes.OVAL, { x, y, w:sz, h:sz, fill:{ color } });
  if (sym) s.addText(sym, { x, y:y+sz*0.05, w:sz, h:sz*0.9, fontSize:sz*0.45, color:C.white, align:"center", valign:"middle", fontFace:"Segoe UI Symbol" });
};

// ─── SECTION SLIDE ───
const sec = (title, sub) => {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.OVAL, { x:-1.5, y:-1.5, w:5, h:5, fill:{ color:"0D1F3C", transparency:55 } });
  s.addShape(pres.shapes.OVAL, { x:9.5, y:4, w:5, h:5, fill:{ color:"0D1F3C", transparency:55 } });
  s.addShape(pres.shapes.RECTANGLE, { x:M, y:2.3, w:2.5, h:0.04, fill:{ color:C.gold } });
  s.addText(title, { x:M, y:2.6, w:CW, h:1.8, fontSize:34, bold:true, color:C.white, fontFace:"Calibri" });
  if (sub) s.addText(sub, { x:M, y:4.3, w:CW-2, h:0.4, fontSize:13, color:C.txMu, fontFace:"Calibri" });
  goldBot(s);
  pn(s);
};

// ─── CONTENT SLIDE (single column) ───
const cont = (title, lines, opts) => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, title, opts?.sub);
  const runs = lines.map((ln, i) => {
    if (ln === "") return { text:"", options:{ breakLine:true, fontSize:6 } };
    const h = ln.startsWith("$");
    return { text:ln.replace(/^\$/,""), options:{ bullet:!h, fontSize:h?14:12, bold:h, color:h?C.bl:C.txDk, fontFace:"Calibri", breakLine:i<lines.length-1, paraSpaceAfter:h?6:2 } };
  });
  s.addText(runs, { x:M+0.1, y:1.1, w:CW-0.2, h:5.7, valign:"top", margin:0 });
  navyBot(s); pn(s);
};

// ─── TWO COLUMN ───
const twoCol = (title, L, R, opts) => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, title, opts?.sub);
  const cw = 5.8;
  [{ items:L, clr:C.bl, x:M }, { items:R, clr:opts?.rc||C.teal, x:M+cw+0.35 }].forEach(({ items, clr, x }) => {
    accentL(s, x, 1.1, cw, 5.4, clr);
    const runs = items.map((ln, i) => {
      if (ln === "") return { text:"", options:{ breakLine:true, fontSize:6 } };
      const h = ln.startsWith("$");
      return { text:ln.replace(/^\$/,""), options:{ bullet:!h, fontSize:h?13:11, bold:h, color:h?clr:C.txDk, fontFace:"Calibri", breakLine:i<items.length-1, paraSpaceAfter:h?4:1 } };
    });
    s.addText(runs, { x:x+0.2, y:1.3, w:cw-0.35, h:5.0, valign:"top", margin:0 });
  });
  navyBot(s); pn(s);
};

// ─── CARD GRID ───
const grid = (title, cards, opts) => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, title, opts?.sub);
  const n = cards.length;
  const cw = n===4 ? 2.85 : n===3 ? 3.85 : 5.85;
  const gap = 0.3;
  const sx = (W - n*cw - (n-1)*gap)/2;
  const sy = 1.1;
  const ch = 5.4;
  cards.forEach((c, i) => {
    const cx = sx + i*(cw+gap);
    const clr = c.clr || C.bl;
    accentL(s, cx, sy, cw, ch, clr);
    if (c.ic) icn(s, cx+0.15, sy+0.15, 0.4, clr, c.ic);
    const ty = sy + (c.ic ? 0.65 : 0.15);
    s.addText(c.t, { x:cx+0.15, y:ty, w:cw-0.3, h:0.35, fontSize:14, bold:true, color:clr, fontFace:"Calibri" });
    const runs = c.items.map((it, j) => ({ text:it, options:{ bullet:true, fontSize:11, color:C.txDk, fontFace:"Calibri", breakLine:j<c.items.length-1, paraSpaceAfter:1 } }));
    s.addText(runs, { x:cx+0.15, y:ty+0.4, w:cw-0.3, h:ch-ty+sy-0.6, valign:"top", margin:0 });
  });
  navyBot(s); pn(s);
};

// ─── TABLE ───
const tbl = (title, hds, rows, opts) => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, title, opts?.sub);
  const tw = 11.7; // table width
  const colW = opts.cw;
  s.addTable(
    [hds.map(h => ({ text:h, options:{ bold:true, color:C.white, fill:{ color:C.navy }, fontSize:11, fontFace:"Calibri", align:"center", valign:"middle" } })),
     ...rows.map((r,i) => r.map((c,j) => ({ text:String(c), options:{ fontSize:11, color:C.txDk, fontFace:"Calibri", fill:{ color:i%2===0?C.ice:C.white }, align:j===0?"left":"center", valign:"middle" } })))],
    { x:M, y:1.15, w:tw, colW, border:{ pt:0.5, color:C.txMu }, rowH:0.42, margin:[2,6,2,6] }
  );
  navyBot(s); pn(s);
};

// ─── CALLOUT ───
const callout = (title, items, opts) => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, title, opts?.sub);
  const n = items.length;
  const cw = n===4 ? 2.8 : n===3 ? 3.8 : 5.8;
  const gap = 0.35;
  const sx = (W - n*cw - (n-1)*gap)/2;
  items.forEach((it, i) => {
    const cx = sx + i*(cw+gap);
    const clr = it.clr || C.bl;
    card(s, cx, 1.3, cw, 2.1);
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:1.3, w:cw, h:0.05, fill:{ color:clr } });
    s.addText(it.num, { x:cx, y:1.6, w:cw, h:0.7, fontSize:32, bold:true, color:clr, align:"center", fontFace:"Calibri" });
    s.addText(it.lb, { x:cx+0.1, y:2.35, w:cw-0.2, h:0.6, fontSize:11, color:C.txGy, align:"center", valign:"middle", fontFace:"Calibri" });
  });
  if (opts?.ex) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:M, y:3.8, w:CW, h:2.5, fill:{ color:C.ice } });
    s.addText(opts.ex, { x:M+0.2, y:3.9, w:CW-0.4, h:2.3, fontSize:12, color:C.txDk, fontFace:"Calibri", valign:"top" });
  }
  navyBot(s); pn(s);
};

// ─── FLOW ───
const flow = (title, steps, opts) => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, title, opts?.sub);
  const n = steps.length;
  const bw = 2.6, gap = 0.5;
  const sx = (W - n*bw - (n-1)*gap)/2;
  steps.forEach((st, i) => {
    const cx = sx + i*(bw+gap);
    const clr = st.clr || C.bl;
    card(s, cx, 1.8, bw, 2.0);
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y:1.8, w:bw, h:0.05, fill:{ color:clr } });
    icn(s, cx+bw/2-0.25, 1.95, 0.5, clr, st.num);
    s.addText(st.t, { x:cx+0.1, y:2.55, w:bw-0.2, h:0.35, fontSize:13, bold:true, color:clr, align:"center", fontFace:"Calibri" });
    s.addText(st.d, { x:cx+0.1, y:2.9, w:bw-0.2, h:0.7, fontSize:10, color:C.txGy, align:"center", valign:"top", fontFace:"Calibri" });
    if (i < n-1) s.addText("›", { x:cx+bw, y:2.4, w:gap, h:0.5, fontSize:24, color:C.gold, align:"center", valign:"middle", fontFace:"Calibri" });
  });
  if (opts?.note) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:M, y:4.3, w:CW, h:1.7, fill:{ color:C.ice } });
    s.addText(opts.note, { x:M+0.2, y:4.35, w:CW-0.4, h:1.5, fontSize:11, color:C.txDk, fontFace:"Calibri", valign:"middle" });
  }
  navyBot(s); pn(s);
};

// ═══════════════════════════════════════════════════
// SLIDES
// ═══════════════════════════════════════════════════

// ─── COVER ───
(() => {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.OVAL, { x:-1, y:-1.5, w:4.5, h:4.5, fill:{ color:"0D1F3C", transparency:55 } });
  s.addShape(pres.shapes.OVAL, { x:8.5, y:-2, w:7, h:7, fill:{ color:"0D1F3C", transparency:55 } });
  s.addShape(pres.shapes.OVAL, { x:10, y:4.5, w:5, h:5, fill:{ color:"0D1F3C", transparency:55 } });
  s.addShape(pres.shapes.RECTANGLE, { x:M, y:2.6, w:4, h:0.04, fill:{ color:C.gold } });
  s.addText("BERITA DAERAH", { x:M, y:0.5, w:5, h:0.35, fontSize:12, color:C.gold, bold:true, fontFace:"Calibri" });
  s.addText("KOTA BEKASI", { x:M, y:0.85, w:5, h:0.35, fontSize:14, color:C.white, bold:true, fontFace:"Calibri" });
  s.addText("PERATURAN WALI KOTA BEKASI", { x:M, y:1.8, w:CW, h:0.45, fontSize:20, bold:true, color:C.white, fontFace:"Calibri" });
  s.addText("NOMOR 51 TAHUN 2024", { x:M, y:2.15, w:CW, h:0.4, fontSize:15, color:C.gold, bold:true, fontFace:"Calibri" });
  s.addText("TENTANG\nPENGELOLAAN PAJAK REKLAME", { x:M, y:3.1, w:CW, h:2.0, fontSize:40, bold:true, color:C.white, fontFace:"Calibri" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:M, y:5.4, w:7.5, h:0.9, fill:{ color:"0D1F3C" } });
  s.addText("Pemerintah Kota Bekasi  ·  20 Desember 2024", { x:M+0.2, y:5.45, w:7, h:0.4, fontSize:12, color:C.ice, fontFace:"Calibri" });
  s.addText("Berlaku sejak diundangkan", { x:M+0.2, y:5.75, w:7, h:0.35, fontSize:10, color:C.txMu, fontFace:"Calibri" });
  goldBot(s);
})();

// ─── DAFTAR ISI ───
(() => {
  const s = pres.addSlide(); s.background = { color:C.offW }; hdr(s, "DAFTAR ISI", "Perwal Bekasi No 51/2024");
  const items = [
    ["1","Ketentuan Umum & Definisi",C.bl],["2","Objek, Subjek & Wajib Pajak",C.teal],["3","Masa Pajak & Tahun Pajak",C.warm],
    ["4","Pendaftaran & Pendataan WP",C.navyM],["5","Nilai Sewa Reklame (NSR)",C.red],["6","Perhitungan & Tarif Pajak",C.bl],
    ["7","Penetapan, Tagihan & Pembayaran",C.teal],["8","Pembetulan, Keberatan & Banding",C.warm],["9","Pemeriksaan, Penagihan & Penghapusan",C.navyM],
    ["10","Keringanan, Kemudahan & Penghargaan",C.bl],["11","Ketentuan Penutup",C.teal],
  ];
  const cw=5.5, gap=0.35, sx=(W-2*cw-gap)/2;
  items.forEach(([num,lb,clr],i)=>{
    const col=i<6?0:1, row=i<6?i:i-6, x=sx+col*(cw+gap), y=1.1+row*0.85;
    card(s,x,y,cw,0.65);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:x+0.1,y:y+0.08,w:0.5,h:0.5,fill:{color:clr}});
    s.addText(num,{x:x+0.1,y:y+0.1,w:0.5,h:0.45,fontSize:14,bold:true,color:C.white,align:"center",valign:"middle",fontFace:"Calibri"});
    s.addText(lb,{x:x+0.75,y:y+0.08,w:cw-0.9,h:0.5,fontSize:13,color:C.txDk,fontFace:"Calibri",valign:"middle"});
  });
  navyBot(s); pn(s);
})();

// ═══════════════════════════════════════════
// BAB I
// ═══════════════════════════════════════════
sec("BAB I\nKETENTUAN UMUM","Pasal 1 — Definisi Kunci");
(()=>{
  const s=pres.addSlide();s.background={color:C.offW};hdr(s,"Definisi Penting");
  [["🏛️","Daerah","Kota Bekasi"],["📊","Bapenda","Badan Pendapatan Daerah Kota Bekasi"],["📢","Reklame","Media untuk promosi & pengenalan komersial"],["💰","Pajak Reklame","Pajak atas penyelenggaraan reklame"],["📐","NSR","Nilai Sewa Reklame — dasar pengenaan pajak"],["🆔","NPWPD","Nomor Pokok Wajib Pajak Daerah"],["👤","Wajib Pajak","Orang pribadi/badan dengan hak & kewajiban pajak"]].forEach(([ic,term,def],i)=>{
    const y=1.1+i*0.78;
    card(s,M,y,CW,0.65);
    icn(s,M+0.12,y+0.08,0.5,C.ice,ic);
    s.addText(term,{x:M+0.75,y:y+0.04,w:2.5,h:0.3,fontSize:13,bold:true,color:C.txDk,fontFace:"Calibri"});
    s.addText(def,{x:M+0.75,y:y+0.32,w:CW-1,h:0.3,fontSize:11,color:C.txGy,fontFace:"Calibri"});
  });
  navyBot(s);pn(s);
})();

// ═══════════════════════════════════════════
// BAB II
// ═══════════════════════════════════════════
sec("BAB II\nOBJEK, SUBJEK & WAJIB PAJAK","Pasal 2–4");
grid("Objek Pajak Reklame",[
  {ic:"📋",t:"10 Jenis Reklame",clr:C.bl,items:["Papan / Billboard","Videotron / Megatron","Kain (Spanduk, Umbul, Baliho)","Melekat / Stiker","Selebaran","Berjalan (Kendaraan)","Udara (Balon Gas)","Apung","Film / Slide","Peragaan"]},
  {ic:"🚫",t:"Dikecualikan",clr:C.teal,items:["Internet, TV, radio, media cetak","Label / merek produk","Nama usaha ≤ 1 m² di tempat","Reklame Pemerintah / Pemda","Tempat ibadah & panti asuhan","Sosial & keagamaan ≤ 30 hari","Kegiatan politik (masa kampanye)","Olahraga KONI ≤ 30 hari"]},
]);
grid("Subjek & Wajib Pajak",[
  {ic:"👤",t:"Subjek Pajak (Pasal 3)",clr:C.bl,items:["Orang pribadi atau Badan","yang menggunakan Reklame"]},
  {ic:"✋",t:"Wajib Pajak (Pasal 4)",clr:C.teal,items:["Orang pribadi atau Badan","yang menyelenggarakan Reklame","Jika pihak ketiga → menjadi WP"]},
]);

// ═══════════════════════════════════════════
// BAB III
// ═══════════════════════════════════════════
sec("BAB III\nMASA PAJAK & TAHUN PAJAK","Pasal 5");
callout("Masa & Tahun Pajak",[
  {num:"12",lb:"Bulan (Permanen)",clr:C.bl},
  {num:"30",lb:"Hari (Insidentil)",clr:C.teal},
  {num:"1",lb:"Tahun Pajak",clr:C.warm},
  {num:"1",lb:"Bulan (Bagian Tahun)",clr:C.navyM},
],{sub:"Pasal 5",ex:"• Masa Pajak Permanen: 12 bulan atau sesuai jangka waktu penayangan reklame\n• Masa Pajak Insidentil: dihitung per hari, maksimal 30 hari\n• Tahun Pajak: 1 tahun kalender atau sesuai tahun buku wajib pajak\n• Bagian Tahun Pajak: 1 bulan penuh (jika tidak mencakup satu tahun penuh)"});

// ═══════════════════════════════════════════
// BAB IV
// ═══════════════════════════════════════════
sec("BAB IV\nPENDAFTARAN & PENDATAAN WP","Pasal 6–8");
twoCol("Pendaftaran & Pendataan",[
  "$PENDAFTARAN (Pasal 6)","","WP wajib mendaftarkan diri & objek pajak","Formulir: ambil/online/dikirim petugas","Lampirkan: KTP, NPWP, Akta, NIB","Bapenda terbitkan NPWPD","Jika tidak mendaftar → NPWPD jabatan","Juga: NOPD & nomor registrasi",
],[
  "$PENDATAAN & NONAKTIF (Pasal 7–8)","","Bapenda mendata WP & objek pajak","Termasuk data geografis","Dapat kerjasama dengan instansi lain","Penonaktifan: WP tak penuhi syarat","Keputusan maksimal 3 bulan","Syarat: tanpa tunggakan & keberatan",
]);

// ═══════════════════════════════════════════
// BAB V
// ═══════════════════════════════════════════
sec("BAB V\nNILAI SEWA REKLAME","Pasal 9 — Dasar Pengenaan Pajak");
(()=>{
  const s=pres.addSlide();s.background={color:C.offW};hdr(s,"Faktor Penentu NSR");
  const f=[["1","Jenis Reklame",C.bl],["2","Bahan",C.teal],["3","Lokasi (Kelas Jalan)",C.warm],["4","Waktu Tayang (detik)",C.navyM],["5","Jangka Waktu (hari)",C.bl],["6","Jumlah Media",C.teal],["7","Ukuran (m²)",C.warm]];
  const bw=2.8,gap=0.3,sx=(W-4*bw-3*gap)/2;
  f.forEach(([num,lb,clr],i)=>{
    const x=sx+(i%4)*(bw+gap),y=1.1+Math.floor(i/4)*1.6;
    card(s,x,y,bw,1.3);
    icn(s,x+0.15,y+0.25,0.55,clr,num);
    s.addText(lb,{x:x+0.8,y:y+0.2,w:bw-1,h:0.9,fontSize:13,color:C.txDk,fontFace:"Calibri",valign:"middle"});
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:M,y:4.4,w:CW,h:1.7,fill:{color:C.ice}});
  s.addText("KLASIFIKASI KELAS JALAN",{x:M+0.2,y:4.45,w:5,h:0.3,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri"});
  ["🏛️  Kelas Jalan Khusus — Tol | Premium 1 | Premium 2","🚗  Kelas Jalan I (Kendali Ketat) — Lebar > 3 m, pusat pelayanan","🏡  Kelas Jalan II (Kendali Sedang) — Lebar ≤ 3 m, jalan lingkungan"].forEach((t,i)=>{
    s.addText(t,{x:M+0.2,y:4.8+i*0.35,w:CW-0.4,h:0.3,fontSize:10,color:C.txDk,fontFace:"Calibri"});
  });
  navyBot(s);pn(s);
})();

// ═══════════════════════════════════════════
// BAB VI
// ═══════════════════════════════════════════
sec("BAB VI\nPERHITUNGAN & TARIF PAJAK","Pasal 10");
callout("Rumus Dasar Perhitungan",[
  {num:"×",lb:"Pajak = Tarif × NSR",clr:C.bl},
  {num:"50%",lb:"Indoor = 50% NSR",clr:C.teal},
  {num:"+20%",lb:"Tinggi > 15 m",clr:C.warm},
  {num:"+50%",lb:"Tembakau & Miras",clr:C.red},
],{sub:"Pasal 10",ex:"Rumus: Pajak Reklame = Tarif Pajak × NSR\nNSR = Nilai Kelas Jalan × Ukuran (m²) × Jumlah × Jangka Waktu\n\nKetentuan Khusus:\n• Reklame indoor: NSR 50% dari NSR normal\n• Ketinggian > 15 meter: tambahan 20%\n• Produk tembakau & minuman beralkohol: tambahan 50%\n• Perubahan naskah/revisi isi reklame: dikecualikan"});

tbl("NSR — Papan / Billboard",["Kelas Jalan","Zona","NSR (Rp/m²/hari)"],[
  ["Kelas Jalan Khusus","Jalan Tol","23.575"],["Kelas Jalan Khusus","Premium 1","16.100"],["Kelas Jalan Khusus","Premium 2","14.950"],
  ["Kelas Jalan I","Kendali Ketat","13.225"],["Kelas Jalan II","Kendali Sedang","11.500"],
],{cw:[3.5,5.5,2.7]});

tbl("NSR — Megatron / Videotron",["Kelas Jalan","Zona","NSR (/30 dtk)","NSR (/m²/thn)"],[
  ["Kelas Jalan Khusus","Jalan Tol","Rp 17,25","13.599.900"],["Kelas Jalan Khusus","Premium 1","Rp 13,80","10.879.920"],
  ["Kelas Jalan Khusus","Premium 2","Rp 9,20","7.253.280"],["Kelas Jalan I","Kendali Ketat","Rp 8,05","6.346.620"],
  ["Kelas Jalan II","Kendali Sedang","Rp 5,75","4.533.300"],
],{cw:[3.5,4,2.5,2.7]});

tbl("NSR — Kain (Spanduk/Umbul/Baliho)",["Kelas Jalan","Zona","NSR (Rp/m²/hari)"],[
  ["Kelas Jalan Khusus","Jalan Tol","30.000"],["Kelas Jalan Khusus","Premium 1","30.000"],["Kelas Jalan Khusus","Premium 2","25.000"],
  ["Kelas Jalan I","Kendali Ketat","20.000"],["Kelas Jalan II","Kendali Sedang","19.000"],
],{cw:[3.5,5.5,2.7]});

grid("NSR — Jenis Reklame Lainnya",[
  {ic:"🏷️",t:"Stiker",clr:C.bl,items:["Rp 7,5/cm²","Min. Rp 750.000/kali"]},
  {ic:"🧱",t:"Melekat",clr:C.teal,items:["Rp 750.000/m²/tahun"]},
  {ic:"📄",t:"Selebaran",clr:C.warm,items:["Rp 600/lembar","Min. Rp 6.000.000/kali"]},
  {ic:"🚌",t:"Berjalan",clr:C.red,items:["Rp 6.000/m²/hari","Termasuk kendaraan"]},
]);
grid("NSR — Jenis Lainnya (lanjutan)",[
  {ic:"🎈",t:"Udara",clr:C.bl,items:["Rp 2.400.000/sekali","Maks. 1 bulan"]},
  {ic:"🌊",t:"Apung",clr:C.teal,items:["Rp 600.000/sekali","Maks. 1 bulan"]},
  {ic:"🎬",t:"Film / Slide",clr:C.warm,items:["Rp 12.000/15 detik"]},
  {ic:"🎭",t:"Peragaan",clr:C.red,items:["Rp 480.000/penyelenggaraan"]},
]);

// ═══════════════════════════════════════════
// BAB VII
// ═══════════════════════════════════════════
sec("BAB VII\nPENETAPAN, TAGIHAN & PEMBAYARAN","Pasal 11–14");
flow("Alur Penetapan → Tagihan → Pembayaran",[
  {num:"1",t:"SKPD",d:"Diterbitkan Bapenda\nMasa berlaku 5 tahun",clr:C.bl},
  {num:"2",t:"Pembayaran",d:"Lunas 1 bulan\nsejak SKPD diterima",clr:C.teal},
  {num:"3",t:"Keterlambatan",d:"Bunga 1%/bln\nDiterbitkan STPD",clr:C.warm},
  {num:"4",t:"STPD",d:"Harus lunas\n≤ 30 hari",clr:C.red},
],{sub:"Pasal 11–14",note:"• Jatuh tempo: 1 bulan sejak tanggal pengiriman SKPD\n• Pembayaran: Kas Daerah / Bank Persepsi / tempat lain yang ditunjuk\n• Stiker sebagai tanda bukti pembayaran reklame\n• STPD dikenakan bunga 1%/bulan (maks. 24 bulan)"});

// ═══════════════════════════════════════════
// BAB VIII
// ═══════════════════════════════════════════
sec("BAB VIII\nPEMBETULAN, KEBERATAN & BANDING","Pasal 15–20, 29–33");
twoCol("Pembetulan Ketetapan",[
  "$PEMBETULAN (Pasal 15–20)","","Kesalahan tulis: nama, alamat, NPWPD","Kesalahan hitung: jumlah, tarif","Kekeliruan penerapan aturan","1 permohonan = 1 ketetapan","Keputusan maksimal 6 bulan","> 6 bulan tanpa putusan → dikabulkan","Dapat dilakukan berulang (Ps 20)","Jenis keputusan: kabul / batal / tolak",
],[
  "$JANGKA WAKTU & SANKSI","","Permohonan diajukan ke Bapenda","Keputusan: kabul (tambah/kurang/hapus)","Keputusan: batal | tolak","Pasal 19: pembetulan jabatan","Pasal 20: berulang jika masih salah",
],{rc:C.warm});

twoCol("Keberatan & Banding",[
  "$KEBERATAN (Pasal 29–31)","","Objek: SKPD, SKPDKB, SKPDKBT, dll","Diajukan maks. 3 bulan sejak SKPD","Sudah bayar min. yang disetujui","Keputusan maks. 12 bulan","Jika ditolak: denda 30%","Jika dikabulkan: + bunga 0,6%/bulan",
],[
  "$BANDING (Pasal 32–33)","","Objek: Surat Keputusan Keberatan","Ke badan peradilan pajak","Maks. 3 bulan sejak keputusan","Menangguhkan kewajiban bayar","Jika ditolak: denda 60%","Jika dikabulkan: + bunga 0,6%/bulan",
]);

// ═══════════════════════════════════════════
// BAB IX
// ═══════════════════════════════════════════
sec("BAB IX\nPEMERIKSAAN, PENAGIHAN & PENGHAPUSAN","Pasal 22–26");
grid("Pemeriksaan & Penagihan",[
  {ic:"🔍",t:"Pemeriksaan (Ps 22–23)",clr:C.bl,items:["Kepala Bapenda berwenang periksa","Menguji kepatuhan WP","WP wajib: buka buku/dokumen","Beri akses tempat & keterangan","Jika tidak → pajak ditetapkan jabatan"]},
  {ic:"📬",t:"Penagihan (Ps 24)",clr:C.teal,items:["Dasar: SKPD, SKPDKB, SKPDKBT","STPD, SK Pembetulan/Keberatan","Putusan Banding"]},
  {ic:"⏳",t:"Kedaluwarsa (Ps 25)",clr:C.warm,items:["5 tahun sejak pajak terutang","Tertangguh jika ada:","Surat Teguran / Paksa","Pengakuan utang dari WP"]},
]);
flow("Penghapusan Piutang Pajak (Pasal 26)",[
  {num:"1",t:"Penelitian",d:"Dilakukan Bapenda",clr:C.bl},
  {num:"2",t:"Penetapan",d:"Keputusan Wali Kota",clr:C.teal},
  {num:"3",t:"Koordinasi",d:"Dengan Inspektorat",clr:C.warm},
  {num:"4",t:"SK Penghapusan",d:"Diterbitkan",clr:C.navyM},
],{note:"Syarat penghapusan piutang pajak:\n• Piutang tidak mungkin ditagih lagi karena kedaluwarsa\n• Ada koordinasi dengan Inspektorat Daerah\n• Dibuktikan dengan dokumen pelaksanaan penagihan"});

// ═══════════════════════════════════════════
// BAB X
// ═══════════════════════════════════════════
sec("BAB X\nKERINGANAN, KEMUDAHAN & PENGHARGAAN","Pasal 27–28, 34–35");
grid("Fasilitas & Penghargaan",[
  {ic:"🎯",t:"Keringanan (Ps 27)",clr:C.bl,items:["Keringanan / Pengurangan","Pembebasan / Penundaan","Atas pokok & sanksi pajak","WP dengan likuiditas rendah","Objek terdampak bencana/kebakaran"]},
  {ic:"🤝",t:"Kemudahan (Ps 28)",clr:C.teal,items:["Perpanjangan waktu bayar","Angsuran maks. 24 bulan","Bunga 0,6%/bulan","Keadaan kahar: bencana, wabah, kerusuhan"]},
  {ic:"🏆",t:"Penghargaan (Ps 34–35)",clr:C.warm,items:["WP Taat Pajak","Bayar tepat waktu ≥ 1 tahun","Tanpa tunggakan 3 tahun","Kontribusi signifikan","Piagam / Hadiah (APBD)"]},
]);

// ═══════════════════════════════════════════
// BAB XI
// ═══════════════════════════════════════════
sec("BAB XI\nKETENTUAN PENUTUP","Pasal 36–37");
twoCol("Pencabutan & Mulai Berlaku",[
  "$PERATURAN YANG DICABUT (Pasal 36)","","Perwal No. 48 Tahun 2012","Petunjuk Pelaksanaan Perda 14/2012","Perwal No. 52 Tahun 2013 (Perubahan)","",
],[
  "$MULAI BERLAKU (Pasal 37)","","Sejak diundangkan","20 Desember 2024","","Pj. WALI KOTA BEKASI,","ttd.","R. GANI MUHAMAD",
],{rc:C.navyM});

// ─── CLOSING ───
(()=>{
  const s=pres.addSlide();
  s.background={color:C.navy};
  s.addShape(pres.shapes.OVAL,{x:-1,y:-1.5,w:4.5,h:4.5,fill:{color:"0D1F3C",transparency:55}});
  s.addShape(pres.shapes.OVAL,{x:8.5,y:-2,w:7,h:7,fill:{color:"0D1F3C",transparency:55}});
  s.addShape(pres.shapes.OVAL,{x:10,y:4.5,w:5,h:5,fill:{color:"0D1F3C",transparency:55}});
  s.addShape(pres.shapes.RECTANGLE,{x:M,y:3.6,w:3.5,h:0.04,fill:{color:C.gold}});
  s.addText("BERITA DAERAH KOTA BEKASI",{x:M,y:1.6,w:CW,h:0.4,fontSize:14,color:C.gold,bold:true,fontFace:"Calibri"});
  s.addText("TERIMA KASIH",{x:M,y:2.4,w:CW,h:1.5,fontSize:48,bold:true,color:C.white,fontFace:"Calibri"});
  s.addText([{text:"Peraturan Wali Kota Bekasi Nomor 51 Tahun 2024",options:{breakLine:true,fontSize:14,color:C.ice}},{text:"Tentang Pengelolaan Pajak Reklame",options:{fontSize:14,color:C.ice}}],{x:M,y:4.1,w:CW,h:0.8,fontFace:"Calibri"});
  s.addText("Sumber: https://jdih.bekasikota.go.id",{x:M,y:5.1,w:CW,h:0.35,fontSize:10,color:C.txMu,fontFace:"Calibri"});
  goldBot(s);
})();

// ══════════════════════════════════════════════
// SAVE
// ══════════════════════════════════════════════
const path = require("path");
const out = path.resolve(__dirname, "Perwal_Bekasi_51_2024_Pajak_Reklame.pptx");
pres.writeFile({ fileName: out })
  .then(() => console.log(`✅ OK: ${out} (${pres.slides.length} slide)`))
  .catch(e => console.error("❌", e));
