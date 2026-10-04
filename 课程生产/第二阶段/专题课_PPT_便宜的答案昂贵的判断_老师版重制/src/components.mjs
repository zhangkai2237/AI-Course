import { C, FONTS } from "./theme.mjs";

export const px = (n) => n / 96;
export const P = (x, y, w, h) => ({ x: px(x), y: px(y), w: px(w), h: px(h) });
const transparent = { color: C.paper, transparency: 100 };

export function shape(slide, pptx, type, box, fill = C.white, line = transparent, extra = {}) {
  return slide.addShape(pptx.ShapeType[type], {
    ...box,
    fill: typeof fill === "string" ? { color: fill } : fill,
    line: typeof line === "string" ? { color: line, width: 1 } : line,
    ...extra,
  });
}

export function text(slide, value, box, opts = {}) {
  return slide.addText(value, {
    ...box,
    fontFace: opts.fontFace || FONTS.body,
    fontSize: opts.size || 18,
    color: opts.color || C.ink,
    bold: opts.bold || false,
    align: opts.align || "left",
    valign: opts.valign || "mid",
    margin: opts.margin ?? 0,
    breakLine: false,
    fit: "shrink",
    paraSpaceAfterPt: 0,
    lineSpacingMultiple: 1,
    isTextBox: true,
    ...opts,
  });
}

export function line(slide, pptx, x1, y1, x2, y2, color = C.line, width = 1.2, arrow = false) {
  slide.addShape(pptx.ShapeType.line, {
    x: px(x1), y: px(y1), w: px(x2 - x1), h: px(y2 - y1),
    line: { color, width, endArrowType: arrow ? "triangle" : undefined },
  });
}

export function badge(slide, pptx, label, x, y, color = C.blue, bg = C.blueSoft, width = 150) {
  shape(slide, pptx, "roundRect", P(x, y, width, 28), bg, { color, width: 0.8 }, { radius: 0.08 });
  text(slide, label, P(x + 10, y + 4, width - 20, 20), { size: 10, color, bold: true, align: "center" });
}

export function chrome(slide, pptx, d, index, total) {
  text(slide, `专题课 · 老师版 v2.0 · ${d.act}`, P(64, 26, 620, 24), { size: 9, color: C.blue, bold: true });
  text(slide, `P${String(index + 1).padStart(2, "0")}`, P(1160, 26, 56, 24), { size: 10, color: C.muted, bold: true, align: "right" });
  shape(slide, pptx, "rect", P(64, 682, 1152, 3), C.line);
  shape(slide, pptx, "rect", P(64, 682, 1152 * (index + 1) / total, 3), C.blue);
}

export function title(slide, pptx, d) {
  text(slide, d.title, P(64, 73, 1110, 55), { size: 27, bold: true });
  if (d.kicker) text(slide, d.kicker, P(66, 130, 1100, 30), { size: 11, color: C.muted });
  shape(slide, pptx, "rect", P(64, 164, 124, 4), C.blue);
  if (d.publicCase) badge(slide, pptx, "公共模拟案例", 1012, 128, C.red, C.redSoft, 164);
}

export function card(slide, pptx, { x, y, w, h, label, body, tone = "blue", index = "" }) {
  const tones = {
    blue: [C.blue, C.blueSoft], green: [C.green, C.greenSoft], yellow: ["9A6A00", C.yellowSoft], red: [C.red, C.redSoft], neutral: [C.ink, C.paper2],
  };
  const [accent, soft] = tones[tone] || tones.blue;
  shape(slide, pptx, "roundRect", P(x, y, w, h), C.white, { color: C.line, width: 1 }, { radius: 0.08, shadow: { type: "outer", color: C.shadow, opacity: 0.12, blur: 1, angle: 45, distance: 1 } });
  shape(slide, pptx, "rect", P(x, y, 7, h), accent);
  if (index) text(slide, index, P(x + 22, y + 14, 50, 22), { size: 9, color: accent, bold: true });
  text(slide, label, P(x + 22, y + (index ? 40 : 17), w - 44, 34), { size: 15, color: accent, bold: true });
  text(slide, body, P(x + 22, y + (index ? 78 : 60), w - 44, h - (index ? 92 : 76)), { size: 11.5, valign: "top", breakLine: true });
  if (soft) shape(slide, pptx, "rect", P(x + w - 8, y, 8, h), soft);
}

export function cards(slide, pptx, items, { cols = 3, y = 205, h = 160, gap = 20, x = 70, totalW = 1140 } = {}) {
  const w = (totalW - gap * (cols - 1)) / cols;
  items.forEach((it, i) => {
    const row = Math.floor(i / cols), col = i % cols;
    card(slide, pptx, { x: x + col * (w + gap), y: y + row * (h + gap), w, h, ...it, index: it.index || String(i + 1).padStart(2, "0") });
  });
}

export function quote(slide, pptx, value, { y = 250, h = 115, tone = "blue", label = "评审结论" } = {}) {
  const accent = tone === "red" ? C.red : tone === "green" ? C.green : tone === "yellow" ? "9A6A00" : C.blue;
  const soft = tone === "red" ? C.redSoft : tone === "green" ? C.greenSoft : tone === "yellow" ? C.yellowSoft : C.blueSoft;
  shape(slide, pptx, "roundRect", P(96, y, 1088, h), soft, { color: accent, width: 1.2 }, { radius: 0.08 });
  text(slide, label, P(126, y + 18, 170, 22), { size: 9.5, color: accent, bold: true });
  text(slide, value, P(126, y + 43, 1028, h - 60), { size: 20, bold: true, color: C.ink, align: "center" });
}

export function checklist(slide, pptx, items, { x = 100, y = 215, w = 1080, gap = 58, tone = "green" } = {}) {
  const accent = tone === "red" ? C.red : tone === "yellow" ? "9A6A00" : C.green;
  const soft = tone === "red" ? C.redSoft : tone === "yellow" ? C.yellowSoft : C.greenSoft;
  items.forEach((item, i) => {
    shape(slide, pptx, "roundRect", P(x, y + i * gap, 36, 36), soft, { color: accent, width: 1 }, { radius: 0.08 });
    text(slide, tone === "red" ? "!" : "✓", P(x, y + i * gap + 2, 36, 30), { size: 14, color: accent, bold: true, align: "center" });
    text(slide, item, P(x + 56, y + i * gap, w - 56, 38), { size: 14, bold: i === 0 });
  });
}

export function columns(slide, pptx, left, right, { y = 210, h = 375 } = {}) {
  [left, right].forEach((d, i) => {
    const x = 70 + i * 580;
    shape(slide, pptx, "roundRect", P(x, y, 550, h), C.white, { color: i ? C.green : C.red, width: 1.2 }, { radius: 0.08 });
    badge(slide, pptx, d.label, x + 24, y + 20, i ? C.green : C.red, i ? C.greenSoft : C.redSoft, 175);
    d.items.forEach((v, j) => {
      text(slide, "—", P(x + 28, y + 78 + j * 62, 28, 34), { size: 15, color: i ? C.green : C.red, bold: true });
      text(slide, v, P(x + 65, y + 75 + j * 62, 450, 46), { size: 12.3, breakLine: true });
    });
  });
}

export function timeline(slide, pptx, items, { y = 300 } = {}) {
  const x = 82, totalW = 1116, gap = 18, w = (totalW - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    if (i < items.length - 1) line(slide, pptx, x + i * (w + gap) + w, y + 55, x + (i + 1) * (w + gap) - 4, y + 55, C.blue, 1.7, true);
    shape(slide, pptx, "roundRect", P(x + i * (w + gap), y, w, 112), C.white, { color: i === items.length - 1 ? C.green : C.line, width: 1 }, { radius: 0.08 });
    text(slide, String(i + 1).padStart(2, "0"), P(x + i * (w + gap) + 14, y + 12, 36, 20), { size: 9, color: C.blue, bold: true });
    text(slide, it, P(x + i * (w + gap) + 12, y + 38, w - 24, 58), { size: 11.5, bold: true, align: "center", breakLine: true });
  });
}

export function matrix(slide, pptx, headers, rows, { x = 70, y = 205, widths, rowH = 58, fontSize = 10.5 } = {}) {
  const hh = 46;
  let xx = x;
  headers.forEach((h, i) => {
    shape(slide, pptx, "rect", P(xx, y, widths[i], hh), C.ink, { color: C.paper, width: 1 });
    text(slide, h, P(xx + 8, y + 8, widths[i] - 16, 30), { size: 11, color: C.white, bold: true, align: "center" });
    xx += widths[i];
  });
  rows.forEach((r, ri) => {
    xx = x;
    r.forEach((v, ci) => {
      shape(slide, pptx, "rect", P(xx, y + hh + ri * rowH, widths[ci], rowH), ri % 2 ? C.paper : C.white, { color: C.line, width: 0.8 });
      text(slide, v, P(xx + 10, y + hh + ri * rowH + 7, widths[ci] - 20, rowH - 14), { size: fontSize, bold: ci === 0, color: ci === 0 ? C.blue : C.ink, align: ci === 0 ? "center" : "left", breakLine: true });
      xx += widths[ci];
    });
  });
}

export function section(slide, pptx, d) {
  badge(slide, pptx, d.act, 72, 110, C.blue, C.blueSoft, 170);
  text(slide, d.title, P(72, 205, 1060, 112), { size: 35, bold: true });
  shape(slide, pptx, "rect", P(72, 342, 180, 6), C.blue);
  text(slide, d.kicker || "", P(76, 378, 1000, 54), { size: 17, color: C.muted });
  if (d.prompt) quote(slide, pptx, d.prompt, { y: 485, h: 96, label: "本幕问题" });
}
