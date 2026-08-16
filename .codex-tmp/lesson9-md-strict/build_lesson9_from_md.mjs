import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const MD = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第9课 - AI产品评测与数据反馈.md";
const OUT = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第09课_正式PPT_严格按MD版.pptx";
const PREVIEW = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-md-strict/rendered";
const MONTAGE = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-md-strict/第09课_严格按MD版_总览.webp";

const C = {
  bg: "#07151F", bg2: "#0B202C", ink: "#F5F2E9", muted: "#A7B5BC",
  teal: "#36D6C2", teal2: "#168E85", orange: "#FF7657", yellow: "#F0C75E",
  line: "#26404A", dark: "#031016", green: "#6EE7B7", redBg: "#1D1112",
};

function cleanInline(s) {
  return s
    .replace(/\\([\-_.])/g, "$1")
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\\-/g, "-")
    .replace(/\\\./g, ".")
    .trim();
}

function isTransition(s) {
  const t = cleanInline(s);
  return /^(我们先讲|下面讲|我们继续使用|前面我们解决了|接下来要解决的是|现在我们已经有了|下一部分是|至此，这节课|这一部分最重要的观点是)/.test(t);
}

function parseMarkdown(md) {
  const lines = md.replace(/\r/g, "").split("\n");
  const blocks = [];
  let para = [];
  let quote = [];
  let list = [];
  let code = [];
  let inCode = false;
  const flushPara = () => {
    if (para.length) {
      const text = cleanInline(para.join(" "));
      if (text && !isTransition(text)) blocks.push({ kind: "p", text });
      para = [];
    }
  };
  const flushQuote = () => {
    if (quote.length) {
      const text = cleanInline(quote.join(" "));
      if (text) blocks.push({ kind: "quote", text });
      quote = [];
    }
  };
  const flushList = () => {
    if (list.length) { blocks.push({ kind: "list", items: list.map(cleanInline).filter(Boolean) }); list = []; }
  };
  const flushCode = () => {
    if (code.length) { blocks.push({ kind: "code", text: code.join("\n").trim() }); code = []; }
  };
  const flushAll = () => { flushPara(); flushQuote(); flushList(); if (!inCode) flushCode(); };

  for (const line of lines) {
    if (/^```/.test(line.trim())) {
      flushPara(); flushQuote(); flushList();
      if (inCode) { inCode = false; flushCode(); } else inCode = true;
      continue;
    }
    if (inCode) { code.push(line); continue; }
    const hm = line.match(/^(#{1,4})\s*(.*)$/);
    if (hm) {
      flushAll();
      const title = cleanInline(hm[2]);
      if (title) blocks.push({ kind: "heading", level: hm[1].length, text: title });
      continue;
    }
    if (/^---+\s*$/.test(line)) { flushAll(); continue; }
    const qm = line.match(/^>\s?(.*)$/);
    if (qm) { flushPara(); flushList(); quote.push(qm[1]); continue; }
    const lm = line.match(/^\s*(?:[-*]|\d+\.)\s+(.*)$/);
    if (lm) { flushPara(); flushQuote(); list.push(lm[1]); continue; }
    if (!line.trim()) { flushAll(); continue; }
    flushQuote(); flushList();
    para.push(line.trim());
  }
  flushAll();
  return blocks;
}

function normalizeHeading(title, sectionNo) {
  if (/第[七八]部分总结/.test(title) && sectionNo === 7) return "第七部分总结";
  if (/第[七八]部分总结/.test(title) && sectionNo === 8) return "第八部分总结";
  return title;
}

function splitSentences(text, max = 240) {
  if (text.length <= max) return [text];
  const parts = text.split(/(?<=[。！？；：])/).filter(Boolean);
  const out = [];
  let buf = "";
  for (const p of parts) {
    if (buf && (buf + p).length > max) { out.push(buf); buf = p; }
    else buf += p;
  }
  if (buf) out.push(buf);
  if (!out.length) for (let i = 0; i < text.length; i += max) out.push(text.slice(i, i + max));
  return out;
}

function expandBlocks(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.kind === "p" && b.text.length > 260) {
      splitSentences(b.text, 235).forEach(text => out.push({ kind: "p", text }));
    } else if (b.kind === "quote" && b.text.length > 220) {
      splitSentences(b.text, 200).forEach(text => out.push({ kind: "quote", text }));
    } else if (b.kind === "list" && b.items.length > 7) {
      for (let i = 0; i < b.items.length; i += 7) out.push({ kind: "list", items: b.items.slice(i, i + 7) });
    } else out.push(b);
  }
  return out;
}

function estimate(b) {
  if (b.kind === "p") return 16 + Math.ceil(b.text.length / 58) * 24;
  if (b.kind === "quote") return 34 + Math.ceil(b.text.length / 52) * 27;
  if (b.kind === "code") return 28 + b.text.split("\n").reduce((n, x) => n + Math.max(1, Math.ceil(x.length / 62)), 0) * 22;
  if (b.kind === "list") return 10 + b.items.reduce((n, x) => n + Math.max(1, Math.ceil(x.length / 52)) * 25 + 5, 0);
  if (b.kind === "subhead") return 44;
  return 50;
}

function makePlan(blocks) {
  const plan = [{ type: "cover", title: "AI产品评测与数据反馈", subtitle: "第09课｜严格按课程讲稿内容与示例制作" }];
  let sectionNo = 0;
  let h2 = "";
  let h3 = "";
  let topic = "课程导入";
  let pending = [];

  const flush = () => {
    const expanded = expandBlocks(pending);
    let page = [];
    let used = 0;
    let continuation = 0;
    for (const b of expanded) {
      const h = estimate(b);
      if (page.length && used + h > 488) {
        plan.push({ type: "content", title: topic, continuation, blocks: page, sectionNo });
        continuation += 1; page = []; used = 0;
      }
      page.push(b); used += h;
    }
    if (page.length) plan.push({ type: "content", title: topic, continuation, blocks: page, sectionNo });
    pending = [];
  };

  for (const b of blocks) {
    if (b.kind !== "heading") { pending.push(b); continue; }
    if (b.level === 4) {
      pending.push({ kind: "subhead", text: normalizeHeading(b.text, sectionNo) });
      continue;
    }
    flush();
    if (b.level === 1) {
      if (/^第9课/.test(b.text)) continue;
      const m = b.text.match(/^第([一二三四五六七八])部分/);
      if (!m) continue;
      sectionNo = "一二三四五六七八".indexOf(m[1]) + 1;
      h2 = ""; h3 = ""; topic = b.text;
      plan.push({ type: "section", no: sectionNo, title: b.text.replace(/^第[一二三四五六七八]部分[、：]?\s*/, "") });
    } else if (b.level === 2) {
      h2 = normalizeHeading(b.text, sectionNo); h3 = ""; topic = h2;
    } else if (b.level === 3) {
      h3 = normalizeHeading(b.text, sectionNo);
      topic = h2 && !/总结/.test(h2) ? `${h2.replace(/^[一二三四五六七八九十]+、/, "")}｜${h3}` : h3;
    }
  }
  flush();
  return plan;
}

function addShape(slide, geometry, pos, fill = "none", lineFill = "none", lineWidth = 0, radius = 0) {
  const cfg = { geometry, position: pos, fill, line: { style: "solid", fill: lineFill, width: lineWidth } };
  if (radius) cfg.borderRadius = radius;
  return slide.shapes.add(cfg);
}

function addText(slide, text, pos, size = 22, color = C.ink, bold = false, align = "left", family = "PingFang SC") {
  const s = addShape(slide, "textbox", pos);
  s.text = text;
  s.text.style = { fontSize: size, color, bold, alignment: align, fontFamily: family };
  return s;
}

function sectionLabel(no) {
  return ["WHY EVAL", "QUALITY", "EVAL SET", "METHODS", "AUTOMATION", "BAD CASE", "REGRESSION", "DATA LOOP"][Math.max(0, no - 1)] || "EVAL";
}

function addChrome(slide, index, total, no) {
  addText(slide, `AI PRODUCT · ${sectionLabel(no)}`, { left: 68, top: 28, width: 340, height: 22 }, 13, C.teal, true);
  addText(slide, String(index + 1).padStart(3, "0"), { left: 1150, top: 28, width: 64, height: 22 }, 13, C.muted, true, "right");
  addShape(slide, "rect", { left: 68, top: 680, width: 1144, height: 2 }, C.line);
  addShape(slide, "rect", { left: 68, top: 680, width: Math.max(16, 1144 * ((index + 1) / total)), height: 2 }, C.teal);
}

function noteText(d) {
  return `[Sources]\n- ${MD}\n- 对应章节：${d.title || "课程封面"}`;
}

function renderBlock(slide, b, y) {
  if (b.kind === "p") {
    const h = estimate(b) - 8;
    addText(slide, b.text, { left: 88, top: y, width: 1100, height: h }, 18, C.ink, false);
    return y + h + 8;
  }
  if (b.kind === "quote") {
    const h = estimate(b);
    addShape(slide, "roundRect", { left: 86, top: y, width: 1110, height: h }, C.bg2, C.line, 1, 12);
    addShape(slide, "rect", { left: 86, top: y, width: 7, height: h }, C.orange);
    addText(slide, b.text, { left: 116, top: y + 15, width: 1048, height: h - 24 }, 20, C.ink, true);
    return y + h + 10;
  }
  if (b.kind === "list") {
    let yy = y + 4;
    for (const item of b.items) {
      const lines = Math.max(1, Math.ceil(item.length / 52));
      const h = lines * 25 + 4;
      addText(slide, "—", { left: 92, top: yy, width: 42, height: 28 }, 19, C.orange, true);
      addText(slide, item, { left: 144, top: yy - 1, width: 1034, height: h }, 18, C.ink, false);
      yy += h + 3;
    }
    return yy + 6;
  }
  if (b.kind === "code") {
    const h = estimate(b);
    addShape(slide, "roundRect", { left: 86, top: y, width: 1110, height: h }, "#031016", C.line, 1, 10);
    addText(slide, b.text, { left: 112, top: y + 14, width: 1055, height: h - 22 }, 16, C.green, false, "left", "Menlo");
    return y + h + 10;
  }
  if (b.kind === "subhead") {
    addText(slide, b.text, { left: 88, top: y + 1, width: 1080, height: 32 }, 22, C.teal, true);
    addShape(slide, "rect", { left: 88, top: y + 35, width: 80, height: 3 }, C.orange);
    return y + 45;
  }
  return y;
}

function renderSlide(p, d, index, total) {
  const slide = p.slides.add();
  slide.background.fill = C.bg;
  slide.speakerNotes.textFrame.setText(noteText(d));
  slide.speakerNotes.setVisible(true);

  if (d.type === "cover") {
    addShape(slide, "rect", { left: 0, top: 0, width: 24, height: 720 }, C.teal);
    addText(slide, "AI PRODUCT MANAGER", { left: 82, top: 70, width: 360, height: 30 }, 16, C.teal, true);
    addText(slide, d.title, { left: 80, top: 210, width: 1050, height: 100 }, 58, C.ink, true);
    addText(slide, d.subtitle, { left: 82, top: 356, width: 880, height: 45 }, 24, C.muted);
    addShape(slide, "rect", { left: 82, top: 474, width: 270, height: 6 }, C.orange);
    addText(slide, "EVAL · BAD CASE · REGRESSION · DATA LOOP", { left: 82, top: 506, width: 700, height: 26 }, 15, C.muted, true);
    addText(slide, "09", { left: 1000, top: 470, width: 160, height: 110 }, 82, C.line, true, "right");
    return;
  }

  if (d.type === "section") {
    slide.background.fill = C.bg2;
    addText(slide, String(d.no).padStart(2, "0"), { left: 72, top: 92, width: 260, height: 170 }, 118, C.teal, true);
    addShape(slide, "rect", { left: 74, top: 300, width: 100, height: 5 }, C.orange);
    const sz = d.title.length > 22 ? 40 : 48;
    addText(slide, d.title, { left: 72, top: 350, width: 1100, height: 120 }, sz, C.ink, true);
    addText(slide, "按照讲稿结构展开概念、方法与完整示例", { left: 74, top: 500, width: 960, height: 42 }, 22, C.muted);
    addChrome(slide, index, total, d.no);
    return;
  }

  addChrome(slide, index, total, d.sectionNo);
  const title = d.continuation ? `${d.title}（续）` : d.title;
  const titleSize = title.length > 30 ? 32 : title.length > 24 ? 35 : 38;
  addText(slide, title, { left: 68, top: 78, width: 1125, height: 66 }, titleSize, C.ink, true);
  addShape(slide, "rect", { left: 68, top: 158, width: 160, height: 4 }, C.teal);
  let y = 180;
  for (const b of d.blocks) y = renderBlock(slide, b, y);
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(PREVIEW, { recursive: true });
  const markdown = await fs.readFile(MD, "utf8");
  const plan = makePlan(parseMarkdown(markdown));
  await fs.writeFile("/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-md-strict/slide-plan.txt", plan.map((d, i) => `${i + 1}\t${d.type}\t${d.title || d.subtitle}`).join("\n"));
  await fs.writeFile("/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-md-strict/source-notes.txt", `Primary source:\n${MD}\n\nThe deck follows the Markdown section order. Obvious section-summary numbering typos are normalized only in slide titles.`);

  const p = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  plan.forEach((d, i) => renderSlide(p, d, i, plan.length));

  for (const [i, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(i + 1).padStart(3, "0")}`;
    await writeBlob(`${PREVIEW}/${stem}.png`, await p.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${PREVIEW}/${stem}.layout.json`, await layout.text());
  }
  await writeBlob(MONTAGE, await p.export({ format: "webp", montage: true, scale: 1 }));
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(OUT);
  console.log(JSON.stringify({ slides: plan.length, output: OUT, preview: PREVIEW }));
}

main().catch(err => { console.error(err); process.exitCode = 1; });
