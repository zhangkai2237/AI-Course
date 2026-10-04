import PptxGenJS from "pptxgenjs";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { slides, TOTAL_MINUTES } from "./content.mjs";
import { C, FONTS, THEME } from "./theme.mjs";
import { SOURCE_POLICY, assertSourcePath } from "./source-policy.mjs";
import { P, shape, text, badge, chrome, title, card, cards, quote, checklist, columns, timeline, matrix, section, line } from "./components.mjs";
import { conversationWireframe, orderConfirmationWireframe, handoffWireframe } from "./wireframes.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const SOURCE = path.resolve(ROOT, SOURCE_POLICY.soleContentSource);
const OUT = path.join(ROOT, "dist", "专题课_便宜的答案昂贵的判断_老师版_v2.0.pptx");
const MANIFEST = path.join(ROOT, "qa", "build-manifest.json");
assertSourcePath(SOURCE);
await fs.access(SOURCE);
await fs.mkdir(path.dirname(OUT), { recursive: true });
await fs.mkdir(path.dirname(MANIFEST), { recursive: true });

const pptx = new PptxGenJS();
pptx.layout = THEME.layout;
pptx.author = "AI-Course";
pptx.company = "AI PRODUCT MANAGER";
pptx.subject = "专题课：便宜的答案，昂贵的判断（老师版）";
pptx.title = "便宜的答案，昂贵的判断";
pptx.lang = "zh-CN";
pptx.theme = { headFontFace: FONTS.head, bodyFontFace: FONTS.body, lang: "zh-CN" };
pptx.defineSlideMaster({ title: "WORKBENCH", background: { color: C.paper }, objects: [] });
pptx.defineSlideMaster({ title: "COVER", background: { color: C.paper }, objects: [] });

function renderStandard(slide, d) {
  title(slide, pptx, d);
  switch (d.type) {
    case "compare": columns(slide, pptx, d.left, d.right, { y: 212, h: 355 }); break;
    case "quote":
      quote(slide, pptx, d.quote, { y: 218, h: 120, tone: d.tone, label: d.label });
      checklist(slide, pptx, d.bullets || [], { x: 135, y: 395, w: 1000, gap: 58, tone: d.tone === "yellow" ? "yellow" : "green" });
      break;
    case "threeColumns": cards(slide, pptx, d.columns, { cols: 3, y: 235, h: 245 }); quote(slide, pptx, d.key, { y: 525, h: 78, label: "使用顺序" }); break;
    case "audit": matrix(slide, pptx, ["表述", "需要的证据", "安全改写"], d.rows, { widths: [330, 400, 410], y: 210, rowH: 72, fontSize: 11 }); break;
    case "ladder":
      d.levels.forEach((it, i) => card(slide, pptx, { x: 105 + i * 275, y: 245 + i * 25, w: 250, h: 185, ...it, index: String(i + 1).padStart(2, "0") }));
      quote(slide, pptx, d.key, { y: 525, h: 76, label: "口径原则" });
      break;
    case "cards": cards(slide, pptx, d.items, { cols: 3, y: 210, h: 170 }); break;
    case "split": columns(slide, pptx, d.left, d.right, { y: 210, h: 365 }); break;
    case "loop":
      timeline(slide, pptx, d.steps, { y: 235 });
      quote(slide, pptx, d.badcase, { y: 420, h: 82, tone: "red", label: "课堂模拟 Bad Case" });
      text(slide, d.key, P(100, 548, 1080, 35), { size: 14, color: C.blue, bold: true, align: "center" });
      break;
    case "stairs":
      d.levels.forEach((it, i) => {
        const w = 820 - i * 150, x = 230 + i * 75, y = 225 + i * 102;
        card(slide, pptx, { x, y, w, h: 84, ...it, index: String(i + 1).padStart(2, "0") });
      });
      text(slide, d.key, P(200, 552, 880, 34), { size: 14, color: C.red, bold: true, align: "center" });
      break;
    case "quiz":
      d.statements.forEach((v, i) => {
        shape(slide, pptx, "roundRect", P(95, 210 + i * 92, 925, 70), C.white, { color: C.line, width: 1 }, { radius: 0.06 });
        text(slide, `${i + 1}`, P(112, 228 + i * 92, 35, 28), { size: 11, color: C.blue, bold: true, align: "center" });
        text(slide, v, P(165, 220 + i * 92, 820, 45), { size: 12.5, bold: true });
        badge(slide, pptx, d.answers[i], 1045, 229 + i * 92, [C.green, C.blue, "9A6A00", C.red][i], [C.greenSoft, C.blueSoft, C.yellowSoft, C.redSoft][i], 145);
      });
      break;
    case "scorecard": matrix(slide, pptx, ["支点", "审查对象", "证据门槛"], d.rows, { widths: [240, 520, 380], y: 220, rowH: 72, fontSize: 12 }); quote(slide, pptx, d.key, { y: 555, h: 60, label: "自评公式" }); break;
    case "overview8":
      cards(slide, pptx, d.items.map((v, i) => ({ label: v, body: i < 4 ? "先建立项目主线" : "再展示判断与证据", tone: i === 7 ? "red" : i >= 5 ? "yellow" : "blue" })), { cols: 4, y: 205, h: 155, gap: 16 });
      text(slide, d.key, P(110, 565, 1060, 32), { size: 13, color: C.red, bold: true, align: "center" });
      break;
    case "projectOverview": matrix(slide, pptx, ["项目概览字段", "安全口径"], d.fields, { widths: [250, 890], y: 205, rowH: 66, fontSize: 12 }); break;
    case "problemGoal":
      timeline(slide, pptx, d.problems, { y: 250 });
      quote(slide, pptx, d.goal, { y: 445, h: 95, tone: "green", label: "产品目标" });
      break;
    case "mvpBoundary":
      quote(slide, pptx, d.positioning, { y: 195, h: 72, label: "产品定位" });
      columns(slide, pptx, { label: "第一版做", items: d.inScope }, { label: "暂不做", items: d.outScope }, { y: 295, h: 280 });
      text(slide, d.key, P(150, 602, 980, 30), { size: 14, color: C.red, bold: true, align: "center" });
      break;
    case "wireConversation": conversationWireframe(slide, pptx); break;
    case "wireOrder": orderConfirmationWireframe(slide, pptx); break;
    case "wireHandoff": handoffWireframe(slide, pptx); break;
    case "matrix": matrix(slide, pptx, d.headers, d.rows, { widths: d.widths, y: 195, rowH: 58, fontSize: 9.8 }); break;
    case "decision":
      quote(slide, pptx, d.constraint, { y: 210, h: 85, tone: "red", label: "约束" });
      timeline(slide, pptx, d.options, { y: 335 });
      quote(slide, pptx, d.choice, { y: 505, h: 82, tone: "green", label: "选择" });
      break;
    case "twoDecisions": cards(slide, pptx, d.decisions, { cols: 2, y: 225, h: 280, gap: 30 }); quote(slide, pptx, d.key, { y: 545, h: 66, label: "判断" }); break;
    case "badcase":
      cards(slide, pptx, [
        { label: "现象", body: d.before, tone: "red" }, { label: "原因", body: d.diagnosis, tone: "yellow" },
        { label: "修改", body: d.fix, tone: "blue" }, { label: "复测", body: d.retest, tone: "green" },
      ], { cols: 4, y: 230, h: 270, gap: 18 });
      quote(slide, pptx, d.key, { y: 545, h: 66, label: "闭环" });
      break;
    case "resultBoundary": cards(slide, pptx, d.columns, { cols: 3, y: 215, h: 260 }); text(slide, d.metrics, P(110, 515, 1060, 45), { size: 12, color: C.red, bold: true, align: "center" }); text(slide, d.key, P(110, 568, 1060, 32), { size: 13, color: C.blue, bold: true, align: "center" }); break;
    case "eightSentences":
      cards(slide, pptx, d.prompts.map((v, i) => ({ label: `${i + 1}`, body: v, tone: i === 7 ? "red" : "blue" })), { cols: 4, y: 205, h: 160, gap: 16 });
      break;
    case "fourProblems": cards(slide, pptx, d.items, { cols: 4, y: 225, h: 300, gap: 18 }); quote(slide, pptx, d.key, { y: 565, h: 55, label: "目标" }); break;
    case "timedStructure": matrix(slide, pptx, ["时间", "段落", "只回答一个问题"], d.segments, { widths: [180, 250, 710], y: 205, rowH: 68, fontSize: 12 }); break;
    case "template":
      d.template.forEach((v, i) => {
        shape(slide, pptx, "roundRect", P(100, 220 + i * 92, 1080, 70), C.white, { color: C.blue, width: 1.1 }, { radius: 0.06 });
        text(slide, v, P(128, 231 + i * 92, 1024, 48), { size: 14, bold: true });
      });
      checklist(slide, pptx, d.evidence, { x: 135, y: 220 + d.template.length * 92 + 35, w: 1000, gap: 57, tone: "yellow" });
      break;
    case "script90":
      d.paragraphs.forEach((v, i) => {
        shape(slide, pptx, "roundRect", P(78 + (i % 2) * 570, 205 + Math.floor(i / 2) * 193, 548, 170), C.white, { color: [C.blue, C.green, "9A6A00", C.red][i], width: 1 }, { radius: 0.06 });
        badge(slide, pptx, ["问题", "形态与流程", "判断与Bad Case", "进展与限制"][i], 96 + (i % 2) * 570, 220 + Math.floor(i / 2) * 193, [C.blue, C.green, "9A6A00", C.red][i], [C.blueSoft, C.greenSoft, C.yellowSoft, C.redSoft][i], 170);
        text(slide, v, P(98 + (i % 2) * 570, 258 + Math.floor(i / 2) * 193, 508, 100), { size: 9.7, valign: "top", breakLine: true });
      });
      break;
    case "keepDelete":
      columns(slide, pptx, { label: "保留", items: d.keep }, { label: "删除", items: d.remove }, { y: 205, h: 345 });
      text(slide, `点评：${d.criteria.join(" · ")}`, P(150, 580, 980, 28), { size: 12, color: C.blue, bold: true, align: "center" });
      break;
    case "fiveLayers": matrix(slide, pptx, ["层次", "面试官在核对什么"], d.layers, { widths: [260, 880], y: 210, rowH: 68, fontSize: 12 }); break;
    case "qaPair":
      d.qa.forEach((item, i) => {
        const y = 220 + i * 180;
        shape(slide, pptx, "roundRect", P(90, y, 1100, 150), C.white, { color: i ? C.green : C.blue, width: 1.2 }, { radius: 0.07 });
        text(slide, item.q, P(118, y + 18, 1038, 40), { size: 16, color: i ? C.green : C.blue, bold: true });
        text(slide, item.a, P(118, y + 67, 1038, 61), { size: 12.5, valign: "top", breakLine: true });
      });
      quote(slide, pptx, d.key, { y: 575, h: 48, label: "边界" });
      break;
    case "whySteps": matrix(slide, pptx, ["步骤", "名称", "回答模板"], d.steps, { widths: [140, 210, 790], y: 205, rowH: 69, fontSize: 11.5 }); break;
    case "closing":
      columns(slide, pptx, { label: "普通回答", items: [d.weak] }, { label: "专业回答", items: [d.strong] }, { y: 200, h: 210 });
      checklist(slide, pptx, d.practice, { x: 120, y: 445, w: 1040, gap: 39, tone: "green" });
      text(slide, d.key, P(100, 638, 1080, 32), { size: 17, color: C.blue, bold: true, align: "center" });
      break;
    case "finalPractice":
      cards(slide, pptx, d.tasks.map((body, i) => ({ label: String(i + 1).padStart(2, "0"), body, tone: i === 4 ? "red" : "blue" })), { cols: 5, y: 205, h: 170, gap: 14 });
      checklist(slide, pptx, d.integrity, { x: 120, y: 420, w: 1040, gap: 48, tone: "green" });
      text(slide, d.key, P(100, 630, 1080, 32), { size: 17, color: C.blue, bold: true, align: "center" });
      break;
    default: throw new Error(`Unknown slide type: ${d.type}`);
  }
}

function renderCover(slide, d) {
  shape(slide, pptx, "rect", P(0, 0, 28, 720), C.blue);
  badge(slide, pptx, "老师版 v2.0 · 120 MIN", 76, 70, C.blue, C.blueSoft, 240);
  text(slide, d.title, P(76, 205, 1080, 105), { size: 39, bold: true });
  text(slide, d.kicker, P(80, 332, 900, 42), { size: 17, color: C.muted });
  quote(slide, pptx, d.key, { y: 430, h: 110, label: "课程命题" });
  text(slide, "FACT · CHOICE · VALIDATION · BOUNDARY", P(82, 608, 800, 28), { size: 10, color: C.blue, bold: true });
  text(slide, "44", P(1040, 570, 130, 70), { size: 42, color: C.line, bold: true, align: "right" });
}

slides.forEach((d, i) => {
  const slide = pptx.addSlide(d.type === "cover" ? "COVER" : "WORKBENCH");
  slide.background = { color: C.paper };
  slide.addNotes(d.notes);
  if (d.type === "cover") renderCover(slide, d);
  else {
    chrome(slide, pptx, d, i, slides.length);
    if (d.type === "section") section(slide, pptx, d);
    else renderStandard(slide, d);
  }
});

await pptx.writeFile({ fileName: OUT });
await fs.writeFile(MANIFEST, JSON.stringify({
  generatedAt: new Date().toISOString(),
  output: path.relative(ROOT, OUT),
  slideCount: slides.length,
  totalMinutes: TOTAL_MINUTES,
  noteCount: slides.filter(d => d.notes).length,
  source: path.relative(ROOT, SOURCE),
  sourcePolicy: SOURCE_POLICY,
  dependency: "pptxgenjs@4.0.1",
  theme: THEME.name,
  buildInputs: ["src/build.mjs", "src/content.mjs", "src/theme.mjs", "src/components.mjs", "src/wireframes.mjs", "src/notes.mjs", "src/source-policy.mjs", path.relative(ROOT, SOURCE)],
}, null, 2), "utf8");
console.log(JSON.stringify({ output: OUT, slides: slides.length, minutes: TOTAL_MINUTES }));
