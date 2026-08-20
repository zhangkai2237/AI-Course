import fs from "node:fs/promises";
import PptxGenJS from "pptxgenjs";

const ROOT = "/Users/keivn/Project/AI-Course";
const LESSON = `${ROOT}/课程生产/第一阶段/第10课_AI产品方案与PRD设计实战_重做工作区`;
const OUT = `${LESSON}/第10课_正式PPT_核心结构版.pptx`;
const OUTLINE = `${LESSON}/第10课_PPT结构与逐页内容_核心结构版.md`;

const SOURCES = [
  `${LESSON}/01_第一节_智能审核项目的业务背景与产品目标.md`,
  `${LESSON}/02_第二节_AI介入判断与智能审核方案的能力分工.md`,
  `${LESSON}/03_第三节_审核结果处理动作与人工复核流程.md`,
  `${LESSON}/04_第四节_智能审核产品的MVP范围与能力优先级.md`,
  `${LESSON}/05_第五节_大模型语义审核的输入输出与能力要求.md`,
  `${LESSON}/06_第六节_智能审核评测集的来源与验收约定.md`,
  `${LESSON}/07_第七节_完整AI产品PRD的组装.md`,
];

const C = {
  bg: "07151F", panel: "0B202C", panel2: "102B37", ink: "F5F2E9",
  muted: "A7B5BC", teal: "36D6C2", teal2: "168E85", orange: "FF7657",
  yellow: "F0C75E", line: "294550", green: "6EE7B7", red: "FF8A7A",
  dark: "031016", white: "FFFFFF",
};

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "AI-Course";
pptx.subject = "第10课 AI产品方案与PRD设计实战";
pptx.title = "AI产品方案与PRD设计实战";
pptx.company = "AI PRODUCT MANAGER";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "PingFang SC",
  bodyFontFace: "PingFang SC",
  lang: "zh-CN",
};
pptx.defineSlideMaster({
  title: "BLANK_DARK",
  background: { color: C.bg },
  objects: [],
  slideNumber: { x: 12.05, y: 0.27, w: 0.55, h: 0.22, color: C.muted, fontFace: "PingFang SC", fontSize: 9.5, bold: true, align: "right", margin: 0 },
});

const px = n => n / 96;
const P = (left, top, width, height) => ({ x: px(left), y: px(top), w: px(width), h: px(height) });
const noLine = { color: C.bg, transparency: 100 };

function addShape(slide, type, p, fill = C.panel, line = noLine, radius = 0) {
  const opts = { ...p, fill: typeof fill === "string" ? { color: fill } : fill, line: typeof line === "string" ? { color: line, width: 1 } : line };
  if (radius) opts.radius = radius;
  return slide.addShape(pptx.ShapeType[type], opts);
}

function addText(slide, value, p, size = 22, color = C.ink, bold = false, align = "left", family = "PingFang SC", extra = {}) {
  return slide.addText(value, {
    // 第9课核心结构版以 1280×720 像素字号设计；PptxGenJS 使用 pt。
    // 统一乘 1.2，使视觉字号与第9课最终版保持同一阅读尺度。
    ...p, fontFace: family, fontSize: size * 1.2, color, bold, align,
    valign: extra.valign || "mid", margin: extra.margin ?? 0,
    breakLine: false, fit: "shrink", lineSpacingMultiple: 1.0,
    paraSpaceAfterPt: 0, isTextBox: true, ...extra,
  });
}

function addLine(slide, x1, y1, x2, y2, color = C.line, width = 1.5, arrowEnd = false) {
  slide.addShape(pptx.ShapeType.line, {
    x: px(x1), y: px(y1), w: px(x2 - x1), h: px(y2 - y1),
    line: { color, width, endArrowType: arrowEnd ? "triangle" : undefined },
  });
}

function addTitle(slide, title, subtitle = "") {
  addText(slide, title, P(70, 72, 1130, 58), 25, C.ink, true);
  addShape(slide, "rect", P(70, 148, 160, 4), C.teal);
  if (subtitle) addText(slide, subtitle, P(250, 134, 930, 32), 10.5, C.muted, false, "right");
}

function addChrome(slide, i, total, phase) {
  addText(slide, `AI PRODUCT · ${phase}`, P(70, 27, 450, 22), 8, C.teal, true);
  addShape(slide, "rect", P(70, 683, 1140, 2), C.line);
  addShape(slide, "rect", P(70, 683, 1140 * (i + 1) / total, 2), C.teal);
}

function addNotes(slide, d) {
  const idx = Math.max(0, Math.min(SOURCES.length - 1, (d.sourceSection || 1) - 1));
  slide.addNotes(`[Sources]\n- ${SOURCES[idx]}\n- 内容依据：${d.title || "第10课讲义"}`);
}

function bulletList(slide, items, x = 95, y = 205, w = 1060, size = 14.5, gap = 47, color = C.ink) {
  items.forEach((it, j) => {
    addText(slide, "—", P(x, y + j * gap, 38, 30), size, C.orange, true);
    addText(slide, it, P(x + 52, y - 1 + j * gap, w - 52, 40), size, color);
  });
}

function quote(slide, value, y = 300, color = C.orange) {
  addShape(slide, "roundRect", P(86, y, 1108, 86), C.panel, { color: C.line, width: 1 });
  addShape(slide, "rect", P(86, y, 8, 86), color);
  addText(slide, value, P(122, y + 18, 1030, 50), 16, C.ink, true);
}

function cards(slide, items, { cols = 3, y = 205, h = 150 } = {}) {
  const gap = 22, x = 76, totalW = 1128, w = (totalW - gap * (cols - 1)) / cols;
  items.forEach((it, j) => {
    const row = Math.floor(j / cols), col = j % cols, yy = y + row * (h + gap), xx = x + col * (w + gap);
    addShape(slide, "roundRect", P(xx, yy, w, h), j % 2 ? C.panel : C.panel2, { color: C.line, width: 1 });
    addText(slide, it.k, P(xx + 22, yy + 18, w - 44, 34), 14.5, it.color || C.teal, true);
    addText(slide, it.v, P(xx + 22, yy + 57, w - 44, h - 69), it.size || 10.5, C.ink, false, "left", "PingFang SC", { valign: "top", breakLine: true });
  });
}

function flow(slide, items, { y = 300, small = false, start = 1 } = {}) {
  const gap = 20, x = 78, totalW = 1124, w = (totalW - gap * (items.length - 1)) / items.length, h = small ? 92 : 116;
  for (let j = 0; j < items.length - 1; j++) addLine(slide, x + j * (w + gap) + w, y + h / 2, x + (j + 1) * (w + gap) - 5, y + h / 2, C.teal, 2, true);
  items.forEach((it, j) => {
    const xx = x + j * (w + gap);
    addShape(slide, "roundRect", P(xx, y, w, h), C.panel2, { color: C.line, width: 1 });
    addText(slide, String(start + j).padStart(2, "0"), P(xx + 14, y + 10, 36, 24), 8.5, C.orange, true);
    addText(slide, it, P(xx + 14, y + (small ? 36 : 41), w - 28, small ? 45 : 56), small ? 10 : 12, C.ink, true, "center", "PingFang SC", { breakLine: true });
  });
}

function compare(slide, left, right) {
  const y = 212, w = 530, h = 380;
  [left, right].forEach((d, j) => {
    const x = 76 + j * 598;
    addShape(slide, "roundRect", P(x, y, w, h), j ? C.panel2 : C.panel, { color: C.line, width: 1 });
    addText(slide, d.h, P(x + 28, y + 20, w - 56, 42), 17, j ? C.teal : C.orange, true);
    bulletList(slide, d.items, x + 25, y + 89, w - 50, 12.5, 55);
  });
}

function matrix(slide, headers, rows, { y = 205, widths = null, rowHeight = 58, fontSize = 10.3 } = {}) {
  const x = 70, total = 1140, hh = 54;
  const ws = widths || headers.map(() => total / headers.length);
  let xx = x;
  headers.forEach((h, j) => {
    addShape(slide, "rect", P(xx, y, ws[j], hh), C.teal2, { color: C.bg, width: 1 });
    addText(slide, h, P(xx + 8, y + 10, ws[j] - 16, 34), 11, C.ink, true, "center");
    xx += ws[j];
  });
  rows.forEach((r, i) => {
    xx = x;
    r.forEach((v, j) => {
      addShape(slide, "rect", P(xx, y + hh + i * rowHeight, ws[j], rowHeight), i % 2 ? C.panel2 : C.panel, { color: C.line, width: 1 });
      addText(slide, v, P(xx + 10, y + hh + i * rowHeight + 8, ws[j] - 20, rowHeight - 16), fontSize, j === 0 ? C.teal : C.ink, j === 0, j === 0 ? "center" : "left", "PingFang SC", { breakLine: true });
      xx += ws[j];
    });
  });
}

function bars(slide, items, y = 216) {
  const max = Math.max(...items.map(d => d.v));
  items.forEach((d, i) => {
    const yy = y + i * 73;
    addText(slide, d.k, P(110, yy, 250, 34), 13, C.ink, i === 0);
    addShape(slide, "roundRect", P(380, yy + 6, 680, 18), C.panel2);
    addShape(slide, "roundRect", P(380, yy + 6, 680 * d.v / max, 18), i === 0 ? C.orange : C.teal);
    addText(slide, d.s || `${d.v}%`, P(1080, yy - 1, 90, 30), 11, C.muted, true, "right");
  });
}

function stacked(slide, items) {
  items.forEach((it, j) => {
    const x = 150 + j * 50, y = 210 + j * 120, w = 980 - j * 100;
    addShape(slide, "roundRect", P(x, y, w, 94), it.color || (j === 0 ? "31191B" : j === 1 ? C.panel2 : C.panel), { color: j === 0 ? C.red : C.line, width: 1 });
    addText(slide, it.k, P(x + 24, y + 15, 230, 36), 14.5, j === 0 ? C.red : C.teal, true);
    addText(slide, it.v, P(x + 255, y + 15, w - 285, 55), 12, C.ink, false, "left", "PingFang SC", { breakLine: true });
  });
}

function caseSlide(slide, d) {
  addTitle(slide, d.title, d.subtitle || "贯穿案例");
  addShape(slide, "roundRect", P(82, 195, 1116, 100), C.panel, { color: C.line, width: 1 });
  addText(slide, d.q, P(112, 218, 1056, 52), 15.5, C.ink, true);
  addShape(slide, "roundRect", P(82, 320, 1116, 105), C.dark, { color: C.line, width: 1 });
  addText(slide, d.a, P(112, 343, 1056, 62), 13, C.green, false, "left", "Menlo", { breakLine: true });
  cards(slide, d.points.map(([k, v]) => ({ k, v })), { cols: d.points.length, y: 455, h: 125 });
}

function splitFlow(slide, top, bottom) {
  flow(slide, top, { y: 225, small: true });
  addLine(slide, 640, 326, 640, 373, C.orange, 2, true);
  flow(slide, bottom, { y: 385, small: true, start: top.length + 1 });
}

function lifecycle(slide, items) {
  const x = 105, y = 305, w = 170, h = 92, gap = 30;
  items.forEach((it, i) => {
    if (i < items.length - 1) addLine(slide, x + i * (w + gap) + w, y + h / 2, x + (i + 1) * (w + gap) - 5, y + h / 2, i === items.length - 2 ? C.orange : C.teal, 2, true);
    addShape(slide, "roundRect", P(x + i * (w + gap), y, w, h), i === items.length - 1 ? "31191B" : C.panel2, { color: i === items.length - 1 ? C.red : C.line, width: 1 });
    addText(slide, it, P(x + i * (w + gap) + 14, y + 20, w - 28, 50), 11.5, i === items.length - 1 ? C.red : C.ink, true, "center", "PingFang SC", { breakLine: true });
  });
}

function codeBlock(slide, code, y = 205, h = 390, size = 10.5) {
  addShape(slide, "roundRect", P(130, y, 1020, h), C.dark, { color: C.line, width: 1 });
  addText(slide, code, P(170, y + 25, 940, h - 50), size, C.green, false, "left", "Menlo", { valign: "top", breakLine: true });
}

const S = [];
const add = (type, title, cfg = {}) => S.push({ type, title, ...cfg });

add("cover", "AI产品方案与PRD设计实战", { subtitle: "第10课｜核心结构版", phase: "OVERVIEW", sourceSection: 1 });
add("quote", "第9课会评测，第10课把评测要求前置到产品方案", { phase: "OVERVIEW", sourceSection: 1, quote: "怎样在开发之前，就把业务问题、AI能力、异常处理和验收要求写清楚？", items: ["不是先选模型，而是先把业务任务拆清楚。", "不是只写页面，而是定义能力、流程和责任边界。", "不是等上线再测，而是在PRD里约定评测与接管机制。"] });
add("map", "第10课的产品设计主线", { phase: "OVERVIEW", sourceSection: 7, items: ["业务问题", "AI介入", "能力分工", "结果流程", "MVP", "能力契约", "评测验收", "组装PRD"], small: true });
add("cards", "学习目标", { phase: "OVERVIEW", sourceSection: 7, cols: 3, items: [
  { k: "会判断", v: "比较人工、规则、大模型和混合方案，判断AI是否值得介入。" },
  { k: "会设计", v: "把机器判断接入修改、复核、版本和异常处理流程。" },
  { k: "会落地", v: "把MVP、能力输入输出和验收约定组装成可开发的PRD。" },
] });

add("section", "01 业务背景与产品目标", { phase: "BUSINESS", sourceSection: 1, subtitle: "先把真实业务问题说清楚，再讨论AI" });
add("splitFlow", "项目从内容生产提速开始", { phase: "BUSINESS", sourceSection: 1, top: ["AI辅助生产", "每日内容激增", "人工审核积压"], bottom: ["退回修改", "再次审核", "通过后发布"] });
add("cards", "人工审核流程出现五类问题", { phase: "BUSINESS", sourceSection: 1, cols: 3, items: [
  { k: "数量", v: "内容从几十条增长到几百条，审核工作量线性上升。" },
  { k: "一致性", v: "同一句话，不同审核人员可能给出不同判断。" },
  { k: "规则变化", v: "行业、渠道、版本和活动阶段持续变化。" },
  { k: "数据沉淀", v: "聊天式反馈无法形成风险类型、依据和修改记录。" },
  { k: "责任边界", v: "高风险内容不能完全交给机器独立放行。" },
] });
add("formula", "生产效率提升，审核成为新瓶颈", { phase: "BUSINESS", sourceSection: 1, formula: "50条 × 5分钟 = 250分钟   →   300条 × 5分钟 = 1500分钟", items: ["AI提高了内容生产效率，却没有同步提高审核效率。", "真正的问题不是“模型能不能审”，而是审核链路能不能扩展。"] });
add("compare", "审核不一致与规则变化同时存在", { phase: "BUSINESS", sourceSection: 1, left: { h: "同一句话，多种判断", items: ["“可能是”是否降低绝对化程度？", "“最适合”是否仍属于高风险？", "是否需要结合渠道和人群判断？"] }, right: { h: "同一规则，多种上下文", items: ["渠道不同，适用标准不同", "行业不同，风险底线不同", "规则更新，需要保留历史版本"] } });
add("compare", "系统既要沉淀数据，也要保留人工责任", { phase: "BUSINESS", sourceSection: 1, left: { h: "结构化沉淀", items: ["风险类型与原文位置", "命中规则和判断依据", "最终修改与人工意见"] }, right: { h: "安全接管", items: ["模型可能漏判或误拦", "新规则可能尚未进入知识库", "高风险结果需要责任人决定"] } });
add("stacked", "先补业务基线，再验证产品假设", { phase: "BUSINESS", sourceSection: 1, layers: [
  { k: "当前基线", v: "每日审核量、平均耗时、驳回率、积压量、风险分布和人工一致性。" },
  { k: "产品假设", v: "规则＋AI＋人工预审能减少重复工作，同时不突破高风险安全底线。" },
  { k: "验证方式", v: "历史样本PoC、小范围试点、真实人工耗时与复核量对比。" },
] });
add("flow", "不要把业务需求直接写成产品需求", { phase: "BUSINESS", sourceSection: 1, items: ["表面需求\n做AI审核工具", "业务问题\n审核积压与不一致", "用户问题\n看不清哪里错、怎么改", "产品机会\n增加机器预审闭环"] });
add("cards", "项目中有四类核心用户", { phase: "BUSINESS", sourceSection: 1, cols: 2, items: [
  { k: "内容运营", v: "想知道哪里有问题、为什么有问题、应该怎么改。" },
  { k: "人工审核", v: "想直接看到机器发现了什么、依据是什么、为何转人工。" },
  { k: "规则运营", v: "关心规则版本、更新影响和历史内容追溯。" },
  { k: "业务负责人", v: "关心效率是否提升，以及业务风险是否可控。" },
] });
add("matrix", "产品目标要能被业务和评测共同验证", { phase: "BUSINESS", sourceSection: 1, headers: ["目标方向", "产品目标", "后续验证"], widths: [220, 570, 350], rows: [
  ["效率", "减少人工从头检查每条内容的重复工作", "单条审核耗时、积压量"],
  ["质量", "提前发现高频、明确和语义型风险", "漏判、误判、风险覆盖"],
  ["一致性", "把规则、依据和输出结构标准化", "人工一致性、解析成功率"],
  ["安全", "不确定与高风险内容能够转人工", "人工接管率、失败放行数"],
] });
add("compare", "第一版同时写清目标与非目标", { phase: "BUSINESS", sourceSection: 1, left: { h: "第一版要做", items: ["文本营销内容机器预审", "风险定位、依据与修改建议", "高风险和不确定结果转人工"] }, right: { h: "第一版不做", items: ["不覆盖所有行业与多模态", "不自动修改并直接发布", "不替代法务承担最终责任"] } });
add("quote", "第一版项目定义", { phase: "BUSINESS", sourceSection: 1, quote: "规则处理明确检查，大模型识别复杂语义，知识库提供当前依据；低风险进入快速流程，高风险与不确定结果交给人工。", items: ["产品对象：营销文本内容。", "产品位置：正式发布前的机器预审。", "产品底线：不自动发布，不替代最终责任人。"] });

add("section", "02 AI介入判断与能力分工", { phase: "SOLUTION", sourceSection: 2, subtitle: "先比较方案，再决定每种能力负责什么" });
add("cards", "智能审核至少有四种候选方案", { phase: "SOLUTION", sourceSection: 2, cols: 2, items: [
  { k: "纯人工", v: "理解上下文、可以担责，但工作量随内容量线性增长。" },
  { k: "纯规则", v: "稳定、便宜、可解释，但难以覆盖隐含语义和上下文。" },
  { k: "纯大模型", v: "语义理解强，但结果波动、依据可能过时、责任不清。" },
  { k: "混合方案", v: "规则＋大模型＋RAG＋人工，能力互补但系统复杂度更高。" },
] });
add("matrix", "四种方案的取舍", { phase: "SOLUTION", sourceSection: 2, headers: ["方案", "语义理解", "稳定性", "当前依据", "人工量", "复杂度"], widths: [180, 190, 190, 210, 180, 190], rowHeight: 62, rows: [
  ["纯人工", "强", "受人员一致性影响", "人工查询", "高", "低"],
  ["纯规则", "弱", "高", "结构化规则", "中", "中"],
  ["纯大模型", "强", "相对较低", "容易缺失", "中", "中"],
  ["混合方案", "强", "可通过分工改善", "RAG管理", "聚焦高风险", "高"],
] });
add("matrix", "AI是否介入，需要回答六个产品问题", { phase: "SOLUTION", sourceSection: 2, headers: ["判断问题", "当前项目", "产品结论"], widths: [340, 470, 330], rowHeight: 59, fontSize: 9.8, rows: [
  ["业务量是否足够大", "内容快速增长并出现积压", "值得自动化"],
  ["规则能否完全覆盖", "隐含表达难以穷举", "需要语义能力"],
  ["是否有可用知识", "规则、制度和案例存在", "可以建设知识库"],
  ["错误能否被接管", "发布前可拦截和复核", "适合人机协作"],
  ["机器能否独立决策", "高风险不允许自动放行", "必须保留人工"],
  ["流程是否固定", "检查与分流步骤明确", "优先Workflow"],
] });
add("cards", "只使用大模型，上线时会暴露六类问题", { phase: "SOLUTION", sourceSection: 2, cols: 3, items: [
  { k: "不稳定", v: "相同输入可能给出不同等级和结论。" },
  { k: "依据过时", v: "模型不一定知道内部最新规则。" },
  { k: "难以追溯", v: "不知道当时使用哪版规则和Prompt。" },
  { k: "成本浪费", v: "格式、词表和阈值不需要大模型判断。" },
  { k: "输出难接入", v: "一段自然语言无法稳定驱动产品流程。" },
  { k: "责任越界", v: "模型不能独立承担高风险最终责任。" },
] });
add("matrix", "先按任务性质分配能力", { phase: "SOLUTION", sourceSection: 2, headers: ["任务性质", "典型任务", "优先能力"], widths: [280, 560, 300], rows: [
  ["确定、可枚举", "禁用词、必要声明、格式、明确阈值", "规则"],
  ["依赖上下文", "暗示、同义改写、组合语义、误导表达", "大模型"],
  ["依赖当前资料", "渠道规则、版本条款、历史争议结论", "RAG"],
  ["高风险或不确定", "依据冲突、信息不足、责任判断", "人工"],
] });
add("cards", "确定性规则负责稳定执行", { phase: "SOLUTION", sourceSection: 2, cols: 3, items: [
  { k: "词与表达", v: "禁用词、敏感词、高风险固定短语。" },
  { k: "声明与格式", v: "必要风险提示、字数、字段和格式。" },
  { k: "明确阈值", v: "可以清楚枚举的业务限制和数值条件。" },
] });
add("cards", "大模型负责规则难以穷举的语义", { phase: "SOLUTION", sourceSection: 2, cols: 3, items: [
  { k: "隐含承诺", v: "没有直接使用禁词，但整体表达暗示保证结果。" },
  { k: "上下文误导", v: "单句正常，组合后可能让用户产生错误理解。" },
  { k: "同义绕过", v: "用改写、否定、反讽或案例包装绕开固定词。" },
] });
add("compare", "RAG提供依据，人工承担责任", { phase: "SOLUTION", sourceSection: 2, left: { h: "RAG知识库", items: ["检索当前有效规则", "提供规则解释和正反案例", "保留业务、渠道和版本信息"] }, right: { h: "人工复核", items: ["处理高风险与信息不足", "解决规则和证据冲突", "作出需要责任人确认的最终判断"] } });
add("splitFlow", "混合审核的完整协作流程", { phase: "SOLUTION", sourceSection: 2, top: ["输入校验", "确定性规则", "检索适用依据"], bottom: ["大模型语义审核", "结果合并", "风险路由"] });
add("compare", "第一版用Workflow，而不是Agent", { phase: "SOLUTION", sourceSection: 2, left: { h: "Workflow适合当前主链路", items: ["必要步骤每次都执行", "每个节点可以单独测试", "失败位置容易定位和追踪"] }, right: { h: "Agent留给后续复杂任务", items: ["多渠道自主收集材料", "动态追问和外部调查", "受控生成完整调查报告"] } });
add("matrix", "能力分工要写到输入、输出和边界", { phase: "SOLUTION", sourceSection: 2, headers: ["模块", "主要输入", "主要输出", "不负责"], widths: [220, 300, 330, 290], rowHeight: 57, fontSize: 9.2, rows: [
  ["输入校验", "文案、渠道、产品信息", "缺失字段与格式问题", "语义风险"],
  ["规则检查", "文案、结构化规则", "命中规则、风险片段", "复杂上下文"],
  ["RAG检索", "业务类型、渠道、风险类别", "有效依据与案例", "最终判断"],
  ["大模型", "文案、元数据、规则与依据", "风险、理由、建议", "最终责任"],
  ["人工复核", "机器结果、依据、原文", "最终结论与原因", "重复检查全部低风险内容"],
] });
add("case", "同一条文案如何被分层审核", { phase: "SOLUTION", sourceSection: 2, q: "这款产品是目前最值得购买的理财产品，预计年化收益可达20%，收益稳定，完全不用担心亏损。", a: "整体风险等级：高风险｜机器建议：不允许直接发布｜下一步：转人工复核", points: [["规则", "绝对化与高风险短语"], ["RAG", "收益保证与数据规范"], ["大模型", "收益误导与风险隐瞒"], ["人工", "驳回并记录复核原因"]] });

add("section", "03 结果动作与人工复核", { phase: "FLOW", sourceSection: 3, subtitle: "发现问题以后，产品要让用户继续完成任务" });
add("matrix", "先把规则、RAG和模型的原始结果放在一起", { phase: "FLOW", sourceSection: 3, headers: ["来源", "发现内容", "产品价值"], widths: [220, 610, 310], rows: [
  ["规则", "命中“最值得购买”“不用担心亏损”", "稳定召回明确问题"],
  ["RAG", "返回绝对化、收益保证与数据来源条款", "提供当前审核依据"],
  ["大模型", "识别收益数据不足、收益误导和风险隐瞒", "补充语义风险"],
] });
add("flow", "结果合并解决三个问题", { phase: "FLOW", sourceSection: 3, items: ["相同风险\n不重复展示", "不同来源\n继续保留证据", "每个风险\n对应处理建议"] });
add("code", "一条正式风险结果必须结构化", { phase: "FLOW", sourceSection: 3, code: `风险片段：完全不用担心亏损\n风险类型：收益保证与风险隐瞒\n风险等级：高\n识别来源：规则＋大模型＋RAG\n审核依据：金融营销内容审核规则V3，第18条\n修改建议：删除保证性表达，补充必要风险提示\n是否需要人工复核：是` });
add("cards", "一篇内容需要四类整体处理结果", { phase: "FLOW", sourceSection: 3, cols: 2, items: [
  { k: "机器预审通过", v: "未发现阻断风险，依据完整，且属于当前支持范围。" },
  { k: "建议修改", v: "只有低风险或提示性问题，用户可以确认或修改。" },
  { k: "退回修改", v: "存在明确且可以通过修改解决的中风险问题。" },
  { k: "转人工复核", v: "高风险、不确定、依据不足、规则冲突或系统异常。" },
] });
add("matrix", "风险等级必须映射到明确产品动作", { phase: "FLOW", sourceSection: 3, headers: ["审核情况", "整体结果", "下一步"], widths: [460, 300, 380], rows: [
  ["未发现风险", "机器预审通过", "进入原有后续流程"],
  ["只有低风险问题", "建议修改", "用户确认、修改或申请人工"],
  ["明确中风险问题", "退回修改", "生成新版本后重新审核"],
  ["高风险、不确定或依据冲突", "待人工复核", "进入人工审核工作台"],
] });
add("splitFlow", "运营端不是只看结果，而是完成修改闭环", { phase: "FLOW", sourceSection: 3, top: ["提交内容V1", "查看风险位置", "理解依据与建议"], bottom: ["修改形成V2", "重新完整审核", "进入后续流程"] });
add("cards", "六类情况触发人工复核", { phase: "FLOW", sourceSection: 3, cols: 3, items: [
  { k: "高风险", v: "命中红线规则或高风险语义。" },
  { k: "不确定", v: "模型无法稳定作出判断。" },
  { k: "信息不足", v: "缺少产品事实或业务上下文。" },
  { k: "依据冲突", v: "规则、知识或版本之间存在冲突。" },
  { k: "用户申诉", v: "运营对机器结论发起复核。" },
  { k: "系统失败", v: "重试后规则、RAG或模型仍未完成。" },
] });
add("compare", "人工复核工作台要减少重复劳动", { phase: "FLOW", sourceSection: 3, left: { h: "人工需要看到", items: ["原始内容与内容版本", "机器风险、依据和原始输出", "转人工原因与历史处理记录"] }, right: { h: "人工可以执行", items: ["复核通过或驳回修改", "调整风险类型与等级", "升级给规则或专业负责人"] } });
add("lifecycle", "审核任务需要明确状态", { phase: "FLOW", sourceSection: 3, items: ["待审核", "机器审核中", "待修改", "待人工复核", "审核完成"] });
add("compare", "机器结果和人工结果必须同时保留", { phase: "FLOW", sourceSection: 3, left: { h: "最小追溯记录", items: ["机器原始结果", "人工最终结果", "修改原因、操作人与时间"] }, right: { h: "最小版本绑定", items: ["内容版本", "规则与知识版本", "模型与Prompt版本"] } });
add("stacked", "内容修改和机器失败都不能偷走安全步骤", { phase: "FLOW", sourceSection: 3, layers: [
  { k: "内容发生实质修改", v: "生成新版本，原审核结论不再适用，必须重新执行完整审核。" },
  { k: "机器审核未完成", v: "保留已有结果，允许有限安全重试，再次失败后转人工。" },
  { k: "统一原则", v: "没有完成审核，不等于审核没有问题。" },
] });
add("flow", "完整案例从V1走到V2", { phase: "FLOW", sourceSection: 3, small: true, items: ["提交V1", "发现4项风险", "转人工复核", "驳回修改", "生成V2", "重新审核", "预审通过"] });

add("section", "04 MVP范围与能力优先级", { phase: "MVP", sourceSection: 4, subtitle: "第一版不是功能最少，而是验证链路完整" });
add("quote", "MVP验证的是最小审核闭环", { phase: "MVP", sourceSection: 4, quote: "真实提交 → 机器预审 → 查看结构化结果 → 修改或人工复核 → 记录最终结果", items: ["功能少不等于能够验证价值。", "第一版必须同时验证能力、效率和安全接管。"] });
add("cards", "第一版要验证三个核心假设", { phase: "MVP", sourceSection: 4, cols: 3, items: [
  { k: "发现更多有效风险", v: "混合方案能识别纯规则难以覆盖的语义风险，同时不过度增加误报。" },
  { k: "减少人工重复工作", v: "机器提前定位问题、提供依据和建议，人工不再从头检查所有内容。" },
  { k: "异常可安全接管", v: "不确定、高风险和失败任务可以完整交给人工处理。" },
] });
add("matrix", "PoC、MVP与完整产品回答不同问题", { phase: "MVP", sourceSection: 4, headers: ["阶段", "核心问题", "需要具备"], widths: [220, 510, 410], rows: [
  ["能力PoC", "技术方案在历史样本上是否基本可行？", "样本、规则、模型与离线对比"],
  ["产品MVP", "进入真实流程后是否提高效率且安全？", "真实提交、人工接管、结果记录"],
  ["完整产品", "如何支持规模化、多渠道和复杂组织？", "多模态、权限、批量、发布与运营平台"],
] });
add("flow", "功能是否进入MVP，连续问四个问题", { phase: "MVP", sourceSection: 4, items: ["是否直接验证\n核心假设", "缺少它是否无法\n形成闭环", "能否暂时由人工\n或简单工具替代", "投入是否会显著\n拖慢试点"] });
add("matrix", "第一版主动收窄试点范围", { phase: "MVP", sourceSection: 4, headers: ["范围维度", "MVP选择", "暂不扩展"], widths: [230, 490, 420], rows: [
  ["内容类型", "文本营销文案", "图片、音频、视频"],
  ["业务范围", "一个业务线或高频场景", "全部行业与组织"],
  ["渠道", "少量规则相对明确的渠道", "全渠道统一覆盖"],
  ["处理方式", "预审、修改、复核和重审", "自动发布和处罚"],
] });
add("cards", "MVP必须具备八类能力", { phase: "MVP", sourceSection: 4, cols: 4, items: [
  { k: "提交", v: "文案和必要上下文" }, { k: "规则", v: "确定性风险检查" },
  { k: "RAG", v: "当前规则依据" }, { k: "大模型", v: "语义风险识别" },
  { k: "结果", v: "结构化展示与合并" }, { k: "分流", v: "通过、修改或转人工" },
  { k: "复核", v: "人工最终处理" }, { k: "追溯", v: "版本、日志和失败记录" },
] });
add("cards", "这些能力暂时不进入第一版", { phase: "MVP", sourceSection: 4, cols: 3, items: [
  { k: "多模态审核", v: "图片、视频和音频会显著扩大技术与标准范围。" },
  { k: "自动改写发布", v: "不能在未确认事实和责任的情况下直接覆盖原文。" },
  { k: "完整规则后台", v: "第一版可以由产品与研发预配置规则。" },
  { k: "复杂审批权限", v: "先用最小角色和人工分配验证主链路。" },
  { k: "Agent自主调查", v: "关键审核步骤不允许被自主跳过。" },
] });
add("stacked", "能力优先级围绕闭环排序", { phase: "MVP", sourceSection: 4, layers: [
  { k: "P0｜没有就无法试点", v: "提交、规则、语义审核、结构化结果、人工接管、重审、日志。" },
  { k: "P1｜提高使用效率", v: "更好的任务筛选、批量查看、修改辅助和基础统计。" },
  { k: "P2｜规模化扩展", v: "多渠道、多模态、完整规则平台、复杂权限和自动发布。" },
] });
add("compare", "MVP采用小范围试点，而不是一次性全量上线", { phase: "MVP", sourceSection: 4, left: { h: "试点方式", items: ["先跑历史样本PoC", "选择一个团队影子运行", "人工仍保留最终控制权"] }, right: { h: "成功标准", items: ["有效风险发现能力提升", "人工平均耗时下降", "高风险无默认放行，复核量可承受"] } });

add("section", "05 大模型能力输入与输出", { phase: "CAPABILITY", sourceSection: 5, subtitle: "把“大模型审核”写成可实现、可接入的能力契约" });
add("cards", "一项AI能力要求至少包含八个部分", { phase: "CAPABILITY", sourceSection: 5, cols: 4, items: [
  { k: "目标", v: "为什么需要这项能力" }, { k: "对象", v: "处理什么内容" },
  { k: "输入", v: "模型能够看到什么" }, { k: "任务", v: "模型具体判断什么" },
  { k: "输出", v: "必须返回哪些字段" }, { k: "约束", v: "哪些规则必须遵守" },
  { k: "异常", v: "不确定与失败怎么处理" }, { k: "边界", v: "明确不负责什么" },
] });
add("compare", "不要让模型判断“是否合法”", { phase: "CAPABILITY", sourceSection: 5, left: { h: "范围过大的写法", items: ["判断文案是否合法", "默认模型掌握全部法律事实", "责任和验收标准都不清楚"] }, right: { h: "可实现的写法", items: ["基于给定规则和事实识别营销表达风险", "标记原文位置并说明依据", "输出后续处理建议和不确定原因"] } });
add("cards", "模型需要五类输入信息", { phase: "CAPABILITY", sourceSection: 5, cols: 3, items: [
  { k: "完整内容", v: "标题、正文、按钮文案与必要上下文。" },
  { k: "业务上下文", v: "产品类型、渠道、用户、活动和内容用途。" },
  { k: "产品事实", v: "功能、适用范围、数据、价格和官方表述。" },
  { k: "审核依据", v: "当前有效规则、解释、案例和版本信息。" },
  { k: "输出约束", v: "风险类型枚举、原文定位和JSON结构要求。" },
] });
add("code", "模型输入不是一句Prompt，而是一组受控上下文", { phase: "CAPABILITY", sourceSection: 5, code: `待审内容：完整营销文案\n业务上下文：金融产品｜小红书｜成年人｜拉新活动\n产品事实：历史收益数据、统计周期、风险说明\n规则依据：当前有效规则V3＋正反案例\n输出要求：限定风险类型；必须定位原文；信息不足不得编造` });
add("code", "输出结构必须能够被系统稳定解析", { phase: "CAPABILITY", sourceSection: 5, code: `{
  "overall_result": "need_human_review",
  "overall_risk_level": "high",
  "risks": [{
    "text_span": "完全不用担心亏损",
    "risk_type": "收益保证与风险隐瞒",
    "reason": "暗示本金和收益不会发生损失",
    "evidence": "规则V3-18",
    "suggestion": "删除保证性表达并补充风险提示",
    "need_human_review": true
  }]
}`, codeSize: 8.8 });
add("matrix", "模型提供判断材料，系统决定业务动作", { phase: "CAPABILITY", sourceSection: 5, headers: ["模型结果", "系统继续检查", "后续动作"], widths: [300, 450, 390], rows: [
  ["未发现语义风险", "规则是否命中、调用是否完整", "进入整体结果合并"],
  ["低风险", "是否允许运营自行处理", "建议修改"],
  ["中风险", "规则要求与业务场景", "退回修改或人工"],
  ["高风险/信息不足", "高风险路由和缺失信息", "转人工或补充信息"],
  ["输出异常/调用失败", "能否安全重试", "重试后转人工"],
] });
add("compare", "修改建议和最终改写是两个任务", { phase: "CAPABILITY", sourceSection: 5, left: { h: "模型可以提供", items: ["应删除或弱化的表达", "需要补充的限制条件", "一条参考性示例表达"] }, right: { h: "第一版不能自动做", items: ["直接覆盖运营原文", "跳过重新审核", "在缺少事实时自行补充内容"] } });
add("cards", "不确定与失败结果也必须有产品定义", { phase: "CAPABILITY", sourceSection: 5, cols: 3, items: [
  { k: "信息不足", v: "指出缺少什么、为何影响判断，以及补充或转人工。" },
  { k: "依据冲突", v: "展示冲突依据，不让模型自行选择最终结论。" },
  { k: "定位不稳定", v: "允许返回段落级风险，并降低自动处理程度。" },
  { k: "结构错误", v: "缺字段、值越界或无法解析都视为机器失败。" },
  { k: "调用失败", v: "有限安全重试，仍失败则转人工并保留记录。" },
] });
add("matrix", "AI能力边界必须写进PRD", { phase: "CAPABILITY", sourceSection: 5, headers: ["模型不负责", "原因"], widths: [460, 680], rows: [
  ["替代法务或规则负责人作最终结论", "模型不承担最终业务与法律责任"],
  ["判断未提供证据的产品事实是否真实", "模型不能凭常识补齐内部事实"],
  ["自动放行高风险或不确定内容", "业务动作由系统规则与人工控制"],
  ["自动覆盖原文、发布、删除或处罚", "建议仍需用户确认并重新审核"],
  ["在缺少依据时编造审核规则", "必须允许返回信息不足"],
] });
add("summary", "AI能力卡：营销文案语义风险识别", { phase: "CAPABILITY", sourceSection: 5, items: [
  "目标：识别必须结合上下文判断的营销表达风险",
  "输入：完整文案、业务上下文、产品事实和当前规则",
  "输出：风险位置、类型、等级、依据、建议和不确定原因",
  "异常：信息不足补充信息；冲突、失败和高风险转人工",
] });

add("section", "06 评测集来源与验收约定", { phase: "EVALUATION", sourceSection: 6, subtitle: "PRD要提前说明拿什么测、谁确认、何时重测" });
add("cards", "首版评测集来自五类真实与补充数据", { phase: "EVALUATION", sourceSection: 6, cols: 3, items: [
  { k: "历史记录", v: "通过、驳回、修改前后版本和人工备注。" },
  { k: "高风险案例", v: "投诉、损失、红线和不能漏掉的典型问题。" },
  { k: "争议案例", v: "审核分歧、申诉、升级和规则冲突内容。" },
  { k: "近期样本", v: "最近业务、新产品、新渠道和新型规避表达。" },
  { k: "边界样本", v: "人工构造并用模型扩写的正反与临界变体。" },
] });
add("stacked", "历史标签不能直接等于标准答案", { phase: "EVALUATION", sourceSection: 6, layers: [
  { k: "历史数据的价值", v: "来自真实业务，能够反映常见表达、风险分布和人工流程。" },
  { k: "历史数据的问题", v: "人工标准不一、规则可能失效、备注不完整、旧错误可能被保留。" },
  { k: "进入评测集前", v: "历史样本抽样复核；高风险和争议样本逐条确认。" },
] });
add("matrix", "评测表每一行都要包含输入、标准和版本", { phase: "EVALUATION", sourceSection: 6, headers: ["字段组", "典型字段", "作用"], widths: [250, 550, 340], rows: [
  ["样本输入", "文案、业务场景、产品事实、渠道", "还原模型实际可见信息"],
  ["标准答案", "风险类型、等级、位置、预期处理", "明确什么算正确"],
  ["样本元数据", "来源、重要性、确认人、规则版本", "保证可追溯"],
  ["运行记录", "模型输出、版本、评分和Bad Case", "支持比较与回归"],
] });
add("splitFlow", "首版评测集的整理过程", { phase: "EVALUATION", sourceSection: 6, top: ["确定MVP风险范围", "提取历史样本", "去重清理并补上下文"], bottom: ["业务复核标签", "补高风险与边界", "形成首版评测集"] });
add("compare", "没有历史数据，也可以建立首版评测集", { phase: "EVALUATION", sourceSection: 6, left: { h: "启动阶段", items: ["从规则文档整理风险类型", "收集培训中的正反案例", "业务人员编写典型违规样本"] }, right: { h: "持续替换", items: ["模型生成表达变体后人工确认", "两名业务人员独立判断分歧", "影子运行后用真实案例替换构造数据"] } });
add("cards", "开发、验收和回归样本要分开管理", { phase: "EVALUATION", sourceSection: 6, cols: 3, items: [
  { k: "开发样本", v: "产品、算法和研发可以反复查看，用于设计规则与Prompt。" },
  { k: "验收样本", v: "方案基本稳定后判断是否达到上线条件，避免针对题目调答案。" },
  { k: "回归样本", v: "重要错误、高风险和线上问题；规则、模型或Prompt变化后重跑。" },
] });
add("matrix", "PRD中的验收约定至少写清六项", { phase: "EVALUATION", sourceSection: 6, headers: ["约定项", "需要明确"], widths: [310, 830], rowHeight: 58, rows: [
  ["评测对象", "单独模型、组合审核结果，还是完整产品闭环"],
  ["样本来源", "历史、近期、高风险、争议和构造样本分别从哪里来"],
  ["风险范围", "本次MVP必须覆盖什么，暂不验收什么"],
  ["标准答案责任人", "谁确认风险类型、等级和预期动作"],
  ["上线门槛", "高风险底线、识别要求、输出可用性和人工承载量"],
  ["版本变化", "规则、知识、模型、Prompt或结构变化后何时重新验收"],
] });

add("section", "07 组装完整AI产品PRD", { phase: "PRD", sourceSection: 7, subtitle: "把前六节的产品判断装进一份可执行文档" });
add("compare", "AI产品PRD比普通PRD多写三类内容", { phase: "PRD", sourceSection: 7, left: { h: "普通PRD通常关注", items: ["用户、场景与业务流程", "页面、功能与状态", "权限、数据和异常路径"] }, right: { h: "AI产品PRD还必须关注", items: ["AI介入依据和能力边界", "输入、输出、版本与失败接管", "评测集来源和上线门槛"] } });
add("matrix", "最终PRD由十三个章节组成", { phase: "PRD", sourceSection: 7, headers: ["01—04 业务", "05—08 方案与范围", "09—13 能力与保障"], widths: [380, 380, 380], rowHeight: 77, fontSize: 9.2, rows: [
  ["项目摘要", "AI介入判断与备选方案", "AI输入、输出与边界"],
  ["业务背景与现状", "能力分工", "数据、规则与知识依赖"],
  ["用户、场景与流程", "审核结果与处理流程", "评测集与验收条件"],
  ["问题、目标与非目标", "MVP与试点范围", "人工复核、异常、风险与待确认项"],
] });
add("summary", "PRD首页要在一页说清项目", { phase: "PRD", sourceSection: 7, items: [
  "业务问题：内容量增加导致审核积压、标准不一和结果难沉淀",
  "产品方案：规则＋RAG＋大模型＋人工的分层机器预审",
  "MVP范围：文本、单一业务试点、修改复核重审闭环",
  "验收方式：能力效果、人工效率、安全底线和失败接管",
] });
add("flow", "前六节内容按同一顺序组装进PRD", { phase: "PRD", sourceSection: 7, small: true, items: ["背景目标", "AI方案", "结果流程", "MVP范围", "能力契约", "评测验收", "风险待确认"] });
add("matrix", "数据、规则和知识依赖要单独说明", { phase: "PRD", sourceSection: 7, headers: ["依赖", "需要说明", "典型风险"], widths: [250, 540, 350], rows: [
  ["业务数据", "内容、产品事实、渠道和历史记录", "字段缺失、事实未结构化"],
  ["规则", "适用范围、版本、责任人和更新机制", "标准冲突、版本不一致"],
  ["知识库", "来源、切分、检索和有效期", "召回不足、过期依据"],
  ["模型与Prompt", "版本、配置、输出结构和回归要求", "波动、格式错误、成本与延迟"],
] });
add("matrix", "风险与待确认项不是缺陷，而是决策清单", { phase: "PRD", sourceSection: 7, headers: ["待确认项", "当前情况", "影响"], widths: [330, 430, 380], rows: [
  ["人工审核基线", "需要从业务系统补充", "效率目标与收益测算"],
  ["高风险判断标准", "需要规则负责人确认", "自动处理范围"],
  ["历史数据质量", "尚未完成抽样检查", "评测集可靠性"],
  ["模型响应与复核承载", "需要PoC和业务测算", "体验与风险阈值"],
] });
add("compare", "课堂用两项交付检验是否真正掌握", { phase: "PRD", sourceSection: 7, left: { h: "练习一｜PRD首页", items: ["业务问题与目标用户", "产品目标和混合审核方案", "MVP范围与验收方式"] }, right: { h: "练习二｜AI能力卡", items: ["模型处理什么、能获得什么", "必须输出什么、何时转人工", "模型明确不负责什么"] } });
add("closing", "第10课的完整产品设计链路", { phase: "OVERVIEW", sourceSection: 7, items: ["发现业务问题", "判断AI介入", "划分能力责任", "设计结果流程", "确定MVP", "定义能力契约", "约定评测验收", "组装完整PRD"] });
add("next", "PRD完成以后，还要回答：这个项目值得投入吗？", { phase: "NEXT", sourceSection: 7, subtitle: "第11课：模型选择、调用成本、人工成本与业务收益", quote: "产品方案可开发、可验收，不代表它在经济上一定值得做。" });

function render(d, i, total) {
  const slide = pptx.addSlide("BLANK_DARK");
  slide.background = { color: C.bg };
  addNotes(slide, d);
  if (d.type === "cover") {
    addShape(slide, "rect", P(0, 0, 22, 720), C.teal);
    addText(slide, "AI PRODUCT MANAGER", P(82, 72, 400, 30), 10, C.teal, true);
    addText(slide, d.title, P(80, 220, 1080, 92), 35, C.ink, true);
    addText(slide, d.subtitle, P(84, 348, 700, 42), 15.5, C.muted);
    addShape(slide, "rect", P(84, 462, 260, 6), C.orange);
    addText(slide, "BUSINESS · SOLUTION · MVP · CAPABILITY · EVALUATION · PRD", P(84, 500, 850, 30), 9.5, C.muted, true);
    addText(slide, "10", P(1010, 478, 150, 90), 47, C.line, true, "right");
    return;
  }

  addChrome(slide, i, total, d.phase || "PRD");
  if (d.type === "section") {
    addText(slide, d.title, P(80, 182, 1080, 88), 34, C.ink, true);
    addShape(slide, "rect", P(80, 300, 120, 6), C.orange);
    addText(slide, d.subtitle, P(84, 345, 980, 50), 16.5, C.muted);
    return;
  }

  if (d.type === "case") return caseSlide(slide, d);
  addTitle(slide, d.title, d.subtitle || "");

  if (d.type === "map" || d.type === "flow") flow(slide, d.items, { y: d.type === "map" ? 292 : 300, small: d.small });
  else if (d.type === "splitFlow") splitFlow(slide, d.top, d.bottom);
  else if (d.type === "cards") cards(slide, d.items, { cols: d.cols || 3, y: 210, h: d.items.length > (d.cols || 3) ? 148 : 180 });
  else if (d.type === "compare") compare(slide, d.left, d.right);
  else if (d.type === "matrix") matrix(slide, d.headers, d.rows, { widths: d.widths, y: 205, rowHeight: d.rowHeight || 58, fontSize: d.fontSize || 10.3 });
  else if (d.type === "quote") { bulletList(slide, d.items || [], 100, 220, 1060, 13.5, 53); quote(slide, d.quote, 450); }
  else if (d.type === "formula") { quote(slide, d.formula, 226, C.teal); bulletList(slide, d.items, 150, 362, 950, 14, 62); }
  else if (d.type === "stacked") stacked(slide, d.layers);
  else if (d.type === "code") codeBlock(slide, d.code, 205, 400, d.codeSize || 10.5);
  else if (d.type === "bars") bars(slide, d.items, 216);
  else if (d.type === "lifecycle") lifecycle(slide, d.items);
  else if (d.type === "summary") cards(slide, d.items.map((v, j) => ({ k: String(j + 1).padStart(2, "0"), v })), { cols: 2, y: 220, h: 140 });
  else if (d.type === "closing") { flow(slide, d.items, { y: 260, small: true }); quote(slide, "把一个模糊的AI想法，变成边界清楚、能够开发、可以验收的产品方案。", 470); }
  else if (d.type === "next") { quote(slide, d.quote, 260, C.orange); addText(slide, d.subtitle, P(150, 400, 980, 52), 18, C.teal, true, "center"); }
}

const outlineLines = [
  "# 第10课 PPT结构与逐页内容｜核心结构版",
  "",
  `- 生成页数：${S.length}页`,
  "- 设计基准：第09课《AI产品评测与数据反馈》核心结构版",
  "- 内容原则：严格保持7节讲义顺序，合并相邻知识点，避免严格按小标题逐页拆分。",
  "",
  ...S.map((d, i) => `${String(i + 1).padStart(2, "0")}. **${d.title}**｜${d.type}｜${d.phase || "PRD"}`),
  "",
];

await fs.writeFile(OUTLINE, outlineLines.join("\n"), "utf8");
S.forEach((d, i) => render(d, i, S.length));
await pptx.writeFile({ fileName: OUT });
console.log(JSON.stringify({ slides: S.length, out: OUT, outline: OUTLINE }));
