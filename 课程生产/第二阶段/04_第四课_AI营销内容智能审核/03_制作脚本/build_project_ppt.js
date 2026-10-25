const pptxgen = require('pptxgenjs');
const path = require('path');

const ASSET_DIR = path.resolve(__dirname, '../01_讲义与材料/配图');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理课程';
pptx.subject = 'AI营销内容智能审核项目实战';
pptx.title = '第四课｜AI营销内容智能审核项目实战';
pptx.company = 'AI-Course';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Hiragino Sans GB',
  bodyFontFace: 'Hiragino Sans GB',
  lang: 'zh-CN'
};
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';
pptx.margin = 0;

const C = {
  bg: '04171E',
  panel: '0B2932',
  panel2: '103641',
  teal: '27D3C1',
  teal2: '7BE7DC',
  orange: 'FF6B4A',
  white: 'F7FBFC',
  muted: 'A7BBC2',
  line: '28505B',
  red: 'FF775F',
  yellow: 'F6C453',
  green: '4ED6A5',
  ink: '071B22'
};

const S = pptx.ShapeType;
let page = 0;

function addBg(slide) {
  slide.background = { color: C.bg };
}

function addFooter(slide, label = 'AI营销内容智能审核项目实战') {
  page += 1;
  slide.addText(label, { x: 0.55, y: 7.12, w: 5.3, h: 0.18, fontFace: 'Hiragino Sans GB', fontSize: 7.5, color: '66828B', margin: 0 });
  slide.addText(String(page).padStart(2, '0'), { x: 12.2, y: 7.08, w: 0.55, h: 0.22, fontFace: 'Arial', fontSize: 8, color: C.muted, align: 'right', margin: 0 });
}

function addHeader(slide, title, kicker = 'PROJECT 04', subtitle = '') {
  slide.addText(kicker, { x: 0.62, y: 0.42, w: 2.8, h: 0.22, fontSize: 9, bold: true, color: C.teal, charSpacing: 1.4, margin: 0 });
  slide.addText(title, { x: 0.62, y: 0.76, w: 11.9, h: 0.52, fontSize: 26, bold: true, color: C.white, margin: 0, breakLine: false, fit: 'shrink' });
  slide.addShape(S.line, { x: 0.62, y: 1.42, w: 12.05, h: 0, line: { color: C.line, width: 1 } });
  slide.addShape(S.line, { x: 0.62, y: 1.42, w: 1.45, h: 0, line: { color: C.teal, width: 3 } });
  if (subtitle) slide.addText(subtitle, { x: 8.1, y: 0.46, w: 4.55, h: 0.22, fontSize: 8.5, color: C.muted, align: 'right', margin: 0, fit: 'shrink' });
}

function addKey(slide, text) {
  slide.addShape(S.roundRect, { x: 0.62, y: 6.55, w: 12.05, h: 0.43, rectRadius: 0.05, fill: { color: '0A222B' }, line: { color: '16414B', width: 0.8 } });
  slide.addText('本页重点', { x: 0.84, y: 6.67, w: 0.9, h: 0.16, fontSize: 8, bold: true, color: C.teal, margin: 0 });
  slide.addText(text, { x: 1.76, y: 6.62, w: 10.55, h: 0.23, fontSize: 10, color: C.white, margin: 0, fit: 'shrink' });
}

function contentSlide(title, key, notes, opts = {}) {
  const slide = pptx.addSlide();
  addBg(slide);
  addHeader(slide, title, opts.kicker || 'PROJECT 04', opts.subtitle || 'AI营销内容智能审核');
  if (key) addKey(slide, key);
  addFooter(slide);
  if (notes) slide.addNotes(notes);
  return slide;
}

function fullImageSlide(relativePath, notes = '') {
  const slide = pptx.addSlide();
  addBg(slide);
  slide.addImage({ path: path.join(ASSET_DIR, relativePath), x: 0, y: 0, w: 13.333, h: 7.5 });
  page += 1;
  if (notes) slide.addNotes(notes);
  return slide;
}

function sectionSlide(no, title, line, timing, notes) {
  const slide = pptx.addSlide();
  addBg(slide);
  slide.addText(String(no).padStart(2, '0'), { x: 0.68, y: 0.65, w: 2.6, h: 1.8, fontFace: 'Arial', fontSize: 90, bold: true, color: C.teal, transparency: 5, margin: 0 });
  slide.addShape(S.line, { x: 3.5, y: 1.04, w: 0, h: 5.2, line: { color: C.line, width: 1 } });
  slide.addText(title, { x: 4.05, y: 1.52, w: 8.3, h: 1.1, fontSize: 34, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addText(line, { x: 4.08, y: 2.9, w: 7.85, h: 1.1, fontSize: 17, color: C.muted, breakLine: false, margin: 0, fit: 'shrink' });
  slide.addShape(S.roundRect, { x: 4.08, y: 4.55, w: 2.25, h: 0.48, rectRadius: 0.05, fill: { color: C.panel2 }, line: { color: C.teal, width: 1 } });
  slide.addText(timing, { x: 4.28, y: 4.7, w: 1.85, h: 0.17, fontSize: 10, bold: true, color: C.teal2, align: 'center', margin: 0 });
  addFooter(slide, 'AI营销内容智能审核项目实战 · SECTION');
  slide.addNotes(notes);
  return slide;
}

function card(slide, x, y, w, h, title, body, accent = C.teal, opts = {}) {
  slide.addShape(S.roundRect, { x, y, w, h, rectRadius: 0.06, fill: { color: opts.fill || C.panel, transparency: opts.transparency || 0 }, line: { color: opts.line || C.line, width: 0.8 } });
  slide.addShape(S.rect, { x, y, w: 0.08, h, fill: { color: accent }, line: { color: accent, transparency: 100 } });
  slide.addText(title, { x: x + 0.28, y: y + 0.2, w: w - 0.5, h: 0.28, fontSize: opts.titleSize || 15, bold: true, color: opts.titleColor || C.white, margin: 0, fit: 'shrink' });
  slide.addText(body, { x: x + 0.28, y: y + 0.65, w: w - 0.5, h: h - 0.83, fontSize: opts.bodySize || 11.5, color: opts.bodyColor || C.muted, margin: 0, breakLine: false, valign: 'top', fit: 'shrink', bullet: opts.bullet ? { type: 'bullet' } : undefined });
}

function pill(slide, x, y, w, text, color = C.teal, fill = C.panel2) {
  slide.addShape(S.roundRect, { x, y, w, h: 0.38, rectRadius: 0.08, fill: { color: fill }, line: { color, width: 0.8 } });
  slide.addText(text, { x: x + 0.06, y: y + 0.11, w: w - 0.12, h: 0.15, fontSize: 9, bold: true, color, align: 'center', margin: 0, fit: 'shrink' });
}

function arrow(slide, x1, y1, x2, y2, color = C.teal) {
  slide.addShape(S.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color, width: 1.6, beginArrowType: 'none', endArrowType: 'triangle' } });
}

function flowNode(slide, x, y, w, h, title, sub = '', accent = C.teal) {
  slide.addShape(S.roundRect, { x, y, w, h, rectRadius: 0.06, fill: { color: C.panel }, line: { color: accent, width: 1.2 } });
  slide.addText(title, { x: x + 0.12, y: y + 0.18, w: w - 0.24, h: 0.25, fontSize: 12, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
  if (sub) slide.addText(sub, { x: x + 0.12, y: y + 0.5, w: w - 0.24, h: h - 0.63, fontSize: 8.5, color: C.muted, align: 'center', margin: 0, fit: 'shrink' });
}

function smallLabel(slide, x, y, text, color = C.teal) {
  slide.addText(text, { x, y, w: 1.4, h: 0.18, fontSize: 8.5, bold: true, color, margin: 0, fit: 'shrink' });
}

// 01 Cover
{
  const slide = pptx.addSlide();
  addBg(slide);
  slide.addShape(S.rect, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.teal }, line: { color: C.teal } });
  slide.addText('SECOND STAGE · PROJECT COURSE', { x: 0.72, y: 0.65, w: 5.5, h: 0.25, fontSize: 10, bold: true, color: C.teal, charSpacing: 1.7, margin: 0 });
  slide.addText('AI营销内容智能审核', { x: 0.72, y: 1.55, w: 8.7, h: 0.75, fontSize: 38, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addText('项目实战', { x: 0.72, y: 2.48, w: 5.1, h: 0.72, fontSize: 38, bold: true, color: C.white, margin: 0 });
  slide.addText('从业务流程、混合审核到产品页面与评测闭环', { x: 0.75, y: 3.55, w: 7.7, h: 0.35, fontSize: 16, color: C.muted, margin: 0 });
  slide.addText('04', { x: 9.92, y: 1.18, w: 2.4, h: 1.55, fontFace: 'Arial', fontSize: 90, bold: true, color: C.teal, align: 'center', margin: 0 });
  slide.addShape(S.roundRect, { x: 9.4, y: 3.05, w: 3.2, h: 1.65, rectRadius: 0.08, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
  slide.addText('180 MIN', { x: 9.8, y: 3.46, w: 2.4, h: 0.32, fontFace: 'Arial', fontSize: 20, bold: true, color: C.teal2, align: 'center', margin: 0 });
  slide.addText('完整项目课', { x: 9.8, y: 4.03, w: 2.4, h: 0.22, fontSize: 11, color: C.white, align: 'center', margin: 0 });
  slide.addText('老师授课版', { x: 0.75, y: 6.55, w: 2.2, h: 0.25, fontSize: 10, bold: true, color: C.orange, margin: 0 });
  addFooter(slide, 'AI产品经理课程 · 第二阶段');
  slide.addNotes('开场先说明：今天不是重新上一遍AI理论课，而是一起完成一个可放进作品集的智能审核项目。整堂课始终围绕同一条营销文案和同一条业务流程推进。');
}

// 02 Outcomes
{
  const slide = contentSlide('今天要完成什么', '结束时，学员拿走的不是一组概念，而是一套完整项目结构。', '快速说明六项产出。提醒学员先关注方案闭环，不要一开始纠结页面美观。');
  const items = [
    ['01', '业务问题', '现状流程、角色、断点'],
    ['02', '项目边界', '第一版做什么、不做什么'],
    ['03', '混合方案', '规则、模型、资料、人工分工'],
    ['04', '规则流程', '场景、原子规则、结果路由'],
    ['05', '核心页面', '列表、提交、工作台、复核'],
    ['06', '评测迭代', '样本、指标、坏案例、版本']
  ];
  items.forEach((it, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.7 + col * 4.15, y = 1.78 + row * 2.15;
    slide.addText(it[0], { x, y: y + 0.04, w: 0.55, h: 0.3, fontFace: 'Arial', fontSize: 18, bold: true, color: C.teal, margin: 0 });
    card(slide, x + 0.55, y, 3.35, 1.7, it[1], it[2], i === 5 ? C.orange : C.teal, { bodySize: 11 });
  });
}

// 03 schedule
{
  const slide = contentSlide('180分钟怎么走', '两次休息把课程切成三段：理解项目、完成方案、做出产品并验证。', '按照时间轴介绍课堂节奏。互动环节可以根据班级情况缩短，但不要压缩核心页面和坏案例。');
  const times = [
    ['0–10', '开场案例'], ['10–40', '业务与边界'], ['40–65', 'AI介入'], ['65–75', '休息'],
    ['75–110', '规则与流程'], ['110–135', 'MVP与架构'], ['135–145', '休息'], ['145–165', '核心页面'], ['165–180', '评测与作业']
  ];
  slide.addShape(S.line, { x: 1.0, y: 3.65, w: 11.25, h: 0, line: { color: C.line, width: 2 } });
  times.forEach((t, i) => {
    const x = 0.85 + i * 1.37;
    const breakItem = t[1] === '休息';
    slide.addShape(S.ellipse, { x, y: 3.43, w: 0.45, h: 0.45, fill: { color: breakItem ? C.orange : C.teal }, line: { color: C.bg, width: 2 } });
    slide.addText(t[0], { x: x - 0.18, y: 2.72, w: 0.82, h: 0.22, fontSize: 9, bold: true, color: breakItem ? C.orange : C.teal2, align: 'center', margin: 0 });
    slide.addText(t[1], { x: x - 0.38, y: 4.18, w: 1.2, h: 0.55, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'mid', margin: 0, fit: 'shrink' });
  });
}

// 04 final artifact preview
{
  const slide = contentSlide('最终产品长什么样', '四个页面覆盖提交、检查、处理和人工接管，页面背后还有规则与状态。', '先给学员看成品结构，建立期待。强调后面每一个业务判断都会落实到这四个页面。');
  const labels = ['审核任务列表', '内容提交页', '智能审核工作台', '人工复核页'];
  labels.forEach((label, i) => {
    const x = 0.72 + i * 3.05;
    slide.addShape(S.roundRect, { x, y: 1.82, w: 2.72, h: 3.82, rectRadius: 0.06, fill: { color: C.panel }, line: { color: i === 2 ? C.orange : C.teal, width: i === 2 ? 2 : 1 } });
    slide.addShape(S.rect, { x: x + 0.18, y: 2.08, w: 2.36, h: 0.3, fill: { color: C.panel2 }, line: { color: C.line, width: 0.5 } });
    for (let r = 0; r < 4; r++) {
      slide.addShape(S.roundRect, { x: x + 0.23, y: 2.65 + r * 0.55, w: 2.22, h: 0.36, rectRadius: 0.03, fill: { color: r === 1 && i === 2 ? '4A2C2B' : '0E333D' }, line: { color: r === 1 && i === 2 ? C.orange : C.line, width: 0.5 } });
    }
    slide.addText(label, { x: x + 0.15, y: 5.08, w: 2.42, h: 0.28, fontSize: 12, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
    slide.addText(`0${i + 1}`, { x: x + 1.03, y: 5.48, w: 0.7, h: 0.22, fontFace: 'Arial', fontSize: 10, bold: true, color: i === 2 ? C.orange : C.teal, align: 'center', margin: 0 });
  });
}

// 05 case
{
  const slide = contentSlide('先审一段文案', '先不要讨论技术，先判断系统到底需要帮助用户完成什么。', '给学员90秒独立思考，再随机邀请两人回答。先不评价法律结论，只收集他们认为需要检查的位置。');
  slide.addShape(S.roundRect, { x: 0.78, y: 1.83, w: 8.05, h: 3.5, rectRadius: 0.06, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
  slide.addText('“', { x: 1.03, y: 1.94, w: 0.55, h: 0.65, fontSize: 44, color: C.teal, margin: 0 });
  slide.addText('全网最值得买的胶原蛋白饮，7 天彻底改善熬夜暗沉，100% 有效。连续三年销量第一，三甲医院专家一致推荐。今晚下单立减 50%，错过再等一年。', { x: 1.37, y: 2.36, w: 6.88, h: 2.2, fontSize: 22, bold: true, color: C.white, breakLine: false, margin: 0.04, fit: 'shrink', valign: 'mid' });
  card(slide, 9.2, 1.83, 3.15, 1.45, '问题 1', '哪些位置需要进一步检查？', C.orange, { bodySize: 12 });
  card(slide, 9.2, 3.55, 3.15, 1.78, '问题 2', '系统只说“有问题”够不够？还应该返回什么？', C.teal, { bodySize: 12 });
  pill(slide, 9.55, 5.7, 2.45, '90 秒独立判断', C.yellow, '332E20');
}

// 06 audience task
{
  const slide = contentSlide('把你的判断写成产品输出', '从“我觉得不妥”升级为“系统可以返回并驱动处理的结果”。', '让学员按五个字段组织答案。此处的目标不是判断对错，而是让他们感受到产品输出需要结构化。');
  const cols = [
    ['问题位置', '原文哪几个字'], ['问题类型', '属于哪类检查'], ['风险等级', '优先级多高'], ['判断说明', '为什么要处理'], ['下一步动作', '修改、补材料或转人工']
  ];
  cols.forEach((c, i) => {
    const x = 0.75 + i * 2.48;
    slide.addShape(S.ellipse, { x: x + 0.78, y: 1.9, w: 0.72, h: 0.72, fill: { color: i === 4 ? C.orange : C.teal }, line: { color: C.bg, width: 2 } });
    slide.addText(String(i + 1), { x: x + 0.96, y: 2.1, w: 0.35, h: 0.2, fontFace: 'Arial', fontSize: 13, bold: true, color: C.ink, align: 'center', margin: 0 });
    card(slide, x, 2.83, 2.28, 2.32, c[0], c[1], i === 4 ? C.orange : C.teal, { titleSize: 14, bodySize: 11, fill: '0A2630' });
  });
}

// 07 reveal
{
  const slide = contentSlide('这不是文案纠错器', '真正的产品要把一次判断转化为可定位、可处理、可复核的业务动作。', '揭示项目定义。注意不要直接说这些表达一定违法，只说需要进一步检查、补充材料或进入人工判断。');
  const nodes = [
    ['输入内容', '文案＋场景＋材料'], ['识别问题', '规则与语义检查'], ['定位原文', '字符位置与片段'], ['给出动作', '修改／补充／转人工'], ['完成复核', '记录最终结论']
  ];
  nodes.forEach((n, i) => {
    const x = 0.72 + i * 2.5;
    flowNode(slide, x, 2.55, 2.05, 1.55, n[0], n[1], i === 3 ? C.orange : C.teal);
    if (i < nodes.length - 1) arrow(slide, x + 2.08, 3.32, x + 2.43, 3.32, C.muted);
  });
  slide.addText('智能审核平台', { x: 4.52, y: 4.65, w: 4.3, h: 0.5, fontSize: 24, bold: true, color: C.teal2, align: 'center', margin: 0 });
}

// 08 section 1
sectionSlide(1, '理解业务与限定项目', '先把原流程、角色、问题和第一版边界说清，再谈技术。', '30 MIN', '进入第一节。提醒学员：项目的第一步不是列功能，而是限定业务问题。');

// 09 setting
{
  const slide = contentSlide('项目发生在什么业务里', '同一家公司、多个内容渠道、持续增加的审核任务。', '介绍案例公司。可以让学员代入常见消费品、电商或内容运营团队。');
  card(slide, 0.75, 1.82, 3.65, 3.9, '内容生产', '商品详情页\n活动页\n短视频口播稿\n达人合作文案\n私域推送', C.teal, { bodySize: 15 });
  card(slide, 4.83, 1.82, 3.65, 3.9, '发布渠道', '自有商城\n电商平台\n内容社区\n短视频平台\n企业私域', C.teal2, { bodySize: 15 });
  card(slide, 8.91, 1.82, 3.65, 3.9, '审核压力', '内容量增长\n活动高峰排队\n规则持续更新\n历史意见分散\n高风险仍需人工', C.orange, { bodySize: 15 });
  arrow(slide, 4.42, 3.76, 4.76, 3.76, C.muted);
  arrow(slide, 8.5, 3.76, 8.84, 3.76, C.muted);
}

// 10 current flow
{
  const slide = contentSlide('当前人工审核流程', '流程并不缺步骤，问题在于大量重复检查和信息无法复用。', '顺着流程询问：每一步会留下什么信息？哪一步容易返工？');
  const nodes = ['运营撰写', '发送审核', '人工逐段检查', '通过／退回／补材料', '运营修改', '发布'];
  nodes.forEach((n, i) => {
    const x = 0.62 + i * 2.08;
    flowNode(slide, x, 2.67, 1.7, 1.15, n, i === 2 ? '高重复工作' : '', i === 2 || i === 3 ? C.orange : C.teal);
    if (i < nodes.length - 1) arrow(slide, x + 1.72, 3.24, x + 2.02, 3.24, C.muted);
  });
  slide.addShape(S.arc, { x: 6.72, y: 4.1, w: 3.3, h: 1.2, adjustPoint: 0.3, rotate: 0, line: { color: C.orange, width: 1.5, beginArrowType: 'none', endArrowType: 'triangle' }, fill: { color: C.bg, transparency: 100 } });
  slide.addText('反复修改与再次提交', { x: 7.13, y: 5.15, w: 2.5, h: 0.23, fontSize: 10, bold: true, color: C.orange, align: 'center', margin: 0 });
}

// 11 pain points
{
  const slide = contentSlide('四类业务问题', '人工不是“没有价值”，而是被大量重复工作占用。', '逐一说明问题，并补充：高风险判断仍然需要人工，目标不是全自动替代。');
  card(slide, 0.75, 1.82, 2.85, 3.85, '数量', '活动高峰任务排队\n常规内容等待时间变长', C.teal, { bodySize: 14 });
  card(slide, 3.75, 1.82, 2.85, 3.85, '一致性', '同一句表达\n不同审核员处理不同', C.teal2, { bodySize: 14 });
  card(slide, 6.75, 1.82, 2.85, 3.85, '变化', '平台、行业与内部口径\n持续更新', C.yellow, { bodySize: 14 });
  card(slide, 9.75, 1.82, 2.85, 3.85, '沉淀', '审核意见散落\n难以统计和复用', C.orange, { bodySize: 14 });
}

// 12 roles
{
  const slide = contentSlide('项目中有四类角色', '如果只画“运营＋AI”，项目还没有进入真实业务。', '让学员注意规则运营和业务负责人，这两个角色决定系统能不能持续维护和评估。');
  const data = [
    ['内容运营', '提交、修改、重审', '哪里有问题，怎么改'],
    ['人工审核员', '处理高风险与不确定项', '判断是否可靠，信息是否完整'],
    ['规则运营', '维护场景、规则和版本', '哪条规则生效，效果如何'],
    ['业务负责人', '评估效率、风险和试点', '是否值得扩大使用']
  ];
  data.forEach((d, i) => {
    const y = 1.78 + i * 1.1;
    slide.addShape(S.roundRect, { x: 0.78, y, w: 11.78, h: 0.82, rectRadius: 0.04, fill: { color: i % 2 ? '0A2630' : C.panel }, line: { color: C.line, width: 0.6 } });
    slide.addText(d[0], { x: 1.02, y: y + 0.24, w: 1.72, h: 0.22, fontSize: 13, bold: true, color: i === 1 ? C.orange : C.teal2, margin: 0 });
    slide.addText(d[1], { x: 3.22, y: y + 0.24, w: 3.0, h: 0.22, fontSize: 12, color: C.white, margin: 0 });
    slide.addText(d[2], { x: 6.65, y: y + 0.24, w: 5.2, h: 0.22, fontSize: 12, color: C.muted, margin: 0 });
  });
  smallLabel(slide, 3.22, 1.5, '核心任务');
  smallLabel(slide, 6.65, 1.5, '最关心的问题');
}

// 13 breakpoints
{
  const slide = contentSlide('原流程的三个断点', 'AI 的介入点来自流程断点，而不是来自技术热度。', '分别对应提交前、审核中、修改后的问题。后续产品页面要逐一解决。');
  card(slide, 0.82, 1.92, 3.68, 3.72, '01 提交前', '没有自查\n低级问题也进入人工队列', C.teal, { bodySize: 16 });
  card(slide, 4.83, 1.92, 3.68, 3.72, '02 退回时', '意见不结构化\n运营只知道“要改”', C.orange, { bodySize: 16 });
  card(slide, 8.84, 1.92, 3.68, 3.72, '03 修改后', '版本关联弱\n无法确认哪些问题已关闭', C.teal2, { bodySize: 16 });
}

// 14 boundary
{
  const slide = contentSlide('第一版边界', '先把一条小链路做完整，再扩内容类型、渠道和自动化程度。', '读完左右两列后追问：为什么多模态、自动覆盖和自动发布不进第一版？');
  card(slide, 0.82, 1.78, 5.72, 4.35, 'MVP 先做', '一个业务团队试点\n商品详情页＋活动文案\n中文纯文本\n提交前机器预审\n四类重点检查\n人工复核与版本记录', C.teal, { bodySize: 15 });
  card(slide, 6.78, 1.78, 5.72, 4.35, 'MVP 暂不做', '图片、视频、音频\n自动覆盖运营原文\n自动发布\n所有渠道与业务线\n自主 Agent 审核', C.orange, { bodySize: 15 });
}

// 15 goals
{
  const slide = contentSlide('项目目标要写成三层', '“提升效率、降低风险”太空；目标必须能对应角色、流程和验证方式。', '用三层目标替代口号。后续指标要分别覆盖业务、用户和产品。');
  const rows = [
    ['业务目标', '缩短常规内容等待时间', '减少人工重复检查'],
    ['用户目标', '运营快速定位并处理问题', '审核员优先处理高价值任务'],
    ['产品目标', '形成提交—检查—处理—复核—回流闭环', '失败时可安全接管']
  ];
  rows.forEach((r, i) => {
    const y = 1.88 + i * 1.4;
    slide.addText(r[0], { x: 0.88, y: y + 0.32, w: 1.5, h: 0.3, fontSize: 15, bold: true, color: i === 2 ? C.orange : C.teal2, margin: 0 });
    card(slide, 2.5, y, 4.65, 1.08, r[1], '', i === 2 ? C.orange : C.teal, { titleSize: 13 });
    card(slide, 7.42, y, 4.72, 1.08, r[2], '', i === 2 ? C.orange : C.teal2, { titleSize: 13 });
  });
}

// 16 exercise
{
  const slide = contentSlide('课堂判断：哪些需求进入第一版', '是否进入 MVP，要看它能否直接验证核心闭环。', '让学员举手投票，再展示推荐结论。');
  const items = [
    ['自动生成十个营销标题', '后续模块', C.muted],
    ['检测高风险表达并定位原文', '进入 MVP', C.green],
    ['自动把通过内容发布到所有平台', '暂不进入', C.orange]
  ];
  items.forEach((it, i) => {
    const y = 1.9 + i * 1.35;
    slide.addShape(S.roundRect, { x: 1.0, y, w: 11.25, h: 0.98, rectRadius: 0.04, fill: { color: C.panel }, line: { color: it[2], width: 1 } });
    slide.addText(String(i + 1), { x: 1.27, y: y + 0.29, w: 0.45, h: 0.25, fontFace: 'Arial', fontSize: 15, bold: true, color: it[2], margin: 0 });
    slide.addText(it[0], { x: 2.05, y: y + 0.28, w: 6.3, h: 0.26, fontSize: 15, bold: true, color: C.white, margin: 0 });
    pill(slide, 9.6, y + 0.3, 1.8, it[1], it[2], '0E3038');
  });
}

// 17 section 2
sectionSlide(2, '决定 AI 怎么介入', '选择混合审核方案，并明确每种能力承担的任务与失败处理。', '25 MIN', '进入第二节。理论快速复习，只在方案选择和职责划分处停留。');

// 18 options
{
  const slide = contentSlide('四种方案放在一起比较', '选择混合方案，不是因为它最“高级”，而是因为它更适合风险审核。', '解释每种方案的适用位置。纯人工保留兜底，纯规则承担确定性任务，大模型负责语义。');
  const items = [
    ['纯人工', '灵活、理解复杂上下文', '慢、贵、口径波动', '保留兜底'],
    ['纯规则', '稳定、快速、容易解释', '难覆盖隐含语义', '负责确定性检查'],
    ['只用大模型', '语义覆盖广、配置快', '结果波动、责任边界弱', '不单独采用'],
    ['混合方案', '兼顾稳定、覆盖和安全', '设计与运营更复杂', '第一版采用']
  ];
  items.forEach((it, i) => {
    const x = 0.72 + i * 3.06;
    card(slide, x, 1.78, 2.76, 4.18, it[0], `${it[1]}\n\n${it[2]}\n\n→ ${it[3]}`, i === 3 ? C.orange : C.teal, { bodySize: 12.5, line: i === 3 ? C.orange : C.line });
  });
}

// 19 capability split
{
  const slide = contentSlide('五类能力，各自负责什么', '技术能力的边界必须被写成产品职责。', '逐列讲解。特别说明业务资料不等于一定使用RAG，也可以来自结构化字段或接口。');
  const items = [
    ['输入校验', '必填字段\n内容格式', C.teal],
    ['确定性规则', '明确表达\n格式与逻辑', C.teal2],
    ['大模型', '语义理解\n上下文判断', C.orange],
    ['业务资料', '商品事实\n授权与口径', C.yellow],
    ['人工复核', '高风险\n冲突与低确定性', C.red]
  ];
  items.forEach((it, i) => {
    const x = 0.67 + i * 2.5;
    card(slide, x, 2.05, 2.18, 3.5, it[0], it[1], it[2], { titleSize: 14, bodySize: 15, line: it[2] });
  });
}

// 20 not model only
{
  const slide = contentSlide('为什么不能只接一个大模型', '模型能“说一段话”，不等于产品已经能处理一个审核任务。', '用“销量第一”例子展开。让学员观察右侧六个产品问题，这些都不是一次自由文本输出能够稳定解决的。');
  slide.addShape(S.roundRect, { x: 0.78, y: 1.95, w: 4.15, h: 3.9, rectRadius: 0.06, fill: { color: C.panel }, line: { color: C.orange, width: 1.2 } });
  slide.addText('“连续三年销量第一”', { x: 1.15, y: 2.55, w: 3.4, h: 0.85, fontSize: 24, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
  slide.addText('只返回：\n“可能存在夸大宣传风险”', { x: 1.18, y: 4.05, w: 3.35, h: 0.75, fontSize: 14, color: C.muted, align: 'center', margin: 0 });
  const qs = ['命中哪类问题？', '原文位置在哪？', '缺少什么信息？', '风险等级多高？', '下一步怎么处理？', '修改后是否重审？'];
  qs.forEach((q, i) => {
    const x = 5.48 + (i % 2) * 3.35, y = 1.92 + Math.floor(i / 2) * 1.27;
    card(slide, x, y, 3.02, 0.92, q, '', i === 4 ? C.orange : C.teal, { titleSize: 13 });
  });
}

// 21 workflow
{
  const slide = contentSlide('产品主流程', '大模型只是中间一个节点；产品要负责加载、校验、合并、路由和记录。', '从左到右走一遍主流程。提醒学员：结果合并和路由是产品设计中的核心步骤。');
  const top = [
    ['选择场景并提交', '文案＋上下文'], ['输入校验', '必填与格式'], ['加载规则版本', '场景启用项'], ['并行检查', '规则＋模型'], ['结果合并', '去重与定级']
  ];
  top.forEach((n, i) => {
    const x = 0.62 + i * 2.52;
    flowNode(slide, x, 2.0, 2.08, 1.22, n[0], n[1], i === 3 ? C.orange : C.teal);
    if (i < top.length - 1) arrow(slide, x + 2.1, 2.61, x + 2.45, 2.61, C.muted);
  });
  const results = [['通过', C.green], ['建议修改', C.teal2], ['退回修改', C.orange], ['转人工', C.red]];
  results.forEach((r, i) => pill(slide, 1.42 + i * 2.75, 4.65, 2.15, r[0], r[1], '0B2A33'));
  arrow(slide, 6.9, 3.35, 6.9, 4.35, C.muted);
}

// 22 exceptions
{
  const slide = contentSlide('失败和不确定性也要设计', '机器失败不能默认通过，不确定结果不能伪装成确定结论。', '逐行讲解异常处理。这里是区分演示型产品和可试点产品的关键。');
  const rows = [
    ['必填上下文缺失', '阻止提交并提示补充'],
    ['模型超时', '自动重试一次；仍失败则进人工队列'],
    ['输出格式失败', '重新请求结构化输出；仍失败则标记机器失败'],
    ['规则与模型冲突', '保留原始结果，按高风险优先或转人工'],
    ['内容已被修改', '旧结果失效，必须重新审核']
  ];
  rows.forEach((r, i) => {
    const y = 1.72 + i * 0.87;
    slide.addShape(S.roundRect, { x: 0.85, y, w: 11.65, h: 0.67, rectRadius: 0.03, fill: { color: i % 2 ? '0A2630' : C.panel }, line: { color: C.line, width: 0.5 } });
    slide.addText(r[0], { x: 1.08, y: y + 0.2, w: 2.6, h: 0.22, fontSize: 12.5, bold: true, color: i === 2 ? C.orange : C.teal2, margin: 0 });
    slide.addText(r[1], { x: 3.95, y: y + 0.2, w: 7.9, h: 0.22, fontSize: 12, color: C.white, margin: 0 });
  });
}

// 23 workflow vs agent
{
  const slide = contentSlide('第一版选择 Workflow', '步骤明确、分支有限、责任敏感的审核链路，更适合先做可控流程。', '用四个条件快速复习Workflow选择。Agent可以作为后续复杂材料收集或多轮追问的扩展。');
  card(slide, 0.85, 1.9, 5.55, 3.95, '为什么适合 Workflow', '步骤可以预先定义\n输入输出可以校验\n超时与重试可以控制\n人工接管位置明确', C.teal, { bodySize: 16 });
  card(slide, 6.9, 1.9, 5.55, 3.95, 'Agent 以后可以做什么', '多轮追问缺失材料\n辅助查找复杂业务资料\n处理开放式运营任务\n但不作为第一版必要条件', C.orange, { bodySize: 16 });
}

fullImageSlide('第二节_AI介入方案/01_智能审核的五类能力分工_v2.png', '用这张图总结五类能力的职责边界。先读每一列，再强调明确问题交给规则、语义问题交给模型、不确定问题交给人工。');
fullImageSlide('第二节_AI介入方案/02_营销内容智能审核主流程_v2.png', '按照五个阶段走一遍主流程。重点说明并行审核、结果处理和状态路由都是产品能力。');
fullImageSlide('第二节_AI介入方案/03_异常与人工接管逻辑_v2.png', '用表格总结四类关键异常。强调机器失败不得默认通过。');

// 24 section 3
sectionSlide(3, '搭建规则系统与审核流程', '把审核标准拆成可配置、可执行、可返回、可统计的产品对象。', '35 MIN', '进入项目最核心部分。此处节奏可以稍慢，确保学员理解规则层不等于关键词库。');

// 25 standards
{
  const slide = contentSlide('审核判断来自三层标准', '产品需要把标准拆成规则、场景和版本，而不是塞进一个超长提示词。', '说明三层来源。真实项目必须由业务、法务或审核负责人确认适用范围。');
  const layers = [
    ['外部通用要求', '法律法规与通用监管要求', C.orange, 1.0, 4.8],
    ['渠道与行业规则', '平台规范、行业特殊要求', C.teal2, 1.65, 3.5],
    ['公司内部口径', '品牌、商品、活动与授权标准', C.teal, 2.3, 2.2]
  ];
  layers.forEach((l, i) => {
    slide.addShape(S.roundRect, { x: 2.7 + i * 0.8, y: l[3] + 0.7, w: 7.9 - i * 1.6, h: l[4] - 1.3, rectRadius: 0.08, fill: { color: C.panel, transparency: i * 7 }, line: { color: l[2], width: 1.5 } });
    slide.addText(l[0], { x: 4.2, y: l[3] + 1.15, w: 4.9, h: 0.35, fontSize: 18 - i, bold: true, color: l[2], align: 'center', margin: 0 });
    slide.addText(l[1], { x: 4.0, y: l[3] + 1.62, w: 5.3, h: 0.35, fontSize: 11.5, color: C.muted, align: 'center', margin: 0 });
  });
}

// 26 taxonomy
{
  const slide = contentSlide('把检查项组织成规则分类', '分类是产品设计，不是通用法规清单；真实项目需要确认适用范围。', '从基础质量讲到渠道口径。说明分类的作用是配置、统计和责任分工。');
  const items = [
    ['基础质量', '错别字、病句、信息缺失'], ['禁限表达', '绝对化、保证性、极限性表达'], ['商品事实一致性', '成分、规格、功效、价格'], ['来源与授权', '排名、数据、专家、机构'], ['渠道与内部口径', '活动规则、声明、品牌禁用词']
  ];
  items.forEach((it, i) => {
    const y = 1.7 + i * 0.9;
    slide.addShape(S.roundRect, { x: 0.82, y, w: 11.72, h: 0.7, rectRadius: 0.03, fill: { color: i % 2 ? '0A2730' : C.panel }, line: { color: C.line, width: 0.5 } });
    slide.addText(`0${i + 1}`, { x: 1.05, y: y + 0.18, w: 0.5, h: 0.22, fontFace: 'Arial', fontSize: 12, bold: true, color: i === 3 ? C.orange : C.teal, margin: 0 });
    slide.addText(it[0], { x: 1.75, y: y + 0.18, w: 2.7, h: 0.22, fontSize: 13, bold: true, color: C.white, margin: 0 });
    slide.addText(it[1], { x: 4.75, y: y + 0.18, w: 6.8, h: 0.22, fontSize: 12, color: C.muted, margin: 0 });
  });
}

// 27 atomic rule
{
  const slide = contentSlide('一条原子规则长什么样', '原子规则能单独启停、测试、统计和迭代。', '用排名表达完整性规则讲一遍每个字段。强调业务规则和执行方式要分开。');
  const fields = [
    ['规则编号', 'R-SOURCE-003'], ['规则名称', '排名表达完整性检查'], ['适用场景', '商品详情页、活动页'],
    ['触发条件', '出现第一、领先、TOP1等表达'], ['检查方式', '召回＋语义判断＋材料核对'], ['风险等级', '中；关键材料缺失时升高'],
    ['返回内容', '位置、缺失项、动作、版本'], ['处理动作', '补材料；无法补充则转人工']
  ];
  fields.forEach((f, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.78 + col * 6.1, y = 1.7 + row * 1.12;
    slide.addShape(S.roundRect, { x, y, w: 5.75, h: 0.88, rectRadius: 0.03, fill: { color: C.panel }, line: { color: i === 7 ? C.orange : C.line, width: 0.8 } });
    slide.addText(f[0], { x: x + 0.2, y: y + 0.16, w: 1.2, h: 0.18, fontSize: 9, bold: true, color: i === 7 ? C.orange : C.teal, margin: 0 });
    slide.addText(f[1], { x: x + 0.2, y: y + 0.45, w: 5.25, h: 0.22, fontSize: 11.5, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  });
}

// 28 scenes
{
  const slide = contentSlide('先选场景，再加载规则', '同一条规则不一定适用于所有内容；场景决定输入和启用项。', '对比三个场景。说明第一版只实现前两个，口播稿用于展示扩展性。');
  const data = [
    ['商品详情页', '商品事实\n禁限表达\n来源材料', '商品ID、规格、商品资料', C.teal],
    ['活动页', '价格与时间\n优惠条件\n禁限表达', '活动周期、优惠门槛', C.orange],
    ['短视频口播稿', '口语化夸大\n功效暗示\n授权', '脚本版本、达人主体', C.muted]
  ];
  data.forEach((d, i) => {
    const x = 0.82 + i * 4.1;
    card(slide, x, 1.82, 3.72, 3.98, d[0], `${d[1]}\n\n特殊输入：\n${d[2]}`, d[3], { bodySize: 13.5, line: d[3] });
    pill(slide, x + 0.68, 5.95, 2.35, i < 2 ? '第一版实现' : '后续扩展', d[3], '0B2B34');
  });
}

// 29 methods
{
  const slide = contentSlide('“规则”不等于关键词库', '业务检查项可以调用不同执行方式，规则层与技术执行层要分开。', '解释五种执行方式，举例即可，不深入算法。');
  const methods = [
    ['关键词', '明显表达快速召回'], ['正则／格式', '电话、价格、必填声明'], ['确定性逻辑', '活动时间与字段一致性'], ['大模型语义', '隐含承诺与上下文'], ['外部数据核对', '商品、授权、报告有效期']
  ];
  methods.forEach((m, i) => {
    const x = 0.7 + i * 2.5;
    slide.addShape(S.ellipse, { x: x + 0.62, y: 1.9, w: 0.8, h: 0.8, fill: { color: i === 3 ? C.orange : C.teal }, line: { color: C.bg, width: 2 } });
    slide.addText(String(i + 1), { x: x + 0.84, y: 2.14, w: 0.36, h: 0.18, fontFace: 'Arial', fontSize: 12, bold: true, color: C.ink, align: 'center', margin: 0 });
    card(slide, x, 3.0, 2.1, 2.35, m[0], m[1], i === 3 ? C.orange : C.teal, { titleSize: 13, bodySize: 11.5 });
  });
}

// 30 run case
{
  const slide = contentSlide('让案例真正跑出结构化结果', '每个问题都要能定位到原文，并对应明确处理动作。', '逐行展示四条风险对象。再次提醒这是教学案例中的风险判断，不代替真实法务结论。');
  const rows = [
    ['全网最值得买', '禁限表达', '高', '删除或改为可验证描述'],
    ['7天彻底改善／100%有效', '功效与保证性表达', '高', '删除确定性承诺'],
    ['连续三年销量第一', '来源与范围缺失', '中', '补充范围、时间和来源'],
    ['专家一致推荐', '主体与授权不完整', '高', '补充材料或删除']
  ];
  rows.forEach((r, i) => {
    const y = 1.72 + i * 1.04;
    slide.addShape(S.roundRect, { x: 0.72, y, w: 11.92, h: 0.82, rectRadius: 0.03, fill: { color: i % 2 ? '0A2730' : C.panel }, line: { color: C.line, width: 0.5 } });
    slide.addText(r[0], { x: 0.95, y: y + 0.25, w: 3.0, h: 0.22, fontSize: 12.5, bold: true, color: C.white, margin: 0, fit: 'shrink' });
    slide.addText(r[1], { x: 4.12, y: y + 0.25, w: 2.55, h: 0.22, fontSize: 11.5, color: C.muted, margin: 0, fit: 'shrink' });
    pill(slide, 7.05, y + 0.2, 0.72, r[2], r[2] === '高' ? C.red : C.yellow, '17313A');
    slide.addText(r[3], { x: 8.1, y: y + 0.25, w: 4.15, h: 0.22, fontSize: 11.5, color: C.teal2, margin: 0, fit: 'shrink' });
  });
}

// 31 merge
{
  const slide = contentSlide('同一个问题，先合并再展示', '前台只展示合并后的风险对象，后台保留各能力的原始结果。', '用“100%有效”演示两条来源如何合并。强调原始结果仍需保存，便于分析和追溯。');
  flowNode(slide, 0.85, 2.0, 3.1, 1.38, '关键词规则命中', '位置：18–24\n类型：保证性表达', C.teal);
  flowNode(slide, 0.85, 4.05, 3.1, 1.38, '大模型语义命中', '位置：15–24\n类型：功效承诺', C.orange);
  arrow(slide, 4.08, 2.68, 5.02, 3.55, C.muted);
  arrow(slide, 4.08, 4.72, 5.02, 3.82, C.muted);
  flowNode(slide, 5.15, 2.9, 3.0, 1.62, '合并与定级', '位置重叠＋类型相近\n保留较高等级', C.yellow);
  arrow(slide, 8.28, 3.72, 9.12, 3.72, C.muted);
  flowNode(slide, 9.25, 2.78, 3.15, 1.86, '前台风险对象', '“7天彻底改善，100%有效”\n高风险｜删除确定性承诺', C.red);
}

// 32 states
{
  const slide = contentSlide('整体状态必须对应下一步动作', '结果不是一个分数，而是业务流程中的路由信号。', '逐项讲解状态与下一步。机器失败要作为独立状态，不允许与通过混用。');
  const items = [
    ['机器预审通过', '无有效风险且检查完整', '进入后续业务流程', C.green],
    ['建议修改', '只有低风险问题', '运营修改或按权限继续', C.teal2],
    ['退回修改', '存在明确中高风险', '修改后重新提交', C.orange],
    ['转人工复核', '冲突、低确定性或高风险要求', '进入人工队列', C.red],
    ['机器失败', '超时、格式错误、依赖不可用', '重试或人工接管', C.yellow]
  ];
  items.forEach((it, i) => {
    const y = 1.67 + i * 0.9;
    pill(slide, 0.83, y + 0.11, 2.1, it[0], it[3], '0C2A33');
    slide.addText(it[1], { x: 3.35, y: y + 0.2, w: 4.3, h: 0.2, fontSize: 11.5, color: C.white, margin: 0, fit: 'shrink' });
    arrow(slide, 7.76, y + 0.3, 8.28, y + 0.3, C.muted);
    slide.addText(it[2], { x: 8.58, y: y + 0.2, w: 3.65, h: 0.2, fontSize: 11.5, bold: true, color: it[3], margin: 0, fit: 'shrink' });
  });
}

// 33 negation
{
  const slide = contentSlide('边界样本：关键词命中不等于违规', '召回和最终判断是两件事；边界样本决定系统是否可用。', '先让学员回答，再揭示误报原因和改进方法。');
  slide.addShape(S.roundRect, { x: 0.9, y: 1.95, w: 5.1, h: 3.65, rectRadius: 0.06, fill: { color: C.panel }, line: { color: C.orange, width: 1.2 } });
  slide.addText('“我们不承诺 100% 有效，\n实际体验因人而异。”', { x: 1.32, y: 2.75, w: 4.25, h: 1.1, fontSize: 22, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
  card(slide, 6.45, 1.95, 5.95, 1.5, '关键词系统的问题', '看到“100%有效”就直接判高风险，忽略否定关系。', C.orange, { bodySize: 13 });
  card(slide, 6.45, 3.72, 5.95, 1.88, '改进', '关键词先召回 → 上下文语义确认\n把否定句、引用句和对比句加入回归测试。', C.teal, { bodySize: 13 });
}

fullImageSlide('第三节_规则系统与审核流程/01_审核标准如何变成可配置规则.png', '从左到右解释审核要求如何成为规则对象，再按场景组成不同的规则包。');
fullImageSlide('第三节_规则系统与审核流程/02_一条原子规则的执行流程.png', '以排名表达完整性检查为例，讲清触发、语义确认、材料核对和处理路由。');
fullImageSlide('第三节_规则系统与审核流程/03_多路审核结果的合并与去重.png', '说明关键词、规则和模型如何合并成一个前台风险对象，同时保留后台原始结果。');
fullImageSlide('第三节_规则系统与审核流程/04_营销内容审核状态流转.png', '按状态机走一遍通过、修改、人工复核和机器失败分支，重点讲重新审核回路。');

// 34 section 4
sectionSlide(4, '定义 MVP 与产品结构', '用最少能力完成一次真实闭环，并明确系统中的对象、版本与异常。', '25 MIN', '进入第四节。此处把前面的方案收束成产品范围和系统结构。');

// 35 hypotheses
{
  const slide = contentSlide('MVP 要验证三件事', '第一版验证的是业务闭环，不是“模型会不会看文案”。', '读出三条假设，并说明每条都需要对应样本与业务指标。');
  card(slide, 0.82, 1.9, 3.65, 3.75, '01 发现问题', '混合审核能否稳定发现一批真正有价值的问题？', C.teal, { bodySize: 16 });
  card(slide, 4.84, 1.9, 3.65, 3.75, '02 减少沟通', '结构化结果能否减少运营与审核员的往返确认？', C.teal2, { bodySize: 16 });
  card(slide, 8.86, 1.9, 3.65, 3.75, '03 安全接管', '机器不确定或失败时，人工能否及时并完整地接手？', C.orange, { bodySize: 16 });
}

// 36 scope
{
  const slide = contentSlide('功能优先级：P0、P1、暂不进入', 'MVP 不是少画页面，而是保留完成闭环所必需的能力。', '快速扫过功能。重点解释人工复核、重审和记录为什么属于P0。');
  card(slide, 0.72, 1.72, 4.15, 4.48, 'P0｜形成闭环', '场景与上下文\n文本版本\n规则＋语义检查\n结果合并与定级\n原文定位\n四类处理路径\n人工复核\n修改后重审\n基础记录', C.teal, { bodySize: 12.5 });
  card(slide, 5.08, 1.72, 3.55, 4.48, 'P1｜试点后', '规则运营后台\n批量任务\n数据看板\n更多渠道\n建议一键带入', C.teal2, { bodySize: 13.5 });
  card(slide, 8.84, 1.72, 3.77, 4.48, '暂不进入', '多模态审核\n自动发布\n自动覆盖原文\n全公司复杂权限\n自主 Agent', C.orange, { bodySize: 13.5 });
}

// 37 architecture
{
  const slide = contentSlide('三层产品结构', '架构图的作用，是说明每一层变化会影响哪些产品模块。', '从下到上讲。底层提供能力和数据，中间层把审核逻辑产品化，上层承载业务动作。');
  const layers = [
    ['业务应用层', '提交自查｜流程审核｜人工复核｜历史任务｜规则运营', C.orange, 1.75],
    ['审核能力层', '输入解析｜场景配置｜原子规则｜语义判断｜合并与路由', C.teal, 3.18],
    ['模型与数据来源', '大模型｜商品中心｜活动信息｜授权材料｜当前业务口径', C.teal2, 4.62]
  ];
  layers.forEach((l, i) => {
    slide.addShape(S.roundRect, { x: 1.05, y: l[3], w: 11.2, h: 1.02, rectRadius: 0.04, fill: { color: C.panel }, line: { color: l[2], width: 1.2 } });
    slide.addText(l[0], { x: 1.35, y: l[3] + 0.3, w: 2.1, h: 0.26, fontSize: 15, bold: true, color: l[2], margin: 0 });
    slide.addText(l[1], { x: 3.55, y: l[3] + 0.31, w: 8.2, h: 0.24, fontSize: 12.5, color: C.white, margin: 0, fit: 'shrink' });
  });
  arrow(slide, 11.75, 4.5, 11.75, 4.25, C.muted);
  arrow(slide, 11.75, 3.07, 11.75, 2.82, C.muted);
}

// 38 data objects
{
  const slide = contentSlide('五个核心数据对象', '任务、内容、规则、风险和人工结果需要各自有版本和关联。', '解释对象之间的关系。特别强调内容修改后生成新版本，旧结果不能直接沿用。');
  const nodes = [
    ['审核任务', 'task_id\n场景／状态／提交人'], ['内容版本', 'version_id\n原文／业务上下文'], ['规则版本', 'rule_set_id\n启用项／适用范围'], ['风险对象', '位置／类型／等级／动作'], ['人工结果', '结论／调整原因／处理人']
  ];
  nodes.forEach((n, i) => {
    const x = 0.68 + i * 2.5;
    flowNode(slide, x, 2.15, 2.12, 2.6, n[0], n[1], i === 3 || i === 4 ? C.orange : C.teal);
    if (i < nodes.length - 1) arrow(slide, x + 2.15, 3.45, x + 2.43, 3.45, C.muted);
  });
  slide.addText('内容一旦实质修改 → 生成新版本 → 重新审核', { x: 3.3, y: 5.32, w: 6.7, h: 0.4, fontSize: 17, bold: true, color: C.teal2, align: 'center', margin: 0 });
}

// 39 exceptions/version
{
  const slide = contentSlide('版本与异常，决定产品能否试点', '同一任务中的内容、规则和人工结果都必须有清晰记录。', '讲三个高频异常。引导学员思考如果不做版本，会导致什么混乱。');
  card(slide, 0.82, 1.9, 3.65, 3.75, '材料不完整', '能补齐：提交前补齐\n无法补齐：转人工\n不能让模型自行猜测', C.teal, { bodySize: 14 });
  card(slide, 4.84, 1.9, 3.65, 3.75, '规则正在更新', '每次审核绑定规则版本\n新版本只影响新任务或重审任务', C.teal2, { bodySize: 14 });
  card(slide, 8.86, 1.9, 3.65, 3.75, '人工调整结论', '机器原始结果与人工最终结果同时保存\n记录调整原因', C.orange, { bodySize: 14 });
}

fullImageSlide('第四节_MVP与产品结构/01_智能审核MVP范围.png', '先看三个验证目标，再解释P0、P1和暂不进入之间的边界。');
fullImageSlide('第四节_MVP与产品结构/02_营销内容智能审核三层产品架构.png', '从底层向上讲模型与业务数据、审核能力和业务应用之间的依赖关系。');
fullImageSlide('第四节_MVP与产品结构/03_智能审核核心数据对象关系.png', '重点解释一个任务如何关联多个内容版本，以及每个版本分别产生风险对象。');

// 40 section 5
sectionSlide(5, '设计四个核心页面', '页面不是展示 AI 回复，而是帮助用户定位、理解、处理与复核。', '20 MIN', '进入产品形式和交互部分。用四张可编辑线框图串起完整使用流程。');

function browserFrame(slide, x, y, w, h, title) {
  slide.addShape(S.roundRect, { x, y, w, h, rectRadius: 0.05, fill: { color: 'F4F7F8' }, line: { color: '6C8992', width: 0.8 } });
  slide.addShape(S.rect, { x, y, w, h: 0.42, fill: { color: 'DDE7EA' }, line: { color: 'DDE7EA' } });
  ['F27867', 'F5C85B', '55C6A4'].forEach((c, i) => slide.addShape(S.ellipse, { x: x + 0.15 + i * 0.22, y: y + 0.14, w: 0.1, h: 0.1, fill: { color: c }, line: { color: c } }));
  slide.addText(title, { x: x + 0.7, y: y + 0.12, w: w - 1.4, h: 0.14, fontSize: 7.5, bold: true, color: '31505A', align: 'center', margin: 0 });
}

// 41 list UI
{
  const slide = contentSlide('页面 1｜审核任务列表', '不同角色看到同一任务，但排序、筛选和待办重点不同。', '介绍列表页信息与角色差异。机器失败必须作为独立状态可筛选。');
  browserFrame(slide, 0.72, 1.72, 8.2, 4.45, '营销内容审核平台 · 任务列表');
  slide.addShape(S.rect, { x: 0.9, y: 2.36, w: 1.25, h: 3.55, fill: { color: '153944' }, line: { color: '153944' } });
  ['任务列表', '待我处理', '规则管理', '数据看板'].forEach((t, i) => slide.addText(t, { x: 1.03, y: 2.72 + i * 0.55, w: 0.95, h: 0.16, fontSize: 7.5, bold: i === 0, color: i === 0 ? C.teal2 : 'B8CAD0', margin: 0 }));
  slide.addShape(S.roundRect, { x: 2.38, y: 2.58, w: 5.9, h: 0.45, rectRadius: 0.02, fill: { color: 'E8EEF0' }, line: { color: 'CBD9DD' } });
  ['任务名称', '场景', '风险', '状态', '待处理人'].forEach((t, i) => slide.addText(t, { x: 2.55 + i * 1.12, y: 2.74, w: 0.9, h: 0.12, fontSize: 6.8, bold: true, color: '31505A', margin: 0 }));
  for (let r = 0; r < 4; r++) {
    slide.addShape(S.rect, { x: 2.38, y: 3.16 + r * 0.6, w: 5.9, h: 0.48, fill: { color: r % 2 ? 'FFFFFF' : 'F7FAFB' }, line: { color: 'D8E2E5', width: 0.4 } });
    slide.addText(`活动文案审核-${104 + r}`, { x: 2.55, y: 3.33 + r * 0.6, w: 1.35, h: 0.12, fontSize: 6.7, color: '294953', margin: 0 });
    slide.addText(r % 2 ? '活动页' : '商品页', { x: 3.74, y: 3.33 + r * 0.6, w: 0.7, h: 0.12, fontSize: 6.7, color: '52717B', margin: 0 });
    slide.addText(r === 1 ? '3 高' : r === 2 ? '失败' : '1 中', { x: 4.88, y: 3.33 + r * 0.6, w: 0.6, h: 0.12, fontSize: 6.7, bold: true, color: r === 1 ? 'DA5B43' : 'C4962C', margin: 0 });
    slide.addText(r === 1 ? '待人工' : r === 2 ? '机器失败' : '待修改', { x: 5.92, y: 3.33 + r * 0.6, w: 0.8, h: 0.12, fontSize: 6.7, color: '294953', margin: 0 });
  }
  card(slide, 9.25, 1.72, 3.35, 1.3, '列表必须看懂', '场景、版本、风险、状态、待处理人', C.teal, { bodySize: 11.5 });
  card(slide, 9.25, 3.28, 3.35, 1.3, '角色默认排序', '运营看自己的任务；审核员优先看高风险与超时任务', C.orange, { bodySize: 11.5 });
  card(slide, 9.25, 4.84, 3.35, 1.3, '失败不能伪装', '机器失败不能混进“通过”结果', C.yellow, { bodySize: 11.5 });
}

// 42 submit UI
{
  const slide = contentSlide('页面 2｜内容提交页', '场景决定必填上下文；提交前先做完整性检查。', '说明左侧业务上下文、右侧正文。切换场景后字段应动态变化。');
  browserFrame(slide, 0.72, 1.72, 8.2, 4.45, '新建审核任务');
  slide.addText('业务上下文', { x: 1.0, y: 2.42, w: 1.4, h: 0.18, fontSize: 8, bold: true, color: '294953', margin: 0 });
  const labels = ['审核场景 *', '发布渠道 *', '商品对象 *', '活动周期', '授权材料'];
  labels.forEach((l, i) => {
    slide.addText(l, { x: 1.0, y: 2.78 + i * 0.6, w: 1.0, h: 0.12, fontSize: 6.5, color: '52717B', margin: 0 });
    slide.addShape(S.roundRect, { x: 2.05, y: 2.68 + i * 0.6, w: 2.15, h: 0.32, rectRadius: 0.02, fill: { color: 'FFFFFF' }, line: { color: 'C3D2D7', width: 0.5 } });
  });
  slide.addText('营销正文 *', { x: 4.58, y: 2.42, w: 1.25, h: 0.18, fontSize: 8, bold: true, color: '294953', margin: 0 });
  slide.addShape(S.roundRect, { x: 4.58, y: 2.7, w: 3.95, h: 2.65, rectRadius: 0.03, fill: { color: 'FFFFFF' }, line: { color: 'C3D2D7', width: 0.5 } });
  slide.addText('全网最值得买的胶原蛋白饮……', { x: 4.83, y: 2.98, w: 3.45, h: 1.2, fontSize: 9, color: '31505A', margin: 0 });
  slide.addShape(S.roundRect, { x: 7.15, y: 5.52, w: 1.38, h: 0.38, rectRadius: 0.04, fill: { color: '1EB8A7' }, line: { color: '1EB8A7' } });
  slide.addText('提交预审', { x: 7.42, y: 5.64, w: 0.83, h: 0.13, fontSize: 7, bold: true, color: 'FFFFFF', align: 'center', margin: 0 });
  card(slide, 9.25, 1.72, 3.35, 1.3, '动态字段', '场景变化后，需要的商品、活动和授权信息随之变化', C.teal, { bodySize: 11.5 });
  card(slide, 9.25, 3.28, 3.35, 1.3, '提交前校验', '缺少必要上下文时先补齐，不把猜测交给模型', C.orange, { bodySize: 11.5 });
  card(slide, 9.25, 4.84, 3.35, 1.3, '生成版本', '提交后形成不可静默覆盖的内容版本', C.teal2, { bodySize: 11.5 });
}

// 43 workbench UI
{
  const slide = contentSlide('页面 3｜智能审核工作台', '左侧原文、右侧风险，点击任一侧都能定位到另一侧。', '这是整套产品最重要的页面。先讲布局，再讲双向定位、版本变化和建议使用方式。');
  browserFrame(slide, 0.58, 1.65, 9.3, 4.62, '任务 MKT-0108 · 审核工作台');
  slide.addText('机器预审：退回修改', { x: 0.92, y: 2.31, w: 2.0, h: 0.18, fontSize: 8, bold: true, color: 'D65C44', margin: 0 });
  slide.addText('高 3  ·  中 1  ·  规则 v1.4', { x: 3.05, y: 2.31, w: 2.2, h: 0.18, fontSize: 7.2, color: '52717B', margin: 0 });
  slide.addShape(S.roundRect, { x: 0.88, y: 2.72, w: 5.2, h: 3.02, rectRadius: 0.03, fill: { color: 'FFFFFF' }, line: { color: 'C3D2D7', width: 0.5 } });
  slide.addText('全网最值得买的胶原蛋白饮，7 天彻底改善熬夜暗沉，100% 有效。连续三年销量第一，三甲医院专家一致推荐。', { x: 1.18, y: 3.08, w: 4.6, h: 1.55, fontSize: 11, color: '294953', margin: 0, breakLine: false });
  [[1.14, 2.99, 1.3], [3.14, 3.39, 2.1], [1.12, 4.08, 2.15]].forEach((b, i) => slide.addShape(S.rect, { x: b[0], y: b[1], w: b[2], h: 0.24, fill: { color: i === 1 ? 'FFE1D9' : 'FFF1C7', transparency: 10 }, line: { color: i === 1 ? 'E4775D' : 'D7A43E', transparency: 35 } }));
  const risks = [
    ['高', '保证性与功效表达', '删除确定性承诺'], ['高', '禁限表达', '改为可验证描述'], ['中', '来源与范围缺失', '补充时间、范围和来源']
  ];
  risks.forEach((r, i) => {
    const y = 2.72 + i * 0.98;
    slide.addShape(S.roundRect, { x: 6.35, y, w: 3.1, h: 0.78, rectRadius: 0.03, fill: { color: i === 0 ? 'FFF0EB' : 'F6FAFB' }, line: { color: i === 0 ? 'E4775D' : 'C3D2D7', width: 0.6 } });
    slide.addText(r[0], { x: 6.55, y: y + 0.16, w: 0.34, h: 0.15, fontSize: 7, bold: true, color: 'D85F47', margin: 0 });
    slide.addText(r[1], { x: 7.0, y: y + 0.14, w: 1.95, h: 0.15, fontSize: 7.2, bold: true, color: '294953', margin: 0 });
    slide.addText(r[2], { x: 7.0, y: y + 0.42, w: 2.05, h: 0.14, fontSize: 6.4, color: '52717B', margin: 0 });
  });
  card(slide, 10.2, 1.65, 2.55, 1.32, '双向定位', '点问题 → 高亮原文\n点原文 → 定位问题', C.teal, { bodySize: 10.5 });
  card(slide, 10.2, 3.2, 2.55, 1.32, '内容变化', '正文修改后提示：需要重新审核', C.orange, { bodySize: 10.5 });
  card(slide, 10.2, 4.75, 2.55, 1.52, '建议不静默覆盖', '可以复制或带入编辑区，但保留用户确认', C.teal2, { bodySize: 10.5 });
}

// 44 human review UI
{
  const slide = contentSlide('页面 4｜人工复核页', '审核员需要看到机器如何判断，也需要能调整并留下原因。', '说明人工页与运营页的差别：信息更多、动作更重、调整需要记录。');
  browserFrame(slide, 0.72, 1.72, 8.2, 4.45, '待复核任务 MKT-0108');
  slide.addShape(S.roundRect, { x: 1.02, y: 2.48, w: 3.45, h: 2.85, rectRadius: 0.03, fill: { color: 'FFFFFF' }, line: { color: 'C3D2D7', width: 0.5 } });
  slide.addText('原文与业务上下文', { x: 1.28, y: 2.76, w: 2.0, h: 0.17, fontSize: 8, bold: true, color: '294953', margin: 0 });
  slide.addText('场景：活动页\n商品：胶原蛋白饮\n渠道：自有商城\n当前版本：V2\n相关材料：销量报告、授权文件', { x: 1.28, y: 3.13, w: 2.7, h: 1.65, fontSize: 7.5, color: '52717B', margin: 0, breakLine: false });
  slide.addShape(S.roundRect, { x: 4.72, y: 2.48, w: 3.75, h: 2.85, rectRadius: 0.03, fill: { color: 'FFFFFF' }, line: { color: 'C3D2D7', width: 0.5 } });
  slide.addText('机器结果与人工处理', { x: 5.0, y: 2.76, w: 2.2, h: 0.17, fontSize: 8, bold: true, color: '294953', margin: 0 });
  ['通过', '驳回修改', '调整结果', '升级处理'].forEach((t, i) => {
    slide.addShape(S.roundRect, { x: 5.03 + (i % 2) * 1.55, y: 3.23 + Math.floor(i / 2) * 0.72, w: 1.32, h: 0.4, rectRadius: 0.03, fill: { color: i === 1 ? 'E66B50' : 'E8EEF0' }, line: { color: i === 1 ? 'E66B50' : 'C3D2D7' } });
    slide.addText(t, { x: 5.2 + (i % 2) * 1.55, y: 3.36 + Math.floor(i / 2) * 0.72, w: 0.98, h: 0.13, fontSize: 7, bold: true, color: i === 1 ? 'FFFFFF' : '31505A', align: 'center', margin: 0 });
  });
  slide.addText('调整原因 *', { x: 5.03, y: 4.72, w: 0.9, h: 0.13, fontSize: 6.5, color: '52717B', margin: 0 });
  slide.addShape(S.roundRect, { x: 5.92, y: 4.62, w: 2.05, h: 0.32, rectRadius: 0.02, fill: { color: 'FFFFFF' }, line: { color: 'C3D2D7', width: 0.5 } });
  card(slide, 9.25, 1.72, 3.35, 1.3, '审核员多看什么', '机器原始结果、业务资料、规则信息、历史修改', C.teal, { bodySize: 11 });
  card(slide, 9.25, 3.28, 3.35, 1.3, '审核员能做什么', '通过、驳回、调整、补充意见、升级', C.orange, { bodySize: 11 });
  card(slide, 9.25, 4.84, 3.35, 1.3, '为什么记录原因', '用于规则优化、问题归类和回归测试', C.teal2, { bodySize: 11 });
}

// 45 sequence
{
  const slide = contentSlide('作品集里，页面要放在流程后面', '先说明为什么这样设计，再展示页面形式，产品判断才有说服力。', '这是作品集表达提醒。不要只堆高保真图，要让页面对应前面的业务决策。');
  const nodes = ['业务流程', '产品主流程', '任务状态', '核心页面', '完整案例流转'];
  nodes.forEach((n, i) => {
    const x = 0.72 + i * 2.5;
    flowNode(slide, x, 2.55, 2.08, 1.6, n, i === 3 ? '页面形式' : '说明设计原因', i === 3 ? C.orange : C.teal);
    if (i < nodes.length - 1) arrow(slide, x + 2.1, 3.35, x + 2.43, 3.35, C.muted);
  });
  slide.addText('页面不是孤立作品，而是前面所有判断的可视化结果。', { x: 2.35, y: 4.92, w: 8.65, h: 0.4, fontSize: 19, bold: true, color: C.teal2, align: 'center', margin: 0 });
}

fullImageSlide('第五节_核心页面原型/01_审核任务列表页.png', '任务列表页先看角色的任务入口、状态和处理优先级。机器失败必须作为独立状态出现。');
fullImageSlide('第五节_核心页面原型/02_内容提交页.png', '提交页左侧收集业务上下文，右侧输入正文。提交前先做完整性检查。');
fullImageSlide('第五节_核心页面原型/03_智能审核工作台.png', '重点讲原文与风险列表的双向定位、风险卡片信息和修改后重新审核。');
fullImageSlide('第五节_核心页面原型/04_人工复核页.png', '人工页承担最终判断，审核员可以调整等级、选择原因并提交处理结论。');

// 46 section 6
sectionSlide(6, '评测、坏案例与项目包装', '用样本和真实失败推动迭代，并把项目整理成可追问的作品集。', '15 MIN', '最后一节节奏紧凑。重点讲评测字段、三类坏案例和作业。');

// 47 dataset
{
  const slide = contentSlide('先建立一套可用的评测样本', '没有统一标注的样本，就无法知道系统真的变好还是只是“看起来不错”。', '介绍样本来源与字段。强调标注人和规则版本也需要保存。');
  card(slide, 0.75, 1.78, 4.05, 4.38, '样本来源', '脱敏历史内容\n审核员典型案例\n规则边界样本\n已上线内容抽检', C.teal, { bodySize: 16 });
  const fields = ['内容与场景', '必要上下文', '问题位置', '问题类型与等级', '期望动作', '标注人及规则版本'];
  fields.forEach((f, i) => {
    const x = 5.25 + (i % 2) * 3.55, y = 1.78 + Math.floor(i / 2) * 1.4;
    card(slide, x, y, 3.22, 1.05, f, '', i === 4 ? C.orange : C.teal2, { titleSize: 12.5 });
  });
}

// 48 metrics
{
  const slide = contentSlide('指标要覆盖效果、稳定和业务', '总准确率不够；高风险漏检和人工接管同样重要。', '讲四组指标。示例门槛只能作为课堂讨论值，不能照搬到真实项目。');
  card(slide, 0.75, 1.82, 2.82, 3.85, '识别效果', '按问题类型看\n准确率\n召回率\nF1', C.teal, { bodySize: 14.5 });
  card(slide, 3.78, 1.82, 2.82, 3.85, '工程稳定', '结构化输出成功率\n超时率\n依赖失败率', C.teal2, { bodySize: 14.5 });
  card(slide, 6.81, 1.82, 2.82, 3.85, '业务处理', '人工复核比例\n平均处理时长\n一次修改通过率', C.yellow, { bodySize: 14.5 });
  card(slide, 9.84, 1.82, 2.82, 3.85, '高风险表现', '严重漏检数量\n人工调整比例\n高风险响应时间', C.orange, { bodySize: 14.5 });
}

// 49 bad cases
{
  const slide = contentSlide('三个坏案例，比一张高分截图更有价值', '坏案例必须落到具体修改动作，并进入回归测试。', '三列快速讲解。让学员看见项目迭代不是“再调一下提示词”。');
  card(slide, 0.75, 1.78, 3.78, 4.45, '01 否定句误报', '输入：不承诺100%有效\n\n原因：关键词直接定性\n\n改进：语义确认＋否定句样本', C.teal, { bodySize: 13.5 });
  card(slide, 4.78, 1.78, 3.78, 4.45, '02 上下文不足', '输入：连续三年销量第一\n\n原因：缺报告仍强行判断\n\n改进：补材料或转人工', C.orange, { bodySize: 13.5 });
  card(slide, 8.81, 1.78, 3.78, 4.45, '03 重复展示', '输入：全网第一，排名第一\n\n原因：多能力重复命中\n\n改进：按位置和类型合并', C.teal2, { bodySize: 13.5 });
}

// 50 iteration
{
  const slide = contentSlide('迭代闭环：从线上问题回到规则版本', '每次修改都要经过归类、改动、回归和发布，而不是直接线上试错。', '顺时针讲解闭环。可以追问学员：某个误报应修改规则、提示、业务资料还是页面？');
  const nodes = [
    ['线上任务', '真实输入'], ['人机结果对比', '发现差异'], ['问题归类', '误报／漏报／分级／定位'], ['决定修改', '规则／提示／资料／交互'], ['回归测试', '旧问题不能复发'], ['发布新版本', '记录适用范围']
  ];
  const pos = [[1.1,2.0],[4.0,1.55],[7.05,1.55],[9.85,2.0],[7.15,4.45],[3.85,4.45]];
  nodes.forEach((n,i)=>flowNode(slide,pos[i][0],pos[i][1],2.35,1.15,n[0],n[1],i===3?C.orange:C.teal));
  for(let i=0;i<5;i++) arrow(slide,pos[i][0]+2.38,pos[i][1]+0.58,pos[i+1][0]-0.12,pos[i+1][1]+0.58,C.muted);
  arrow(slide,3.75,5.04,1.8,3.35,C.muted);
}

// 51 portfolio
{
  const slide = contentSlide('作品集用七部分讲清项目', '专业性来自前后判断能对应，而不是页面数量。', '快速说明作品集结构。建议学员每一部分都使用本课同一个案例，避免散。');
  const items = [
    ['01','业务背景'],['02','项目边界'],['03','方案选择'],['04','产品设计'],['05','AI设计'],['06','测试迭代'],['07','项目复盘']
  ];
  items.forEach((it,i)=>{
    const x=0.65+i*1.78;
    slide.addShape(S.roundRect,{x,y:2.35,w:1.5,h:2.15,rectRadius:0.05,fill:{color:i===3?C.panel2:C.panel},line:{color:i===3?C.orange:C.line,width:i===3?1.5:0.8}});
    slide.addText(it[0],{x:x+0.43,y:2.72,w:0.64,h:0.25,fontFace:'Arial',fontSize:16,bold:true,color:i===3?C.orange:C.teal,align:'center',margin:0});
    slide.addText(it[1],{x:x+0.15,y:3.45,w:1.2,h:0.55,fontSize:12,bold:true,color:C.white,align:'center',valign:'mid',margin:0,fit:'shrink'});
  });
  slide.addText('每一页都回答：为什么这样做？它解决了哪个流程问题？如何验证？', { x: 1.7, y: 5.15, w: 9.95, h: 0.42, fontSize: 18, bold: true, color: C.teal2, align: 'center', margin: 0 });
}

fullImageSlide('第六节_评测与项目包装/01_智能审核评测与迭代闭环.png', '从建立样本开始，依次讲运行评测、判断门槛、发布版本和线上观察形成的循环。');
fullImageSlide('第六节_评测与项目包装/02_坏案例分析与问题定位.png', '坏案例先分类，再定位具体模块。不要把所有问题都归结为模型或提示词。');
fullImageSlide('第六节_评测与项目包装/03_作品集项目叙事结构.png', '按照七部分回顾项目表达结构，并用一条完整任务案例串联方案、产品和评测。');

// 52 pitch
{
  const slide = contentSlide('面试时的一分钟讲法', '先讲问题和边界，再讲方案、产品闭环与验证，不从模型名称开始。', '老师可完整读一遍，也可以让学员先用自己的话复述，再对照结构。');
  slide.addShape(S.roundRect, { x: 0.9, y: 1.82, w: 11.52, h: 4.35, rectRadius: 0.07, fill: { color: C.panel }, line: { color: C.teal, width: 1 } });
  slide.addText('我做的是一个面向电商营销文本的智能审核平台。原流程依赖人工逐篇检查，高峰期等待时间长，审核口径也不够稳定。第一版我限定商品详情页与活动文案两个场景，采用确定性规则、大模型语义判断、业务资料核对和人工复核的混合流程。产品核心不是返回一段 AI 结论，而是把每个问题定位到原文，给出类型、等级和处理动作，并支持修改后重审与人工接管。评测除了准确率和召回率，也会关注高风险漏检、人工复核比例和处理时长，再通过坏案例持续更新规则和测试集。', { x: 1.35, y: 2.36, w: 10.62, h: 3.2, fontSize: 17, color: C.white, breakLine: false, valign: 'mid', margin: 0.05, fit: 'shrink' });
}

// 53 assignment
{
  const slide = contentSlide('课后作业｜完成一个项目包', '作业重点不是做完整系统，而是让六份材料前后对应。', '说明必做六项与进阶迁移。下一次可以用十条样本和工作台原型做课堂点评。');
  card(slide, 0.78, 1.72, 7.0, 4.52, '必做', '1. 一页业务现状与项目边界\n2. 一张混合审核主流程图\n3. 三条原子规则\n4. 一张智能审核工作台原型\n5. 十条测试样本，至少三条边界样本\n6. 一页坏案例与改进方案', C.teal, { bodySize: 14.5 });
  card(slide, 8.05, 1.72, 4.52, 4.52, '进阶', '迁移到一个新行业：\n\n招聘信息审核\n教育宣传审核\n金融营销材料审核\n\n必须重新定义角色、场景、规则分类、必要上下文和人工接管条件。', C.orange, { bodySize: 13.5 });
}

// 54 close
{
  const slide = pptx.addSlide();
  addBg(slide);
  slide.addText('PROJECT COMPLETE', { x: 0.75, y: 0.72, w: 4.5, h: 0.25, fontSize: 10, bold: true, color: C.teal, charSpacing: 1.6, margin: 0 });
  slide.addText('最后记住三句话', { x: 0.75, y: 1.42, w: 6.8, h: 0.7, fontSize: 34, bold: true, color: C.white, margin: 0 });
  const lines = [
    ['01', '智能审核不是一个模型调用，而是一条有人机分工的业务流程。'],
    ['02', '规则、模型、业务资料和人工各自负责不同问题。'],
    ['03', '可落地项目必须讲清输入、判断、处理、复核、失败和迭代。']
  ];
  lines.forEach((l, i) => {
    const y = 2.55 + i * 1.15;
    slide.addText(l[0], { x: 0.82, y: y + 0.12, w: 0.65, h: 0.3, fontFace: 'Arial', fontSize: 18, bold: true, color: i === 2 ? C.orange : C.teal, margin: 0 });
    slide.addShape(S.line, { x: 1.58, y: y + 0.27, w: 0.72, h: 0, line: { color: C.line, width: 1 } });
    slide.addText(l[1], { x: 2.62, y, w: 9.55, h: 0.55, fontSize: 18, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  });
  slide.addText('下一步：把课堂方案做成你的项目材料', { x: 0.78, y: 6.36, w: 5.8, h: 0.32, fontSize: 14, bold: true, color: C.teal2, margin: 0 });
  addFooter(slide, 'AI产品经理课程 · 第二阶段');
  slide.addNotes('最后只强调三句话，然后布置作业。不要再延伸成面试理论课，保持项目收口。');
}

const out = path.resolve(__dirname, '../02_PPT_正式版/第四课_AI营销内容智能审核项目实战_配图完整版.pptx');
pptx.writeFile({ fileName: out });
