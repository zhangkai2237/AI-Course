const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第02课：大模型基础、上下文与Prompt产品化设计';
pptx.title = '大模型基础、上下文与Prompt产品化设计';
pptx.company = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };

const S = pptx.ShapeType;
const OUT = '课程生产/第一阶段/第02课_大模型基础、上下文与Prompt产品化设计_重做工作区/10_正式PPT_第02课_大模型基础、上下文与Prompt产品化设计.pptx';
const ROOT = '课程生产/第一阶段/第02课_大模型基础、上下文与Prompt产品化设计_重做工作区';

const C = {
  bg: 'F4F0E8', paper: 'FBF9F4', ink: '20211F', text: '292A27', muted: '77766F', line: 'D8D0C2',
  red: 'C35331', teal: '387B77', purple: '74546A', orange: 'D18A38', green: '527553', blue: '405D78',
  softRed: 'F2E2DB', softTeal: 'DDEBE8', softPurple: 'E8E0E7', softOrange: 'F2E7D5', softGreen: 'E2E9DE', softBlue: 'DDE5EB',
  white: 'FFFFFF', dark: '1D1E1C', pale: 'ECE7DE'
};

const moduleOf = n => n <= 3 ? '课程任务' : n <= 6 ? '01 会聊天不等于会做产品' : n <= 13 ? '02 大模型怎样生成答案' : n <= 18 ? '03 Token与上下文窗口' : n <= 23 ? '04 从目标输出反推上下文' : n <= 29 ? '05 结构化Prompt六要素' : n === 30 ? '短休息' : n <= 35 ? '06 从Prompt到产品' : n <= 44 ? '07 销售拜访记录助手案例' : n <= 46 ? '08 Bad Case与Prompt边界' : '09 总结与课后作业';
const kickerOf = n => moduleOf(n).replace(/^\d+\s*/, '');

function t(slide, text, x, y, w, h, o = {}) {
  slide.addText(text, { x, y, w, h, margin: 0, fontFace: 'PingFang SC', fontSize: 13, color: C.text,
    bold: false, align: 'left', valign: 'top', breakLine: false, fit: 'shrink', ...o });
}
function shape(slide, type, x, y, w, h, fill, line = fill, extra = {}) {
  slide.addShape(type, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 1 }, ...extra });
}
function rule(slide, x, y, w, color = C.line, width = 1) {
  slide.addShape(S.line, { x, y, w, h: 0, line: { color, width } });
}
function pill(slide, text, x, y, w, color = C.red, fill = color, textColor = C.white) {
  shape(slide, S.roundRect, x, y, w, 0.34, fill, fill, { rectRadius: 0.1 });
  t(slide, text, x + 0.07, y + 0.075, w - 0.14, 0.17, { fontSize: 9, bold: true, color: textColor, align: 'center' });
}
function base(n, title, kicker = kickerOf(n), opts = {}) {
  const s = pptx.addSlide();
  s.background = { color: opts.bg || C.bg };
  if (!opts.noHeader) {
    pill(s, String(n).padStart(2, '0'), 0.62, 0.3, 0.62, C.red);
    t(s, kicker, 1.42, 0.38, 5.4, 0.18, { fontSize: 9.5, bold: true, color: C.muted, charSpacing: 1.05 });
    t(s, title, 0.62, 0.78, 12.0, 0.52, { fontSize: 25, bold: true, color: C.ink });
    rule(s, 0.62, 1.42, 12.05);
  }
  if (!opts.noFooter) {
    t(s, 'AI 产品经理系统课程 · 第一阶段', 0.62, 7.06, 4.3, 0.16, { fontSize: 8, color: C.muted });
    t(s, moduleOf(n), 7.4, 7.06, 4.5, 0.16, { fontSize: 8, color: C.muted, align: 'right' });
    t(s, String(n).padStart(2, '0'), 12.1, 7.04, 0.55, 0.18, { fontSize: 8.5, color: C.muted, align: 'right' });
  }
  return s;
}
function card(slide, x, y, w, h, head, body, accent = C.red, o = {}) {
  shape(slide, S.roundRect, x, y, w, h, o.fill || C.paper, o.line || C.line, { rectRadius: 0.05,
    shadow: o.shadow === false ? undefined : { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: 0.07 } });
  shape(slide, S.rect, x, y, o.top ? w : 0.07, o.top ? 0.08 : h, accent, accent);
  t(slide, head, x + 0.27, y + 0.23, w - 0.52, 0.38, { fontSize: o.headSize || 15, bold: true, color: o.headColor || C.ink, align: o.headAlign || 'left' });
  if (body) t(slide, body, x + 0.27, y + 0.79, w - 0.54, h - 1.02, { fontSize: o.bodySize || 11.4, color: o.bodyColor || C.muted, valign: o.valign || 'top', paraSpaceAfterPt: 5, align: o.bodyAlign || 'left' });
}
function label(slide, text, x, y, w, color = C.red, fill = C.softRed) {
  shape(slide, S.roundRect, x, y, w, 0.38, fill, fill, { rectRadius: 0.08 });
  t(slide, text, x + 0.08, y + 0.09, w - 0.16, 0.17, { fontSize: 9.5, bold: true, color, align: 'center' });
}
function arrow(slide, x, y, w = 0.42, color = C.line) { shape(slide, S.chevron, x, y, w, 0.46, color, color); }
function quote(slide, text, x, y, w, h, accent = C.red, size = 22, align = 'left') {
  shape(slide, S.rect, x, y, 0.08, h, accent, accent);
  t(slide, text, x + 0.35, y + 0.1, w - 0.35, h - 0.2, { fontSize: size, bold: true, color: C.ink, valign: 'mid', align });
}
function flow(slide, items, y = 2.0, x = 0.62, totalW = 12.0, accents = []) {
  const gap = 0.22, w = (totalW - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const xx = x + i * (w + gap), a = accents[i] || [C.red, C.orange, C.teal, C.purple, C.green][i % 5];
    shape(slide, S.roundRect, xx, y, w, 1.7, C.paper, C.line, { rectRadius: 0.05 });
    t(slide, String(i + 1).padStart(2, '0'), xx + 0.18, y + 0.19, 0.48, 0.2, { fontSize: 10, bold: true, color: a });
    t(slide, it[0], xx + 0.15, y + 0.62, w - 0.3, 0.36, { fontSize: it[2] || 13.2, bold: true, color: C.ink, align: 'center' });
    if (it[1]) t(slide, it[1], xx + 0.17, y + 1.13, w - 0.34, 0.3, { fontSize: 9.7, color: C.muted, align: 'center' });
  });
}
function grid(slide, items, cols = 3, y = 1.7, h = 1.75, colors = [C.red, C.teal, C.orange, C.purple, C.green, C.blue]) {
  const x0 = 0.72, gap = 0.24, w = (11.9 - gap * (cols - 1)) / cols;
  items.forEach((it, i) => {
    const x = x0 + (i % cols) * (w + gap), yy = y + Math.floor(i / cols) * (h + 0.24);
    card(slide, x, yy, w, h, it[0], it[1], colors[i % colors.length], { shadow: false, headSize: it[2] || 15, bodySize: it[3] || 11.2, bodyAlign: it[4] || 'left' });
  });
}
function row(slide, idx, head, body, y, accent = C.red, x = 0.82, w = 11.5) {
  t(slide, String(idx).padStart(2, '0'), x, y + 0.04, 0.55, 0.25, { fontSize: 11.5, bold: true, color: accent });
  t(slide, head, x + 0.82, y, 2.25, 0.34, { fontSize: 14.5, bold: true, color: C.ink });
  t(slide, body, x + 3.35, y + 0.02, w - 3.35, 0.36, { fontSize: 11.8, color: C.muted });
  rule(slide, x, y + 0.6, w);
}
function note(slide, n, extra = '') {
  slide.addNotes(`第${n}页｜${moduleOf(n)}\n逐页内容：${ROOT}/09_PPT结构与逐页内容\n老师讲解：${ROOT}/05_老师版讲义\n${extra}`);
}
function twoCol(slide, left, right, y = 1.72, h = 4.65, leftAccent = C.teal, rightAccent = C.orange) {
  card(slide, 0.78, y, 5.78, h, left[0], left[1], leftAccent, { headSize: left[2] || 18, bodySize: left[3] || 13 });
  card(slide, 6.78, y, 5.78, h, right[0], right[1], rightAccent, { headSize: right[2] || 18, bodySize: right[3] || 13 });
}

// 01 封面
{
  const s = base(1, '', '', { noHeader: true, noFooter: true });
  shape(s, S.rect, 0, 0, 0.22, 7.5, C.red, C.red);
  t(s, '02', 0.88, 0.72, 1.1, 0.66, { fontSize: 35, bold: true, color: C.red });
  t(s, '大模型基础、上下文与\nPrompt产品化设计', 0.88, 1.6, 8.6, 1.55, { fontSize: 36, bold: true, color: C.ink });
  t(s, '从一次好回答，到可重复交付的AI产品', 0.92, 3.53, 8.3, 0.42, { fontSize: 17, color: C.muted });
  rule(s, 0.92, 4.18, 3.0, C.red, 3);
  const xs = [8.2, 9.18, 10.16, 11.14];
  ['输入', '上下文', '生成', '产品'].forEach((x, i) => { label(s, x, xs[i], 1.3 + i * 0.62, 1.35, [C.red, C.teal, C.purple, C.orange][i], [C.softRed, C.softTeal, C.softPurple, C.softOrange][i]); if (i < 3) arrow(s, xs[i] + 0.45, 1.82 + i * 0.62, 0.35); });
  pill(s, '第一阶段 · 第02课', 9.66, 6.43, 2.55, C.red);
  note(s, 1, '开场只说课程定位，不展开技术名词。');
}

// 02 四个问题
{
  const s = base(2, '今天回答四个问题');
  flow(s, [['怎样生成？', '逐步生成'], ['能看到什么？', '上下文窗口'], ['Prompt怎样写？', '产品定义'], ['怎样变成产品？', '交互与风险']], 1.9, 0.62, 12.0, [C.red, C.teal, C.purple, C.orange]);
  t(s, '生成  →  上下文  →  结构化Prompt  →  产品化', 1.1, 5.15, 11.1, 0.45, { fontSize: 19, bold: true, color: C.red, align: 'center' });
  note(s, 2);
}

// 03 课堂产出
{
  const s = base(3, '今天不是只听懂，而是完成一份设计');
  flow(s, [['V0.1', '输出＋验收＋上下文'], ['V0.2', 'Prompt关键骨架'], ['V0.3', '信息供给＋风险处理'], ['案例验证', '找错＋修改＋复测']], 1.82, 0.62, 12.0, [C.blue, C.purple, C.teal, C.orange]);
  quote(s, '课堂案例｜销售拜访记录助手', 1.0, 4.35, 7.6, 0.8, C.red, 20);
  label(s, '教学虚构', 9.45, 4.5, 1.55, C.orange, C.softOrange);
  t(s, '公司、人物、产品和数据均为虚构', 8.2, 5.18, 3.7, 0.28, { fontSize: 10.5, color: C.muted, align: 'right' });
  note(s, 3, '基准版只完成V0.2关键骨架。');
}

// 04A 互动起始：先投票，不提前展示结果
{
  const s = base(4, '同一句话，运行两次，结果会一样吗？');
  card(s, 0.78, 1.68, 5.0, 4.75, '输入完全相同', '请为一款「销售拜访记录助手」\n写一句15字以内的产品介绍。', C.red, { headSize: 18, bodySize: 19, valign: 'mid', bodyAlign: 'center' });
  card(s, 6.08, 1.68, 6.3, 4.75, '请先投票', 'A  完全一样\n\nB  基本一样\n\nC  可能明显不同\n\nD  不确定', C.teal, { headSize: 18, bodySize: 18, bodyAlign: 'center' });
  pill(s, '先收答案，再翻页', 9.77, 0.82, 2.25, C.orange);
  note(s, 4, '分步揭示第1页：先收集学员判断，不展示模型结果。');
}

// 04B 双次输出占位：投票后揭示
{
  const s = base(4, '投票之后：查看两次真实运行结果');
  card(s, 0.78, 1.68, 3.6, 4.75, '输入完全相同', '请为一款「销售拜访记录助手」\n写一句15字以内的产品介绍。\n\nA 完全一样\nB 基本一样\nC 可能明显不同\nD 不确定', C.red, { headSize: 17, bodySize: 14 });
  card(s, 4.7, 1.68, 3.72, 3.55, '第1次', '【课前替换】\n真实模型输出 A', C.teal, { headSize: 17, bodySize: 17, valign: 'mid', bodyAlign: 'center' });
  card(s, 8.7, 1.68, 3.72, 3.55, '第2次', '【课前替换】\n真实模型输出 B', C.purple, { headSize: 17, bodySize: 17, valign: 'mid', bodyAlign: 'center' });
  t(s, '模型：________  日期：________  可见设置：________', 4.72, 5.48, 7.7, 0.3, { fontSize: 10.5, color: C.muted, align: 'center' });
  t(s, '同一输入，结果可能不同；也可能相同或接近。', 4.72, 6.02, 7.7, 0.38, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 4, '必须替换为同一模型、同一输入的真实双次输出；不允许伪造。');
}

// 05A 互动起始：只展示模糊任务
{
  const s = base(5, '“帮我整理一下这次销售拜访”，还缺什么？');
  quote(s, '帮我整理一下这次销售拜访。', 1.0, 1.85, 11.2, 1.45, C.red, 28, 'center');
  t(s, '请先独立想30秒，再说出至少3项缺失信息。', 1.0, 4.25, 11.2, 0.5, { fontSize: 20, bold: true, color: C.red, align: 'center' });
  t(s, '提示：不要急着写Prompt，先判断这个任务为什么无法验收。', 1.0, 5.2, 11.2, 0.35, { fontSize: 13, color: C.muted, align: 'center' });
  pill(s, '先收答案，再翻页', 9.77, 0.82, 2.25, C.orange);
  note(s, 5, '分步揭示第1页：只展示模糊任务，至少收集3项学员答案。');
}

// 05B 模糊任务：收答案后揭示检查问题
{
  const s = base(5, '收答案后：至少检查这五类信息');
  quote(s, '帮我整理一下这次销售拜访。', 0.88, 1.72, 5.2, 1.25, C.red, 25);
  const qs = [['给谁使用？', C.red], ['具体要什么？', C.teal], ['哪些不能推测？', C.orange], ['怎样算合格？', C.purple], ['结果进入什么动作？', C.green]];
  qs.forEach((d, i) => label(s, d[0], 6.45 + (i % 2) * 2.75, 1.72 + Math.floor(i / 2) * 1.18, i === 4 ? 5.1 : 2.45, d[1], [C.softRed, C.softTeal, C.softOrange, C.softPurple, C.softGreen][i]));
  t(s, '先让学员说出至少3项，再逐步揭示。', 1.0, 5.75, 11.2, 0.4, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 5, '互动：先收答案，后揭示五个问题。');
}

// 06 个人使用vs产品交付
{
  const s = base(6, '会使用聊天工具，不等于已经完成AI产品设计');
  twoCol(s,
    ['个人使用', '结果不好，可以继续追问\n自己补背景、手工修改\n自己判断哪些内容可信\n偶尔成功，也可能有帮助'],
    ['产品交付', '不同用户稳定完成同类任务\n产品主动收集必要信息\n高风险内容有证据和确认\n结果可验收并进入后续流程'], 1.72, 4.65, C.blue, C.red);
  t(s, '训练目标：让专业要求被重复执行。', 1.0, 6.25, 11.2, 0.4, { fontSize: 18, bold: true, color: C.red, align: 'center' });
  note(s, 6);
}

// 07 生成不是查询
{
  const s = base(7, '大模型生成答案，不是简单找到一条现成答案');
  grid(s, [['搜索引擎', '找到相关内容或页面'], ['数据库', '按明确条件返回已有记录'], ['大模型', '依据当前上下文，逐步生成后续内容']], 3, 1.78, 3.25, [C.blue, C.teal, C.red]);
  t(s, '不简单等同', 1.0, 5.45, 3.45, 0.3, { fontSize: 13, bold: true, color: C.muted, align: 'center' });
  t(s, '不简单等同', 4.98, 5.45, 3.45, 0.3, { fontSize: 13, bold: true, color: C.muted, align: 'center' });
  t(s, '最低必要心智模型｜不是完整技术原理', 8.95, 5.45, 3.45, 0.3, { fontSize: 12, bold: true, color: C.red, align: 'center' });
  note(s, 7);
}

// 08A 互动起始：先猜下一步
{
  const s = base(8, '给定当前上下文，模型下一步会做什么？');
  card(s, 0.9, 1.78, 5.25, 4.5, '当前上下文', '用户输入\n＋\n已经生成的内容', C.blue, { headSize: 19, bodySize: 23, valign: 'mid', bodyAlign: 'center' });
  card(s, 6.45, 1.78, 5.95, 4.5, '请先猜一猜', '直接生成完整答案？\n\n查找固定答案？\n\n还是先判断下一个Token？', C.orange, { headSize: 19, bodySize: 18, valign: 'mid', bodyAlign: 'center' });
  pill(s, '先提问，再翻页', 10.02, 0.82, 2.0, C.orange);
  note(s, 8, '分步揭示第1页：让学员先预测生成过程，再展示完整流程。');
}

// 08B 完整生成流程
{
  const s = base(8, '大模型怎样一步一步生成答案？');
  flow(s, [['当前上下文', '已有输入与已生成内容'], ['候选Token', '计算后续可能性'], ['选择一个', '按生成策略选择'], ['加入上下文', '形成新的前文'], ['继续生成', '直到结束']], 1.82, 0.45, 12.42, [C.blue, C.teal, C.orange, C.purple, C.green]);
  quote(s, '每生成一步，后续判断的上下文也随之变化。', 1.1, 4.43, 10.9, 0.82, C.red, 19, 'center');
  t(s, '候选短语仅为可读化教学示意，可能包含一个或多个Token。', 1.0, 5.75, 11.2, 0.32, { fontSize: 11, color: C.muted, align: 'center' });
  note(s, 8);
}

// 09 背景改变后续
{
  const s = base(9, '只增加一句背景，后续为什么会变化？');
  twoCol(s,
    ['下班后，我准备去公园……', '可能续写\n\n散步｜跑步｜看花｜见朋友'],
    ['为了备战周末马拉松……', '可能续写\n\n跑步｜训练｜完成五公里'], 1.75, 3.9, C.blue, C.teal);
  t(s, '上下文会改变后续候选的可能性；错误或冲突背景也可能误导结果。', 1.0, 5.92, 11.2, 0.45, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 9);
}

// 10 候选概率
{
  const s = base(10, '一个位置，通常不只有一个合理后续');
  card(s, 0.78, 1.75, 4.0, 4.55, '当前上下文', '“下班后，\n我准备去公园……”', C.blue, { headSize: 17, bodySize: 24, valign: 'mid', bodyAlign: 'center' });
  const vals = [['散步', 0.34, C.teal], ['跑步', 0.29, C.blue], ['见朋友', 0.18, C.purple], ['拍照片', 0.11, C.orange], ['其他', 0.08, C.green]];
  vals.forEach((d, i) => {
    const y = 1.82 + i * 0.77;
    t(s, d[0], 5.22, y + 0.06, 1.2, 0.25, { fontSize: 11.5, bold: true, color: C.ink });
    shape(s, S.roundRect, 6.42, y, 5.4, 0.36, C.pale, C.pale, { rectRadius: 0.05 });
    shape(s, S.roundRect, 6.42, y, 5.4 * d[1] / 0.34, 0.36, d[2], d[2], { rectRadius: 0.05 });
    t(s, `${Math.round(d[1] * 100)}%`, 11.95, y + 0.055, 0.48, 0.23, { fontSize: 10, bold: true, color: d[2], align: 'right' });
  });
  t(s, '候选短语和数字仅为教学示意，不是模型真实Token切分或内部概率。', 5.2, 5.85, 7.2, 0.35, { fontSize: 10, color: C.muted, align: 'center' });
  note(s, 10);
}

// 11 Temperature
{
  const s = base(11, 'Temperature影响发散程度，不是“正确率按钮”');
  t(s, '较低', 0.92, 1.86, 1.0, 0.3, { fontSize: 14, bold: true, color: C.blue });
  t(s, '较高', 11.42, 1.86, 1.0, 0.3, { fontSize: 14, bold: true, color: C.orange, align: 'right' });
  shape(s, S.roundRect, 1.65, 1.92, 9.95, 0.28, C.softBlue, C.softBlue, { rectRadius: 0.05 });
  shape(s, S.ellipse, 5.75, 1.72, 0.66, 0.66, C.red, C.red);
  t(s, '更集中、更保守', 0.92, 2.52, 3.2, 0.32, { fontSize: 14, bold: true, color: C.blue });
  t(s, '更多样、更发散', 9.2, 2.52, 3.2, 0.32, { fontSize: 14, bold: true, color: C.orange, align: 'right' });
  grid(s, [['≠ 越低越正确', ''], ['≠ 设为0后绝对一致', ''], ['不能补充缺失知识', ''], ['不能替代任务与规则', '']], 4, 3.25, 1.5, [C.red, C.orange, C.purple, C.teal]);
  note(s, 11);
}

// 12 差异vs错误
{
  const s = base(12, '结果不同，与结果错误，是两件事');
  twoCol(s,
    ['合理差异', 'A｜拜访结束，关键信息自动成稿\n\nB｜让每次客户沟通，都留下可跟进的记录\n\n表达不同，但都可能合理'],
    ['更需要警惕', '编造材料中不存在的事实\n把未知内容说成确定事实\n与给定材料发生冲突\n生成无法验证的数据或承诺'], 1.72, 4.65, C.teal, C.red);
  t(s, '不同 ≠ 错误｜先检查是否与任务和证据冲突', 1.0, 6.28, 11.2, 0.35, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 12);
}

// 13 Demo
{
  const s = base(13, '一次Demo成功，只能证明“这一次可以”');
  card(s, 0.78, 1.72, 4.1, 4.65, '一次成功', '一次输入\n一次运行\n一条满意结果\n\n只能说明\n这一次得到可接受输出', C.orange, { headSize: 18, bodySize: 16, bodyAlign: 'center' });
  card(s, 5.25, 1.72, 2.18, 1.45, '用户表达', '是否变化', C.red, { headSize: 14, bodySize: 12, bodyAlign: 'center' });
  card(s, 7.62, 1.72, 2.18, 1.45, '材料质量', '是否完整', C.teal, { headSize: 14, bodySize: 12, bodyAlign: 'center' });
  card(s, 9.99, 1.72, 2.18, 1.45, '多次运行', '是否稳定', C.purple, { headSize: 14, bodySize: 12, bodyAlign: 'center' });
  card(s, 5.25, 3.42, 3.45, 1.45, '高风险字段', '是否可信', C.orange, { headSize: 14, bodySize: 12, bodyAlign: 'center' });
  card(s, 8.89, 3.42, 3.28, 1.45, '失败处理', '是否可恢复', C.green, { headSize: 14, bodySize: 12, bodyAlign: 'center' });
  t(s, 'Prompt可以降低部分错误，但不能把生成模型变成绝对规则系统。', 5.25, 5.45, 6.92, 0.62, { fontSize: 15.5, bold: true, color: C.red, align: 'center', valign: 'mid' });
  note(s, 13, '本课只建立测试意识，完整评测留到后续课程。');
}

// 14 Token
{
  const s = base(14, 'Token：模型处理和生成信息的基本单位');
  flow(s, [['自然语言', ''], ['Tokenizer', '切分规则'], ['Token序列', ''], ['模型处理', '']], 1.85, 0.8, 8.0, [C.blue, C.teal, C.purple, C.green]);
  card(s, 9.15, 1.75, 3.15, 4.2, 'Token不固定等于', '× 一个汉字\n× 一个中文词\n× 一个英文单词\n× 固定长度字符\n\n不同模型可能不同', C.red, { headSize: 16, bodySize: 13 });
  t(s, '若展示真实切分，必须标明模型或Tokenizer。', 1.0, 5.72, 7.7, 0.36, { fontSize: 12, color: C.muted, align: 'center' });
  note(s, 14);
}

// 15 上下文组成
{
  const s = base(15, '用户只输入一句话，不代表模型只看到一句话');
  shape(s, S.roundRect, 4.05, 1.72, 5.25, 4.5, C.paper, C.teal, { rectRadius: 0.06, line: { color: C.teal, width: 2 } });
  t(s, '本次上下文工作区', 4.35, 1.98, 4.65, 0.35, { fontSize: 17, bold: true, color: C.teal, align: 'center' });
  const ins = [['内部指令', C.red], ['用户输入', C.blue], ['历史对话', C.purple], ['上传文档', C.orange], ['业务材料', C.green], ['知识片段', C.teal]];
  ins.forEach((d, i) => label(s, d[0], 0.75 + (i % 2) * 1.55, 1.8 + Math.floor(i / 2) * 1.12, 1.35, d[1], [C.softRed, C.softBlue, C.softPurple, C.softOrange, C.softGreen, C.softTeal][i]));
  ins.forEach((d, i) => label(s, d[0], 4.45 + (i % 2) * 2.2, 2.65 + Math.floor(i / 2) * 0.86, 1.9, d[1], [C.softRed, C.softBlue, C.softPurple, C.softOrange, C.softGreen, C.softTeal][i]));
  card(s, 9.72, 1.72, 2.55, 4.5, '输出预算', '需要为本次生成结果预留空间\n\n具体计算与接口行为依模型而异', C.orange, { headSize: 16, bodySize: 12, bodyAlign: 'center' });
  t(s, '界面看不到的内容，也可能占用本次Token空间。', 1.0, 6.35, 11.2, 0.35, { fontSize: 15, bold: true, color: C.red, align: 'center' });
  note(s, 15);
}

// 16 窗口概念
{
  const s = base(16, '上下文窗口：一次调用可使用的有限Token空间');
  grid(s, [['上下文窗口', '一次调用可容纳输入与输出Token的有限空间'], ['实际上下文', '本次真正放进去的指令、输入和材料'], ['永久记忆', '产品是否保存并在未来重新提供'], ['模型已有知识', '训练中学到的能力，不等于本次材料'], ['产品数据库', '数据存在系统里，不代表已经给模型']], 3, 1.62, 1.72, [C.red, C.teal, C.purple, C.orange, C.blue]);
  t(s, '系统里有  ≠  模型这一次看到了', 1.0, 5.92, 11.2, 0.45, { fontSize: 20, bold: true, color: C.red, align: 'center' });
  note(s, 16, '具体窗口上限依模型和接口而异。');
}

// 17 容量vs质量
{
  const s = base(17, '上下文问题有两类：容量不够，和信息不好');
  twoCol(s,
    ['容量问题', '无法提交\n内容被截断\n早期信息被移除\n输出空间不足\n系统先摘要或压缩'],
    ['质量问题', '无关信息太多\n多个版本冲突\n重要事实埋得太深\n材料没有来源和时间\n任务本身过于复杂'], 1.72, 4.55, C.orange, C.red);
  t(s, '窗口回答“最多能装多少”｜设计回答“当前应该装什么”', 1.0, 6.24, 11.2, 0.4, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 17, '截断与报错行为依产品实现而异。');
}

// 18 上下文练习
{
  const s = base(18, '课堂练习：这五类材料，本次应该怎样提供？');
  const items = [['本次拜访转写', '含说话人标签'], ['客户基本资料', '当前商机阶段'], ['产品方案说明', '本次讨论版本'], ['全部CRM记录', '过去三年'], ['公司所有资料', '产品、制度、培训']];
  grid(s, items, 5, 1.72, 2.55, [C.teal, C.blue, C.purple, C.orange, C.red]);
  flow(s, [['必须提供', ''], ['按需提供', ''], ['暂不需要', '']], 4.55, 2.35, 8.65, [C.teal, C.orange, C.red]);
  t(s, '判断提示：删除这项信息，哪一项目标输出会明显变差？｜时间 3分钟', 1.0, 6.35, 11.2, 0.35, { fontSize: 14, bold: true, color: C.red, align: 'center' });
  note(s, 18, '先练习，不在屏幕显示老师答案。');
}

// 19 长Prompt
{
  const s = base(19, 'Prompt写得很长，为什么还是不能验收？');
  card(s, 0.78, 1.72, 7.35, 4.55, '看起来很专业', '你是一位非常资深、非常专业、逻辑严密、洞察敏锐、表达清晰的销售专家。请认真、全面、深入地分析下面的内容，输出一份高质量、专业、有价值、让人满意的销售拜访总结。', C.purple, { headSize: 17, bodySize: 16, valign: 'mid' });
  grid(s, [['给谁使用？', ''], ['交付什么？', ''], ['怎样判断可用？', '']], 1, 1.72, 1.25, [C.red, C.orange, C.teal]);
  t(s, '“专业、深入、高质量”是形容词，不是可检查的产品要求。', 1.0, 6.32, 11.2, 0.35, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 19);
}

// 20 设计顺序
{
  const s = base(20, '写Prompt之前，先按这个顺序思考');
  flow(s, [['用户与时点', '谁在何时用'], ['目标输出', '要交付什么'], ['验收标准', '怎样算可用'], ['必要信息', '模型需要什么'], ['Prompt与流程', '最后组织处理']], 1.85, 0.45, 12.42, [C.red, C.blue, C.teal, C.orange, C.purple]);
  quote(s, '这是产品设计顺序，不是系统运行顺序。', 1.1, 4.45, 10.9, 0.82, C.red, 19, 'center');
  note(s, 20);
}

// 21 用户与时点
{
  const s = base(21, '同一段拜访录音，给不同人，目标输出会不同');
  grid(s, [['一线销售', '诉求、异议、行动项'], ['销售主管', '阶段、缺口、风险'], ['客户', '结论、双方行动和时间'], ['CRM系统', '字段、状态和证据']], 4, 1.72, 2.5, [C.red, C.orange, C.teal, C.blue]);
  quote(s, '本课统一假设｜一线销售在拜访后10分钟内检查AI草稿、录入CRM并准备跟进。', 1.0, 4.72, 11.0, 1.0, C.red, 17, 'center');
  note(s, 21);
}

// 22 三个概念
{
  const s = base(22, '“要什么、怎么放、怎样算合格”是三个问题');
  grid(s, [['目标输出', 'AI要交付什么？\n下一步行动'], ['输出格式', '结果怎样呈现？\nCRM字段或表格'], ['验收标准', '怎样判断可用？\n人、事、时间、状态齐全']], 3, 1.72, 3.55, [C.red, C.blue, C.teal]);
  t(s, '“用表格输出”只能规定形式，不能保证表格里的事实正确。', 1.0, 5.78, 11.2, 0.42, { fontSize: 17, bold: true, color: C.red, align: 'center' });
  note(s, 22);
}

// 23 V0.1
{
  const s = base(23, '课堂练习：从目标输出反推必要上下文');
  card(s, 0.75, 1.65, 5.35, 4.85, '已给定', '用户｜一线销售\n时点｜拜访后10分钟内\n示例输出｜客户核心诉求\n\n请完成\n① 再补1项目标输出＋验收\n② 选择至少3项上下文\n③ 标记事实／背景／规则\n④ 写清支持哪项目标输出', C.teal, { headSize: 17, bodySize: 12.7 });
  const contextItems = [
    ['拜访转写', C.red], ['说话人标签', C.teal], ['客户资料', C.blue],
    ['商机阶段', C.orange], ['CRM字段', C.purple], ['缺失处理规则', C.green],
    ['全部历史资料', C.red], ['无关聊天', C.orange], ['产品方案', C.blue],
  ];
  contextItems.forEach((d, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    card(s, 6.35 + col * 2.02, 1.65 + row * 1.52, 1.84, 1.24, d[0], '', d[1], { headSize: 11.5, bodySize: 9.5, bodyAlign: 'center' });
  });
  pill(s, '04:00', 11.15, 0.82, 1.07, C.red);
  note(s, 23, '基准版不要求写完整Prompt。');
}

// 24 六要素
{
  const s = base(24, '六要素：检查任务说明有没有漏掉重要信息');
  grid(s, [['01 角色与能力', '承担什么职责？'], ['02 任务目标', '这次必须完成什么？'], ['03 背景与输入', '依据哪些事实？'], ['04 规则与约束', '哪些边界必须遵守？'], ['05 输出与验收', '怎样组织、怎样算可用？'], ['06 示例与迭代', '是否需要示例、检查或补问？']], 3, 1.62, 1.84, [C.red, C.orange, C.teal, C.purple, C.blue, C.green]);
  t(s, '本课程基于参考框架重新整理｜不是固定语法｜不要求全部写满', 1.0, 5.82, 11.2, 0.38, { fontSize: 14, bold: true, color: C.red, align: 'center' });
  note(s, 24);
}

// 25 角色任务
{
  const s = base(25, '角色不能替代任务，形容词不能替代责任');
  twoCol(s,
    ['角色', '不建议\n“全世界最顶级的销售专家”\n\n更具体\n“依据材料生成供销售检查的CRM草稿”'],
    ['任务', '不够\n“请专业地分析一下”\n\n更具体\n“整理诉求、异议、行动项和待确认信息”'], 1.72, 4.65, C.orange, C.teal);
  t(s, '角色回答“承担什么职责”｜任务回答“完成什么、用于什么”', 1.0, 6.25, 11.2, 0.35, { fontSize: 15, bold: true, color: C.red, align: 'center' });
  note(s, 25);
}

// 26 输入规则
{
  const s = base(26, '输入材料告诉模型“依据什么”，规则告诉模型“不能越过什么”');
  twoCol(s,
    ['业务背景与输入材料', '客户与商机信息\n本次拜访转写\nCRM字段说明\n当前产品方案\n特殊备注'],
    ['规则与约束', '只依据提供材料\n不把推测写成事实\n缺失标记待确认\n高风险结论附证据\n不生成成交概率'], 1.72, 4.65, C.blue, C.red);
  t(s, '广义上下文可以包含规则；为了检查，本课把规则单独列出。', 1.0, 6.25, 11.2, 0.35, { fontSize: 14.5, bold: true, color: C.red, align: 'center' });
  note(s, 26);
}

// 27 输出验收示例
{
  const s = base(27, '结果“长什么样”和“能不能用”，要分别定义');
  card(s, 0.75, 1.72, 3.75, 4.65, '输出格式', '表格或CRM字段\n固定字段顺序\n确定性状态标签', C.blue, { headSize: 18, bodySize: 15 });
  card(s, 4.78, 1.72, 3.75, 4.65, '验收标准', '关键事实有证据\n缺失信息明确标记\n行动项包含人、事、时间、状态', C.teal, { headSize: 18, bodySize: 14 });
  card(s, 8.8, 1.72, 3.75, 4.65, '示例与边界', '已确认｜明确同意小范围验证\n待确认｜是否扩大仍需讨论\n\n示例帮助理解，但不能替代测试', C.orange, { headSize: 18, bodySize: 12.8 });
  note(s, 27);
}

// 28 V0-V3
{
  const s = base(28, 'Prompt变好，不是因为更长，而是每次修改解决一个问题');
  flow(s, [['V0', '整理一下\n全部不清楚'], ['V1', '明确提取内容\n补任务'], ['V2', '依据＋事实边界\n补输入和规则'], ['V3', '证据＋完整行动项\n补验收']], 1.8, 0.62, 12.0, [C.red, C.orange, C.teal, C.green]);
  t(s, '每一项新增要求，都应该对应一个具体问题。', 1.0, 5.05, 11.2, 0.45, { fontSize: 19, bold: true, color: C.red, align: 'center' });
  note(s, 28);
}

// 29 V0.2
{
  const s = base(29, '课堂练习：补全结构化Prompt关键骨架');
  grid(s, [['任务目标', '这次完成什么？'], ['关键规则', '先写最重要的一条'], ['输出验收', '怎样判断可用？']], 3, 1.72, 3.2, [C.red, C.orange, C.teal]);
  t(s, '同伴只检查一处最重要缺口', 1.0, 5.35, 11.2, 0.4, { fontSize: 18, bold: true, color: C.red, align: 'center' });
  t(s, '基准版：V0.2关键骨架｜充分互动版：补齐完整骨架', 1.0, 5.92, 11.2, 0.32, { fontSize: 12, color: C.muted, align: 'center' });
  pill(s, '04:00', 11.15, 0.82, 1.07, C.red);
  note(s, 29);
}

// 30 休息
{
  const s = base(30, '', '', { noHeader: true, noFooter: true, bg: C.dark });
  t(s, '休息 5 分钟', 0.82, 0.92, 5.5, 0.65, { fontSize: 34, bold: true, color: C.white });
  t(s, '回来后：把Prompt变成产品', 0.86, 1.85, 7.5, 0.4, { fontSize: 17, color: 'D9D4CC' });
  t(s, '05:00', 8.75, 0.9, 3.5, 1.0, { fontSize: 55, bold: true, color: 'E09A78', align: 'right' });
  const rules = [
    ['保存V0.2', '保留刚才的关键骨架'],
    ['打开课堂设计单', '06_课堂练习与作业材料 / 01_Prompt产品化设计单_课堂版.md'],
    ['打开案例转写', '06_课堂练习与作业材料 / 03_销售案例背景与转写_学员版.md'],
  ];
  rules.forEach((d, i) => { rule(s, 0.85, 3.05 + i * 1.02, 11.4, '4A4B48'); t(s, String(i + 1).padStart(2, '0'), 0.88, 3.28 + i * 1.02, 0.55, 0.22, { fontSize: 10, bold: true, color: 'E09A78' }); t(s, d[0], 1.7, 3.22 + i * 1.02, 2.25, 0.32, { fontSize: 14.2, bold: true, color: C.white }); t(s, d[1], 4.18, 3.24 + i * 1.02, 7.75, 0.34, { fontSize: i === 0 ? 12.2 : 10.4, color: 'C7C4BE' }); });
  note(s, 30);
}

// 31 Prompt还不是产品
{
  const s = base(31, '如果每位销售都要自己补完整Prompt，产品解决了什么？');
  card(s, 0.78, 1.72, 4.2, 4.65, '一段很长的Prompt', '角色＋任务＋输入＋规则＋输出＋示例……', C.purple, { headSize: 18, bodySize: 18, bodyAlign: 'center', valign: 'mid' });
  const userBurden = [
    ['理解框架', C.red], ['寻找资料', C.orange],
    ['记住规则', C.teal], ['重复输入', C.purple],
    ['发现缺失', C.green], ['判断风险', C.blue],
  ];
  userBurden.forEach((d, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    card(s, 5.3 + col * 3.57, 1.72 + row * 1.42, 3.3, 1.18, d[0], '', d[1], { headSize: 13, bodySize: 10, bodyAlign: 'center' });
  });
  t(s, '好Prompt是产品内部能力的原材料，不是最终用户界面。', 5.3, 6.05, 6.87, 0.42, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 31);
}

// 32 两个维度
{
  const s = base(32, '产品化需要回答两个不同问题');
  card(s, 0.78, 1.72, 5.78, 4.65, '维度一｜信息怎样进入？', '产品固定\n用户提供\n系统获取\n\n使用蓝色系', C.blue, { headSize: 19, bodySize: 20, bodyAlign: 'center' });
  card(s, 6.78, 1.72, 5.78, 4.65, '维度二｜结果怎样处理风险？', '普通展示\n建议检查\n必须确认\n\n使用橙色系', C.orange, { headSize: 19, bodySize: 20, bodyAlign: 'center' });
  t(s, '两个维度可以同时成立，不是“四选一”。', 1.0, 6.28, 11.2, 0.35, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 32);
}

// 33 信息供给
{
  const s = base(33, '维度一：产品固定、用户提供、系统获取');
  grid(s, [['产品固定', '长期稳定？\n角色、字段、事实规则'], ['用户提供', '只有用户知道或必须选择？\n本次重点、特殊备注'], ['系统获取', '系统已有且有权使用？\n转写、客户资料、商机阶段']], 3, 1.72, 3.75, [C.blue, C.teal, C.purple]);
  t(s, '每次变化 ≠ 必须手填｜系统获取 ≠ 默认获取全部数据', 1.0, 5.85, 11.2, 0.4, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 33);
}

// 34 风险处理
{
  const s = base(34, '维度二：普通展示、建议检查、必须确认');
  grid(s, [['普通展示', '错误影响低\n直接查看或编辑'], ['建议检查', '可能影响判断\n显示证据、状态和修改入口'], ['必须确认', '高风险结论或外部动作\n有权限用户确认后执行']], 3, 1.72, 3.45, [C.green, C.orange, C.red]);
  label(s, '日期', 1.25, 5.55, 1.2, C.orange, C.softOrange); label(s, '金额', 2.75, 5.55, 1.2, C.orange, C.softOrange); label(s, '责任人', 4.25, 5.55, 1.35, C.orange, C.softOrange); label(s, '客户承诺', 5.9, 5.55, 1.55, C.red, C.softRed); label(s, '写入CRM', 7.75, 5.55, 1.55, C.red, C.softRed); label(s, '发送消息', 9.6, 5.55, 1.55, C.red, C.softRed);
  note(s, 34);
}

// 35 V0.3
{
  const s = base(35, '课堂练习：把两个维度变成产品交互');
  card(s, 0.78, 1.72, 5.78, 4.65, 'A 信息供给', '至少2项产品固定\n至少1项用户提供\n至少2项系统获取\n\n删除1项无意义手填', C.blue, { headSize: 19, bodySize: 17, bodyAlign: 'center' });
  card(s, 6.78, 1.72, 5.78, 4.65, 'B 风险处理', '至少1项执行前必须确认\n\n确认什么？\n依据什么？\n谁确认？\n确认后发生什么？', C.orange, { headSize: 19, bodySize: 15, bodyAlign: 'center' });
  pill(s, '03:00', 11.15, 0.82, 1.07, C.red);
  note(s, 35);
}

// 36 第0课 vs 第2课
{
  const s = base(36, '同一个案例，两节课训练目标不同');
  twoCol(s,
    ['第0课', '场景是否值得做？\nAI可以承担哪些环节？\n人与AI怎样初步分工？'],
    ['第2课', '需要哪些上下文和规则？\n怎样找错并最小修改？\n怎样设计信息供给和风险处理？'], 1.72, 4.0, C.blue, C.red);
  flow(s, [['基线结果', ''], ['证据找错', ''], ['最小修改', ''], ['二次验收', ''], ['修正流程', '']], 5.05, 1.2, 10.9, [C.red, C.orange, C.teal, C.purple, C.green]);
  note(s, 36);
}

// 37 案例背景
{
  const s = base(37, '销售要检查的不是漂亮纪要，而是可进入CRM的草稿');
  pill(s, '教学虚构案例', 10.45, 0.82, 1.78, C.orange);
  card(s, 0.75, 1.68, 5.2, 4.75, '项目背景', '供应商｜北辰云科\n客户｜禾味餐饮，48家直营门店\n阶段｜初次演示后，讨论试点与技术可行性\n用户｜一线销售林然\n时点｜拜访结束后10分钟内\n下一步｜检查事实、录入CRM、推进评审', C.teal, { headSize: 18, bodySize: 13.2 });
  card(s, 6.25, 1.68, 2.9, 4.75, '本次只整理', '客户诉求\n异议与风险\n试点状态\n行动项\n待确认信息', C.blue, { headSize: 17, bodySize: 14, bodyAlign: 'center' });
  card(s, 9.45, 1.68, 2.9, 4.75, '本次不做', '成交预测\n自动报价\n完整方案\n自动写入CRM', C.red, { headSize: 17, bodySize: 14, bodyAlign: 'center' });
  note(s, 37);
}

// 38 证据扫描
{
  const s = base(38, '先扫描高风险证据，不从头逐字朗读');
  card(s, 0.78, 1.72, 5.55, 4.65, '重点圈出', '① 数字、日期和人名\n② 确认、可能、暂定、不承诺\n③ 客户主动纠正的语句\n④ 负责人和下一步行动', C.red, { headSize: 19, bodySize: 17 });
  card(s, 6.65, 1.72, 5.75, 4.65, '扫描标记', 'F  事实或明确确认\nU  不确定、待确认或附带条件\nA  行动项\nR  风险或限制\n\n不要在PPT上找答案，请打开完整转写。', C.teal, { headSize: 19, bodySize: 15 });
  pill(s, '90秒', 11.15, 0.82, 1.07, C.red);
  note(s, 38, '不在本页摘录关键原文，避免提前泄题。');
}

// 39 基线结果
{
  const s = base(39, '第一轮基线结果：读起来顺，但能直接使用吗？');
  pill(s, '教学构造的错误样例｜不代表任何特定模型', 8.65, 0.82, 3.58, C.orange);
  card(s, 0.78, 1.62, 11.55, 4.95, '第一轮结果', '禾味餐饮拥有48家门店，目前数据分散、人工整理效率较低，希望通过FlowSight实现经营数据自动汇总和整改闭环。\n\n客户认可产品方案，计划先在8家门店试点，试点成功后采购并推广到全部门店。客户预算虽然紧张，但不是主要问题，预计6月底上线。技术上需要对接POS、会员和巡检系统，预计两周完成。\n\n下一步：北辰云科提供试点和技术方案，禾味餐饮提供数据，双方安排技术会议。', C.orange, { headSize: 17, bodySize: 13.6 });
  t(s, '你能指出至少3个“有证据的具体问题”吗？', 1.0, 6.32, 11.2, 0.35, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 39, '禁止制作伪模型聊天截图。');
}

// 40 证据化找错
{
  const s = base(40, '“感觉不准确”不够，必须写清四件事');
  flow(s, [['结果写了什么？', ''], ['具体问题？', ''], ['原文证据？', ''], ['违反哪条验收？', '']], 1.75, 0.62, 12.0, [C.red, C.orange, C.teal, C.purple]);
  card(s, 0.88, 4.15, 4.25, 1.55, '不合格', '“太笼统、不够专业、感觉有问题”', C.red, { shadow: false, headSize: 14, bodySize: 11.5 });
  card(s, 5.45, 4.15, 6.85, 1.55, '合格表达示例', '结果遗漏了材料中明确的截止时间；原文包含具体日期；违反“行动项必须包含截止时间”。', C.teal, { shadow: false, headSize: 14, bodySize: 11.2 });
  pill(s, '03:00', 11.15, 0.82, 1.07, C.red);
  note(s, 40, '示例不对应本案例核心答案。');
}

// 41 参考讲评
{
  const s = base(41, '语言通顺，仍然可能发生这些业务错误');
  const items = [
    ['8家试点', '可能方案写成确定事实'], ['采购推广', '销售推测写成客户承诺'],
    ['6月底上线', '希望写成确定计划'], ['两周对接', '条件估计写成正式工期'],
    ['预算不重要', '自行弱化商务风险'], ['行动项', '缺负责人和日期'],
    ['安全风险', '遗漏关键异议'], ['待确认', '预算、决策人、时间未列出']
  ];
  grid(s, items, 4, 1.58, 2.05, [C.red, C.orange, C.purple, C.blue, C.teal, C.green, C.red, C.orange]);
  t(s, 'AI产品验收不能只看“读起来像不像好答案”。', 1.0, 6.08, 11.2, 0.38, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 41, '课堂重点讲4—5条，其余作为完整答案保留。');
}

// 42 最小修改
{
  const s = base(42, '根据Bad Case做最小修改，不从头重写');
  flow(s, [['选择问题', '最重要1—3条'], ['定位Prompt', '任务／输入／规则／输出'], ['写通用修改', '能迁移到其他样例']], 1.82, 1.25, 10.8, [C.red, C.orange, C.teal]);
  twoCol(s,
    ['不建议', '必须输出5家试点'],
    ['更通用', '区分已确认、讨论中和待确认；\n可能方案不能改写为确定事实'], 4.15, 1.55, C.red, C.teal);
  pill(s, '03:00', 11.15, 0.82, 1.07, C.red);
  note(s, 42);
}

// 43 真实二轮占位
{
  const s = base(43, '第二轮真实输出：改完以后，发生了什么？');
  card(s, 0.78, 1.68, 3.55, 4.75, '本轮修改', '填写课堂最终采用的1—3条通用修改\n\n完整Prompt放在课堂材料或演讲者备注', C.orange, { headSize: 17, bodySize: 13 });
  card(s, 4.65, 1.68, 7.72, 4.75, '【课前替换】真实第二轮输出', '模型：________\n生成日期：________\n可见设置：________\n输出是否节选：是／否\n\n改善了什么？\n仍然有什么问题？\n是否出现新的Bad Case？', C.teal, { headSize: 18, bodySize: 14, bodyAlign: 'center' });
  note(s, 43, '必须使用真实原始输出；裁剪时标注“节选”；不允许润色后冒充。');
}

// 44 二轮验收
{
  const s = base(44, '一次改善之后，还要完成验收和产品流程');
  card(s, 0.78, 1.68, 5.75, 4.8, '第二轮验收', '□ 确定性状态正确\n□ 未把讨论或估计写成承诺\n□ 关键事实完整且可核对\n□ 行动项支持后续执行\n□ 缺失信息明确处理\n□ 高风险内容进入确认流程', C.teal, { headSize: 18, bodySize: 14.2 });
  card(s, 6.8, 1.68, 5.55, 4.8, '产品流程三句话', '1 系统已经准备________\n\n2 用户仍需填写、选择或确认________\n\n3 必须确认后才能写入CRM的是________', C.orange, { headSize: 18, bodySize: 14.5 });
  t(s, '一次结果改善 ≠ 系统已经稳定 ≠ 可以取消人工确认', 1.0, 6.35, 11.2, 0.35, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 44, '补入第43页真实结果后，更新本轮通过项、剩余问题和新问题。');
}

// 45 四类初筛
{
  const s = base(45, 'Bad Case先别急着改Prompt：问题首先在哪里？');
  grid(s, [['任务', '服务谁、完成什么？\n“整理一下”'], ['上下文', '事实足够、可靠吗？\n缺材料、版本冲突'], ['规则', '推测和缺失怎样处理？\n可能写成确定'], ['输出', '能否检查并进入下一步？\n缺字段、时间或状态']], 2, 1.62, 2.05, [C.red, C.blue, C.orange, C.teal]);
  quote(s, '“可能8家”被写成“确定8家”｜优先检查：规则与约束', 1.0, 5.92, 11.0, 0.65, C.red, 16.5, 'center');
  note(s, 45, '这是Prompt相关第一层初筛，不是完整根因体系。');
}

// 46 Prompt边界
{
  const s = base(46, '出现这些信号时，不要继续无限堆Prompt');
  card(s, 0.78, 1.65, 6.0, 4.9, '五个边界信号', '① 缺少企业私有、最新或大量知识\n② 多个稳定步骤、条件分支和中间检查\n③ 需要查询、写入、发送或创建任务\n④ 原始录音、转写或上游数据已经错误\n⑤ 任务等已说清，多条样本仍持续不达标', C.red, { headSize: 18, bodySize: 14.5 });
  card(s, 7.08, 1.65, 5.25, 4.9, '统一判断句式', '这条问题首先属于________\n\n我会先验证________\n\n如果仍不能解决，再考虑________', C.teal, { headSize: 18, bodySize: 15, bodyAlign: 'center' });
  t(s, '技术名词不是升级路线，问题类型才是下一步方案的起点。', 1.0, 6.34, 11.2, 0.35, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 46, '不展开RAG、Workflow、Agent、Function Calling或MCP定义。');
}

// 47 四个结论
{
  const s = base(47, '今天只带走四个结论');
  const a = [['生成', '不是固定查询\n一次Demo不能证明稳定'], ['上下文', '有限工作区\n信息要必要、可靠、清楚'], ['Prompt', '先定义用户、输出和验收'], ['产品化', '产品承担重复输入与风险处理']];
  a.forEach((d, i) => card(s, 0.72 + i * 3.08, 1.72, 2.82, 4.05, d[0], d[1], [C.red, C.teal, C.purple, C.orange][i], { headSize: 19, bodySize: 14, bodyAlign: 'center' }));
  t(s, 'V0.1 输出与上下文 → V0.2关键骨架 → V0.3产品化 → 案例验证', 1.0, 6.18, 11.2, 0.38, { fontSize: 14, bold: true, color: C.red, align: 'center' });
  note(s, 47);
}

// 48 作业
{
  const s = base(48, '课后完成V1.0：先交必做版，进阶版自选');
  card(s, 0.75, 1.62, 7.65, 4.98, '必做版｜完成即合格', '□ 场景：谁在什么时点完成什么任务\n□ 输出：至少2项结果＋可检查标准\n□ 设计：必要上下文＋Prompt关键骨架\n□ 产品化：信息供给＋至少1项高风险处理\n□ 验证：运行1条样例＋1条Bad Case＋最小修改', C.teal, { headSize: 19, bodySize: 15.2 });
  card(s, 8.7, 1.62, 3.65, 4.98, '进阶版', '再补2—3条样例\n保留每轮原始记录\n补齐证据、权限与确认\n画产品流程', C.orange, { headSize: 18, bodySize: 14.2 });
  t(s, '没有工作场景可使用公开或虚构材料｜禁止提交敏感数据、隐私或公司机密', 1.0, 6.4, 11.2, 0.28, { fontSize: 10.5, bold: true, color: C.red, align: 'center' });
  note(s, 48, '详细字段见课后作业模板；不要逐条朗读。');
}

// 49 结束
{
  const s = base(49, '', '', { noHeader: true, noFooter: true });
  shape(s, S.rect, 0, 0, 0.22, 7.5, C.red, C.red);
  t(s, '从一次好回答，\n走向可重复交付。', 0.9, 1.1, 10.8, 1.25, { fontSize: 35, bold: true, color: C.ink });
  rule(s, 0.92, 2.85, 3.0, C.red, 3);
  t(s, '不要先追求“万能AI产品”', 0.92, 3.45, 5.4, 0.35, { fontSize: 18, bold: true, color: C.red });
  t(s, '先让一个具体任务被定义、被验证、被稳定交付。', 0.92, 4.0, 7.2, 0.42, { fontSize: 15, color: C.text });
  t(s, '下一课', 8.2, 3.42, 3.8, 0.3, { fontSize: 11, bold: true, color: C.muted, align: 'right', charSpacing: 1.2 });
  t(s, '场景判断与\nAI产品形态选择', 7.0, 3.92, 5.0, 0.95, { fontSize: 22, bold: true, color: C.red, align: 'right' });
  pill(s, '117—120分钟：集中答疑／确认作业选题', 8.0, 6.42, 4.0, C.red);
  t(s, 'AI 产品经理系统课程 · 第一阶段', 0.92, 6.65, 4.2, 0.2, { fontSize: 9, color: C.muted });
  note(s, 49);
}

if (pptx._slides.length !== 52) throw new Error(`Expected 52 slides, got ${pptx._slides.length}`);
pptx.writeFile({ fileName: OUT });
