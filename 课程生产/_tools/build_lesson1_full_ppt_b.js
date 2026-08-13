const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第01课：AI产品经理岗位认知与转型路径';
pptx.title = 'AI产品经理岗位认知与转型路径';
pptx.company = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };

const S = pptx.ShapeType;
const OUT = '课程生产/第一阶段/第01课_AI产品经理岗位认知与转型路径_重做工作区/20_正式PPT_第01课_AI产品经理岗位认知与转型路径.pptx';
const ROOT = '课程生产/第一阶段/第01课_AI产品经理岗位认知与转型路径_重做工作区';

const C = {
  bg: 'F4F0E8', paper: 'FBF9F4', ink: '20211F', text: '292A27', muted: '77766F', line: 'D8D0C2',
  red: 'C35331', teal: '387B77', purple: '74546A', orange: 'D18A38', green: '527553', blue: '405D78',
  softRed: 'F2E2DB', softTeal: 'DDEBE8', softPurple: 'E8E0E7', softOrange: 'F2E7D5', softGreen: 'E2E9DE', softBlue: 'DDE5EB',
  white: 'FFFFFF', dark: '1D1E1C'
};

const moduleOf = n => n <= 6 ? '01 开场与认知诊断' : n <= 17 ? '02 AI产品经理到底在做什么' : n <= 30 ? '03 岗位为什么有不同方向' : n <= 38 ? '04 结合背景判断方向' : n <= 43 ? '05 完全零基础怎么走' : n <= 47 ? '06 现场方向诊断' : '07 总结与课后作业';
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
function pill(slide, text, x, y, w, color = C.red, fill = color) {
  shape(slide, S.roundRect, x, y, w, 0.34, fill, fill, { rectRadius: 0.1 });
  t(slide, text, x + 0.08, y + 0.075, w - 0.16, 0.16, { fontSize: 9, bold: true, color: C.white, align: 'center' });
}
function base(n, title, kicker = kickerOf(n), opts = {}) {
  const s = pptx.addSlide(); s.background = { color: opts.bg || C.bg };
  if (!opts.noHeader) {
    pill(s, String(n).padStart(2, '0'), 0.62, 0.3, 0.62, C.red);
    t(s, kicker, 1.42, 0.38, 4.8, 0.18, { fontSize: 9.5, bold: true, color: C.muted, charSpacing: 1.1 });
    t(s, title, 0.62, 0.78, 11.9, 0.52, { fontSize: 25, bold: true, color: C.ink });
    rule(s, 0.62, 1.42, 12.05);
  }
  if (!opts.noFooter) {
    t(s, 'AI 产品经理系统课程 · 第一阶段', 0.62, 7.06, 4.1, 0.16, { fontSize: 8, color: C.muted });
    t(s, moduleOf(n), 8.2, 7.06, 3.7, 0.16, { fontSize: 8, color: C.muted, align: 'right' });
    t(s, String(n).padStart(2, '0'), 12.1, 7.04, 0.55, 0.18, { fontSize: 8.5, color: C.muted, align: 'right' });
  }
  return s;
}
function card(slide, x, y, w, h, head, body, accent = C.red, o = {}) {
  shape(slide, S.roundRect, x, y, w, h, o.fill || C.paper, o.line || C.line, { rectRadius: 0.05, shadow: o.shadow === false ? undefined : { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: 0.08 } });
  shape(slide, S.rect, x, y, o.top ? w : 0.07, o.top ? 0.08 : h, accent, accent);
  t(slide, head, x + 0.27, y + 0.25, w - 0.52, 0.34, { fontSize: o.headSize || 15, bold: true, color: o.headColor || C.ink });
  t(slide, body, x + 0.27, y + 0.79, w - 0.54, h - 1.03, { fontSize: o.bodySize || 11.4, color: o.bodyColor || C.muted, breakLine: false, valign: o.valign || 'top', paraSpaceAfterPt: 6 });
}
function label(slide, text, x, y, w, color = C.red, fill = C.softRed) {
  shape(slide, S.roundRect, x, y, w, 0.38, fill, fill, { rectRadius: 0.08 });
  t(slide, text, x + 0.08, y + 0.09, w - 0.16, 0.17, { fontSize: 9.5, bold: true, color, align: 'center' });
}
function arrow(slide, x, y, w = 0.4, color = C.line) {
  shape(slide, S.chevron, x, y, w, 0.46, color, color);
}
function row(slide, idx, head, body, y, accent = C.red, x = 0.82, w = 11.5) {
  t(slide, String(idx).padStart(2, '0'), x, y + 0.04, 0.55, 0.25, { fontSize: 11.5, bold: true, color: accent });
  t(slide, head, x + 0.82, y, 2.05, 0.32, { fontSize: 15, bold: true, color: C.ink });
  t(slide, body, x + 3.2, y + 0.03, w - 3.2, 0.32, { fontSize: 12.2, color: C.muted });
  rule(slide, x, y + 0.58, w);
}
function quote(slide, text, x, y, w, h, accent = C.red, size = 22) {
  shape(slide, S.rect, x, y, 0.08, h, accent, accent);
  t(slide, text, x + 0.35, y + 0.1, w - 0.35, h - 0.2, { fontSize: size, bold: true, color: C.ink, valign: 'mid' });
}
function flow(slide, items, y = 2.2, x = 0.62, totalW = 12.0, accentList = []) {
  const gap = 0.24, w = (totalW - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const xx = x + i * (w + gap), a = accentList[i] || [C.red, C.orange, C.teal, C.purple, C.green][i % 5];
    shape(slide, S.roundRect, xx, y, w, 1.7, C.paper, C.line, { rectRadius: 0.05 });
    t(slide, String(i + 1).padStart(2, '0'), xx + 0.18, y + 0.2, 0.48, 0.2, { fontSize: 10, bold: true, color: a });
    t(slide, it[0], xx + 0.18, y + 0.65, w - 0.36, 0.32, { fontSize: 13.5, bold: true, color: C.ink, align: 'center' });
    if (it[1]) t(slide, it[1], xx + 0.18, y + 1.14, w - 0.36, 0.3, { fontSize: 9.8, color: C.muted, align: 'center' });
  });
}
function grid(slide, items, cols = 3, y = 1.7, h = 1.75, colors = [C.red, C.teal, C.orange, C.purple, C.green, C.blue]) {
  const x0 = 0.72, gap = 0.24, w = (11.9 - gap * (cols - 1)) / cols;
  items.forEach((it, i) => {
    const x = x0 + (i % cols) * (w + gap), yy = y + Math.floor(i / cols) * (h + 0.24);
    card(slide, x, yy, w, h, it[0], it[1], colors[i % colors.length], { shadow: false, headSize: it[2] || 15, bodySize: it[3] || 11.2 });
  });
}

// Parse teacher-only source notes from detailed markdown files.
const sourceNotes = {};
for (const f of fs.readdirSync(ROOT).filter(x => /^1[2-8]_PPT逐页内容_.*\.md$/.test(x))) {
  const text = fs.readFileSync(path.join(ROOT, f), 'utf8');
  const re = /^## 第(\d+)页[^\n]*\n([\s\S]*?)(?=^## 第\d+页|\Z)/gm;
  let m;
  while ((m = re.exec(text))) {
    const n = Number(m[1]);
    const sm = m[2].match(/### 老师备注来源（不投屏）\s*```text\s*([\s\S]*?)```/);
    sourceNotes[n] = sm ? sm[1].trim() : '本课程教学设计';
  }
}
function note(slide, n, extra = '') {
  slide.addNotes(`老师备注来源（不投屏）\n${sourceNotes[n] || '本课程教学设计'}${extra ? `\n\n${extra}` : ''}\n\n详细讲解见对应老师版讲义。`);
}

// 01 Cover
{
  const s = base(1, '', '', { noHeader: true, noFooter: true, bg: C.bg });
  shape(s, S.rect, 0, 0, 0.22, 7.5, C.red, C.red);
  t(s, '01', 0.88, 0.72, 1.1, 0.66, { fontSize: 35, bold: true, color: C.red });
  t(s, 'AI产品经理岗位认知\n与转型路径', 0.88, 1.68, 8.2, 1.42, { fontSize: 37, bold: true, color: C.ink, breakLine: false });
  t(s, '从岗位地图到个人方向假设', 0.92, 3.48, 7.3, 0.4, { fontSize: 17, color: C.muted });
  rule(s, 0.92, 4.15, 3.0, C.red, 3);
  t(s, 'ROLE  ×  DIRECTION  ×  EVIDENCE', 8.35, 1.05, 3.95, 0.25, { fontSize: 10, bold: true, color: C.muted, charSpacing: 2.0, align: 'right' });
  t(s, '职业判断不是选一个热门标签，\n而是用证据形成可验证的假设。', 8.25, 4.72, 4.05, 1.0, { fontSize: 17, bold: true, color: C.red, align: 'right', breakLine: false });
  pill(s, '第一阶段 · 第01课', 9.66, 6.43, 2.55, C.red);
  note(s, 1);
}

// 02 Three questions
{
  const s = base(2, '今天只回答三个问题');
  const a = [['岗位是什么？', '真正负责什么，而不只是会什么工具'], ['为什么分方向？', '同名岗位为什么工作内容差异很大'], ['我优先验证什么？', '从已有证据出发形成方向假设']];
  a.forEach((d, i) => { card(s, 0.82 + i * 4.12, 1.78, 3.55, 3.75, d[0], d[1], [C.red, C.teal, C.orange][i], { headSize: 18, bodySize: 12.5 }); if (i < 2) arrow(s, 4.48 + i * 4.12, 3.28, 0.35); });
  t(s, '最终产出｜个人方向判断卡 V0.1', 2.15, 5.98, 9.0, 0.38, { fontSize: 17, bold: true, color: C.red, align: 'center' });
  note(s, 2);
}

// 03 Poll
{
  const s = base(3, '你认为公司招聘AI产品经理时，最看重什么？');
  const opts = [['A', 'AI工具熟练度'], ['B', '算法与模型原理'], ['C', '问题与业务判断'], ['D', '产品方案与落地'], ['E', '效果评估与优化']];
  opts.forEach((d, i) => {
    const x = 0.62 + i * 2.48;
    shape(s, S.roundRect, x, 1.85, 2.2, 3.55, C.paper, C.line, { rectRadius: 0.06 });
    t(s, d[0], x, 2.18, 2.2, 0.55, { fontSize: 32, bold: true, color: [C.red, C.teal, C.orange, C.purple, C.green][i], align: 'center' });
    t(s, d[1], x + 0.18, 3.23, 1.84, 0.72, { fontSize: 15, bold: true, color: C.ink, align: 'center', valign: 'mid' });
  });
  pill(s, '现场投票', 10.7, 0.82, 1.5, C.red);
  t(s, '先选择，再说明你的理由。', 3.2, 5.97, 6.9, 0.35, { fontSize: 14, bold: true, color: C.muted, align: 'center' });
  note(s, 3, '互动：先让学员投票，不要立即公布“标准答案”。');
}

// 04 Contextual answer
{
  const s = base(4, '这道题没有脱离岗位场景的唯一答案');
  quote(s, '不同岗位，能力权重不同；\n共同点是：都不止会用工具。', 0.85, 1.85, 7.1, 2.0, C.red, 25);
  const tags = [['To B应用', C.teal, C.softTeal], ['To C应用', C.purple, C.softPurple], ['平台／能力', C.blue, C.softBlue], ['解决方案', C.orange, C.softOrange]];
  tags.forEach((d, i) => label(s, d[0], 8.28 + (i % 2) * 2.05, 1.95 + Math.floor(i / 2) * 1.1, 1.7, d[1], d[2]));
  t(s, '今天先建立判断方法，后面再展开四个方向。', 0.9, 5.25, 11.1, 0.42, { fontSize: 16, bold: true, color: C.red });
  note(s, 4);
}

// 05 Misconceptions
{
  const s = base(5, '转AI最常见的三个误解');
  const a = [['误解 01', '会用AI工具\n＝AI产品经理', '工具只是手段，岗位对问题和结果负责'], ['误解 02', 'AI产品岗位\n都差不多', '服务对象、目标、形态和交付方式不同'], ['误解 03', '过去的经验\n全部作废', '旧经验要转成可验证的资产证据']];
  a.forEach((d, i) => {
    const x = 0.82 + i * 4.12;
    card(s, x, 1.75, 3.55, 4.25, d[0], d[1], [C.red, C.orange, C.teal][i], { headSize: 13, bodySize: 20, bodyColor: C.ink });
    t(s, '修正｜' + d[2], x + 0.27, 4.85, 3.0, 0.65, { fontSize: 11.2, bold: true, color: [C.red, C.orange, C.teal][i] });
  });
  note(s, 5);
}

// 06 Hypothesis loop
{
  const s = base(6, '今天形成的是方向假设，不是职业判决');
  flow(s, [['方向假设', '基于当前证据'], ['真实JD', '核验岗位要求'], ['项目行动', '创造新证据'], ['持续修正', '允许改变方向']], 2.0, 1.0, 11.3, [C.red, C.teal, C.orange, C.green]);
  t(s, '目标：知道下一步优先验证什么，而不是给自己贴一个永久标签。', 1.1, 5.05, 11.0, 0.5, { fontSize: 18, bold: true, color: C.red, align: 'center' });
  note(s, 6);
}

// 07 Case background
{
  const s = base(7, '案例：门店员工每天都在重复问同样的问题');
  pill(s, '教学模拟案例', 10.5, 0.82, 1.72, C.orange);
  card(s, 0.78, 1.68, 4.2, 4.75, '连锁餐饮门店', '门店员工每天都要确认：\n\n新品怎么制作？\n活动什么时候开始？\n这张券能不能用？\n客诉应该如何处理？', C.red, { headSize: 18, bodySize: 14 });
  const src = [['群聊消息', '信息零散，容易被淹没'], ['文档与表格', '版本多，更新不同步'], ['总部人员', '重复回答，且口径可能不一致']];
  src.forEach((d, i) => card(s, 5.35, 1.68 + i * 1.57, 6.75, 1.25, d[0], d[1], [C.teal, C.orange, C.purple][i], { shadow: false, headSize: 14, bodySize: 11 }));
  note(s, 7);
}

// 08 Question
{
  const s = base(8, '“做一个AI问答助手”是完整需求吗？');
  quote(s, '业务方说：\n“能不能做一个AI助手，自动回答门店问题？”', 1.0, 1.9, 8.8, 2.35, C.red, 26);
  t(s, '？', 10.4, 1.72, 1.65, 1.65, { fontSize: 76, bold: true, color: C.red, align: 'center' });
  pill(s, '30秒思考', 10.35, 4.25, 1.75, C.red);
  t(s, '它描述的是问题、方案，还是两者的混合？', 2.0, 5.52, 9.2, 0.4, { fontSize: 17, bold: true, color: C.muted, align: 'center' });
  note(s, 8, '参考收口：这是一个方案方向，还不是完整需求。');
}

// 09 Five product questions
{
  const s = base(9, 'AI产品仍要先回答五个产品问题');
  const a = [['用户', '谁每天在使用？'], ['场景', '什么时候会问？'], ['问题', '当前为什么低效？'], ['价值', '改善什么结果？'], ['边界', '哪些不能自动答？']];
  flow(s, a, 1.9, 0.62, 12.0, [C.red, C.orange, C.teal, C.purple, C.green]);
  t(s, '先把产品问题说清楚，再讨论AI方案。', 1.2, 5.15, 10.9, 0.45, { fontSize: 19, bold: true, color: C.red, align: 'center' });
  note(s, 9);
}

// 10 Seven-step work map
{
  const s = base(10, '从一句需求到一个可用AI产品');
  const a = [['问题', '真实需求'], ['材料', '数据与知识'], ['可行性', '能力边界'], ['人机分工', '谁做什么'], ['效果标准', '怎样算可用'], ['失败案例', '问题出在哪'], ['持续优化', '怎样迭代']];
  flow(s, a, 1.85, 0.45, 12.42, [C.red, C.orange, C.teal, C.purple, C.green, C.red, C.blue]);
  t(s, 'AI产品经理负责的是完整闭环，不是其中一个工具节点。', 1.0, 5.2, 11.3, 0.45, { fontSize: 18, bold: true, color: C.red, align: 'center' });
  note(s, 10);
}

// 11 Demo vs validation
{
  const s = base(11, '一个Demo答对几个问题，不等于可以进入真实业务');
  card(s, 0.78, 1.72, 4.25, 4.6, '演示成功', '几个简单问题都回答通顺\n\n≠\n\n真实产品已经可用', C.orange, { headSize: 18, bodySize: 20, bodyColor: C.ink, valign: 'mid' });
  card(s, 5.35, 1.72, 6.75, 4.6, '真实验证', '验证材料｜历史咨询、有效规则、版本范围、模糊与高风险问题\n\n验证问题｜找对版本？说明依据？不确定时拒答？必要时转人工？', C.teal, { headSize: 18, bodySize: 14 });
  note(s, 11);
}

// 12 Human-AI lanes
{
  const s = base(12, 'AI不是默认独立完成所有任务');
  const rows = [['普通知识问题', 'AI直接回答', C.teal], ['活动与权益承诺', 'AI给候选答案，人确认', C.orange], ['退款／食品安全等高风险', '直接转人工处理', C.red], ['没有依据或信息不足', '拒绝编造，先追问', C.purple]];
  rows.forEach((d, i) => {
    const y = 1.68 + i * 1.12;
    label(s, d[0], 0.82, y, 3.15, d[2], d[2] === C.red ? C.softRed : d[2] === C.orange ? C.softOrange : d[2] === C.purple ? C.softPurple : C.softTeal);
    arrow(s, 4.23, y - 0.02, 0.45, C.line);
    card(s, 4.95, y - 0.15, 7.15, 0.72, d[1], '', d[2], { shadow: false, headSize: 14 });
  });
  t(s, '产品设计要明确：AI什么时候答、什么时候问、什么时候停。', 1.0, 6.15, 11.2, 0.35, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 12);
}

// 13 Metrics
{
  const s = base(13, '“回答很流畅”不是验收标准');
  const a = [['正确', '结论是否可信'], ['依据', '能否说明来源'], ['风险', '是否正确拒答／转人工'], ['效率', '是否缩短查找时间'], ['采用', '用户是否愿意继续使用']];
  flow(s, a, 1.9, 0.62, 12.0, [C.red, C.teal, C.orange, C.purple, C.green]);
  t(s, '指标没有统一阈值；标准要结合场景、风险和业务目标制定。', 1.1, 5.17, 11.1, 0.4, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 13);
}

// 14 Bad Case
{
  const s = base(14, '一条错误回答，问题一定出在模型吗？');
  card(s, 0.72, 1.72, 4.28, 4.55, '失败案例（Bad Case）', '员工问：“这张优惠券今天还能用吗？”\n\nAI回答：“可以使用。”\n\n实际：引用了上月已失效的活动规则', C.red, { headSize: 17, bodySize: 13 });
  const causes = [['资料', '过期规则没有下线', C.blue], ['检索', '没有找到有效版本', C.purple], ['模型', '生成内容与依据不一致', C.orange], ['交互', '没有确认适用范围', C.teal], ['流程', '高风险承诺没有人工确认', C.green]];
  causes.forEach((d, i) => card(s, 5.32 + (i % 2) * 3.55, 1.72 + Math.floor(i / 2) * 1.4, i === 4 ? 6.82 : 3.28, 1.06, d[0], d[1], d[2], { shadow: false, headSize: 13, bodySize: 10.5 }));
  t(s, '正确归因，才能找到正确的优化对象。', 5.35, 6.2, 6.7, 0.33, { fontSize: 14.5, bold: true, color: C.red });
  note(s, 14);
}

// 15 Ability map
{
  const s = base(15, 'AI产品经理的五类核心能力');
  const a = [['问题与业务判断', 'AI是否必要'], ['AI边界与可行性', '能力、数据、风险'], ['产品与人机协作', '流程、交互、兜底'], ['效果与Bad Case', '标准、评测、归因'], ['落地与持续优化', '协作、上线、迭代']];
  a.forEach((d, i) => row(s, i + 1, d[0], d[1], 1.68 + i * 0.91, [C.red, C.teal, C.purple, C.orange, C.green][i]));
  t(s, '本课程能力地图｜不是行业官方评分模型', 8.65, 0.84, 3.55, 0.24, { fontSize: 10, bold: true, color: C.red, align: 'right' });
  note(s, 15);
}

// 16 Definition
{
  const s = base(16, '本课程对AI产品经理的定义');
  label(s, '本课程定义', 0.85, 1.78, 1.45, C.red, C.softRed);
  quote(s, 'AI产品经理以真实问题为起点，\n利用AI设计可落地的产品方案，\n并对可行性、体验、效果和业务价值负责。', 0.85, 2.42, 11.35, 2.45, C.red, 25);
  const tags = [['可行性', C.teal, C.softTeal], ['体验', C.purple, C.softPurple], ['效果', C.orange, C.softOrange], ['业务价值', C.green, C.softGreen]];
  tags.forEach((d, i) => label(s, d[0], 2.1 + i * 2.45, 5.45, 1.8, d[1], d[2]));
  note(s, 16);
}

// 17 Definition exercise
{
  const s = base(17, '用一句话重新定义AI产品经理');
  quote(s, '面向【谁】，解决【什么问题】，\nAI完成【什么任务】，并对【什么结果】负责。', 0.85, 1.72, 7.55, 2.0, C.teal, 23);
  card(s, 8.75, 1.72, 3.45, 2.0, '不合格反例', '“会Prompt、RAG、Agent、向量库……”\n\n只有技术名词，没有用户、问题和结果。', C.red, { headSize: 14, bodySize: 11.5 });
  pill(s, '1分钟练习', 10.6, 0.82, 1.6, C.red);
  t(s, '先独立写，再邀请2名学员分享。', 1.0, 5.3, 11.0, 0.45, { fontSize: 17, bold: true, color: C.red, align: 'center' });
  note(s, 17);
}

// 18 Four jobs
{
  const s = base(18, '同样叫AI产品经理，可能在做四种完全不同的工作');
  grid(s, [['企业知识助手', '让员工更快找到规则与答案'], ['个人口语陪练', '让学习者持续练习并获得反馈'], ['AI建设平台', '让多个团队复用模型和能力'], ['客户PoC交付', '让方案在客户环境中验证与采用']], 4, 1.75, 3.75, [C.teal, C.purple, C.blue, C.orange]);
  t(s, '先看工作，不急着贴方向标签。', 1.4, 5.95, 10.5, 0.38, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 18);
}

// 19 Four questions
{
  const s = base(19, '读岗位，先问四个问题');
  grid(s, [['服务对象', '主要为谁创造价值？'], ['业务目标', '最终改善什么结果？'], ['产品形态', '应用、平台还是解决方案？'], ['交付方式', '标准产品还是客户项目？']], 2, 1.72, 2.05, [C.red, C.teal, C.purple, C.orange]);
  t(s, '四维判断法｜本课程教学框架', 8.72, 0.84, 3.5, 0.24, { fontSize: 10, bold: true, color: C.red, align: 'right' });
  note(s, 19);
}

// 20 Direction map
{
  const s = base(20, '本课的四方向岗位地图');
  card(s, 0.78, 1.72, 5.65, 1.75, '核心｜To B AI应用', '服务企业用户，改善业务流程、效率或经营结果', C.teal, { shadow: false });
  card(s, 6.72, 1.72, 5.65, 1.75, '核心｜To C AI应用', '服务个人用户，创造体验、留存和商业价值', C.purple, { shadow: false });
  card(s, 0.78, 3.78, 5.65, 1.75, '补充｜AI平台／能力', '经营通用能力与复用效率｜待真实JD验证', C.blue, { shadow: false, line: C.blue });
  card(s, 6.72, 3.78, 5.65, 1.75, '补充｜AI解决方案', '围绕客户问题完成试点、交付与采用｜待真实JD验证', C.orange, { shadow: false, line: C.orange });
  t(s, '不是官方、完整、互斥的分类；真实岗位可能交叉。', 1.05, 6.08, 11.0, 0.38, { fontSize: 15, bold: true, color: C.red, align: 'center' });
  note(s, 20);
}

function fourDimCase(n, titleText, scenario, values, accent, sourceExtra = '') {
  const s = base(n, titleText);
  card(s, 0.75, 1.7, 4.15, 4.55, '案例场景', scenario, accent, { headSize: 17, bodySize: 13 });
  const labels = ['服务对象', '业务目标', '产品形态', '交付方式'];
  values.forEach((v, i) => card(s, 5.25 + (i % 2) * 3.55, 1.7 + Math.floor(i / 2) * 2.0, 3.28, 1.7, labels[i], v, [C.red, C.teal, C.purple, C.orange][i], { shadow: false, headSize: 13, bodySize: 11.2 }));
  note(s, n, sourceExtra);
}
fourDimCase(21, 'To B案例：门店运营知识助手', '门店员工每天重复查询新品、活动、券规则和客诉流程。', ['企业员工', '流程提效与口径一致', '企业知识应用', '接入内部工作流程'], C.teal);

// 22 ToB priorities
{
  const s = base(22, 'To B真正难在企业流程，而不只是对话');
  const a = [['多角色', '员工、主管、总部'], ['业务流程', '问题如何进入工作链路'], ['投入产出', '效率、质量、经营结果'], ['权限安全', '谁能看什么、答什么'], ['采用推动', '为什么员工愿意使用']];
  a.forEach((d, i) => row(s, i + 1, d[0], d[1], 1.63 + i * 0.84, [C.teal, C.red, C.orange, C.purple, C.green][i], 0.72, 8.2));
  card(s, 9.18, 1.7, 3.05, 4.0, '真实JD验证区', 'JD-B1／B2待补\n\n补充岗位职责摘要与采集日期', C.teal, { shadow: false, line: C.teal, headSize: 14, bodySize: 11.5 });
  note(s, 22);
}

fourDimCase(23, 'To C案例：AI英语口语陪练', '个人学习者希望随时练口语，但真人陪练贵、开口压力大、反馈不连续。', ['个人学习者', '持续练习与能力提升', 'AI对话学习应用', '在线标准产品'], C.purple);

// 24 ToC journey
{
  const s = base(24, '对话自然，不等于用户会持续使用');
  flow(s, [['看见产品', ''], ['开始对话', ''], ['首次感知价值', 'Aha Moment'], ['继续练习', ''], ['长期使用／付费', '']], 1.85, 0.62, 9.0, [C.muted, C.purple, C.red, C.purple, C.green]);
  card(s, 9.92, 1.85, 2.35, 3.3, '首次价值时刻', 'AI能延续我的水平对话，并自然给出个性化纠正。', C.red, { shadow: false, headSize: 13.5, bodySize: 11.2 });
  label(s, '目标用户', 0.85, 4.95, 1.45, C.purple, C.softPurple); label(s, '具体场景', 2.52, 4.95, 1.45, C.purple, C.softPurple);
  label(s, '首次价值', 4.19, 4.95, 1.45, C.red, C.softRed); label(s, '留存转化', 5.86, 4.95, 1.45, C.green, C.softGreen);
  label(s, '内容安全', 7.53, 4.95, 1.45, C.orange, C.softOrange);
  t(s, '真实JD验证区｜JD-C1／C2待补', 9.55, 5.65, 2.7, 0.28, { fontSize: 10, bold: true, color: C.muted, align: 'right' });
  note(s, 24);
}

// 25 Platform pain
{
  const s = base(25, '为什么五个团队不应该重复建设五遍？');
  const teams = ['客服团队', '销售团队', '运营团队', '研发团队', 'HR团队'];
  teams.forEach((x, i) => label(s, x, 0.62, 1.65 + i * 0.83, 2.0, C.blue, C.softBlue));
  arrow(s, 2.95, 3.25, 0.6, C.line);
  card(s, 3.9, 1.72, 3.6, 4.1, '重复建设', '模型接入\n知识配置\n效果评测\n权限管理\n成本治理', C.red, { headSize: 18, bodySize: 15, bodyColor: C.ink });
  arrow(s, 7.86, 3.25, 0.6, C.line);
  card(s, 8.8, 1.72, 3.6, 4.1, '平台化问题', '哪些能力值得抽象？\n如何配置与复用？\n怎样降低接入和维护成本？', C.blue, { headSize: 18, bodySize: 13 });
  note(s, 25);
}

// 26 Platform profile
{
  const s = base(26, '平台产品经营的是通用能力和复用');
  pill(s, '待真实JD验证', 10.35, 0.82, 1.87, C.orange);
  const layers = [['业务团队', '客服／销售／运营等应用'], ['配置与运营', '知识、评测、权限、成本'], ['通用能力', '模型接入、工具与工作流'], ['基础设施', '模型、数据与系统资源']];
  layers.forEach((d, i) => {
    const y = 1.65 + i * 1.08;
    shape(s, S.roundRect, 0.78 + i * 0.45, y, 7.2 - i * 0.9, 0.78, [C.softRed, C.softOrange, C.softTeal, C.softBlue][i], C.line, { rectRadius: 0.04 });
    t(s, d[0] + '｜' + d[1], 1.0 + i * 0.45, y + 0.23, 6.75 - i * 0.9, 0.23, { fontSize: 12.5, bold: true, color: C.ink, align: 'center' });
  });
  card(s, 8.55, 1.65, 3.72, 4.55, '岗位判断重点', '平台用户是谁？\n经营通用能力还是具体应用？\n需要多深的系统抽象能力？\n接入效率怎样衡量？\n\nJD-P1／P2待补', C.blue, { shadow: false, bodySize: 12.2 });
  note(s, 26);
}

// 27 Solution need
{
  const s = base(27, '为什么标准产品仍需要客户方案？');
  card(s, 0.72, 1.75, 3.15, 4.2, '标准产品', '一套通用能力\n一套默认流程\n一套标准配置', C.teal, { headSize: 18, bodySize: 15 });
  arrow(s, 4.1, 3.15, 0.55, C.line);
  const clients = [['客户A', '流程不同'], ['客户B', '系统不同'], ['客户C', '权限与验收不同']];
  clients.forEach((d, i) => card(s, 4.98, 1.75 + i * 1.39, 3.05, 1.08, d[0], d[1], [C.red, C.orange, C.purple][i], { shadow: false, headSize: 13, bodySize: 10.5 }));
  arrow(s, 8.3, 3.15, 0.55, C.line);
  card(s, 9.18, 1.75, 3.15, 4.2, '解决方案问题', '怎样理解客户问题？\n如何做Demo或PoC？\n标准与定制边界在哪？\n怎样验收并推动采用？', C.orange, { headSize: 18, bodySize: 13 });
  note(s, 27);
}

// 28 Solution profile
{
  const s = base(28, '解决方案不只是做PPT');
  pill(s, '待真实JD验证', 10.35, 0.82, 1.87, C.orange);
  flow(s, [['客户问题', '调研与定义'], ['方案方向', '产品＋技术组合'], ['Demo／PoC', '验证关键假设'], ['范围边界', '标准与定制'], ['交付采用', '验收与使用']], 1.85, 0.62, 12.0, [C.red, C.teal, C.purple, C.orange, C.green]);
  quote(s, '责任不是“讲方案”，而是让方案在客户环境里被验证、交付和采用。', 1.0, 4.35, 10.9, 0.85, C.orange, 17);
  t(s, 'JD-S1／S2待补', 9.95, 5.75, 2.2, 0.25, { fontSize: 10, bold: true, color: C.muted, align: 'right' });
  note(s, 28);
}

// 29 Comparison table
{
  const s = base(29, '四个方向，工作重心有什么不同？');
  const cols = [0.72, 3.25, 5.88, 8.51, 11.14], widths = [2.3, 2.4, 2.4, 2.4, 1.4];
  const headers = ['方向', '首要关注', '典型结果', '可能优势背景', '证据'];
  headers.forEach((h, i) => { shape(s, S.rect, cols[i], 1.65, widths[i], 0.65, C.dark, C.dark); t(s, h, cols[i] + 0.08, 1.86, widths[i] - 0.16, 0.22, { fontSize: 10.5, bold: true, color: C.white, align: 'center' }); });
  const rows = [
    ['To B应用', '企业流程与采用', '效率／质量／经营改善', '行业、流程、B端产品', '核心'],
    ['To C应用', '用户体验与持续价值', '激活／留存／商业化', '用户、内容、增长', '核心'],
    ['平台／能力', '通用能力与复用', '接入效率与治理', '平台、中后台、技术协作', '待JD'],
    ['解决方案', '客户验证与交付', 'PoC／验收／采用', '客户、项目、行业交付', '待JD']
  ];
  rows.forEach((r, ri) => r.forEach((v, ci) => {
    const y = 2.3 + ri * 0.9, fill = ri % 2 ? C.paper : 'F0EBE2';
    shape(s, S.rect, cols[ci], y, widths[ci], 0.9, fill, C.line);
    t(s, v, cols[ci] + 0.09, y + 0.2, widths[ci] - 0.18, 0.5, { fontSize: ci === 0 ? 11.5 : 10, bold: ci === 0, color: ci === 4 && v === '待JD' ? C.orange : C.text, align: ci === 0 || ci === 4 ? 'center' : 'left', valign: 'mid' });
  }));
  note(s, 29);
}

// 30 Exercise
{
  const s = base(30, '这份岗位，更像哪一个方向？');
  card(s, 0.75, 1.68, 5.25, 4.65, '模拟岗位描述', '面向企业客户设计销售话术与跟进助手，连接CRM，参与客户试点并跟踪使用效果。', C.orange, { headSize: 17, bodySize: 16 });
  const q = [['服务谁？', ''], ['业务目标是什么？', ''], ['更像哪个方向？', ''], ['还缺什么信息？', '']];
  q.forEach((d, i) => card(s, 6.35 + (i % 2) * 2.95, 1.68 + Math.floor(i / 2) * 2.18, 2.72, 1.88, d[0], '请先写下你的判断', [C.red, C.teal, C.purple, C.orange][i], { shadow: false, headSize: 14, bodySize: 10.5 }));
  pill(s, '先讨论，再揭示参考分析', 9.55, 0.82, 2.68, C.red);
  note(s, 30, '参考分析：服务企业销售人员与管理者；目标可能是跟进效率、话术质量或转化；主体更接近To B应用，同时带解决方案属性；仍需确认标准产品还是客户项目、是否长期参与PoC和验收、定制程度。');
}

// 31 Evidence first
{
  const s = base(31, '不先问哪个最热门，先问我有什么证据');
  card(s, 0.82, 1.75, 4.4, 3.65, '热门导向', '“现在什么最火？”\n“哪个方向工资最高？”\n“大家都学Agent，我也去学。”', C.red, { headSize: 18, bodySize: 14 });
  t(s, '×', 5.52, 2.75, 1.0, 0.8, { fontSize: 44, bold: true, color: C.red, align: 'center' });
  card(s, 6.75, 1.75, 5.55, 3.65, '证据导向', '我已经有哪些可迁移资产？\n它们在哪类工作里更有价值？\n真实JD是否支持这个判断？', C.teal, { headSize: 18, bodySize: 14 });
  flow(s, [['资产证据', ''], ['方向假设', ''], ['JD验证', '']], 5.6, 3.1, 7.2, [C.teal, C.orange, C.green]);
  note(s, 31);
}

// 32 Six assets
{
  const s = base(32, '盘点六类可以迁移的经验资产');
  grid(s, [['行业资产', '理解行业结构、规则与角色'], ['用户资产', '长期接触并理解某类用户'], ['业务资产', '理解目标、流程与经营逻辑'], ['产品资产', '做过需求、方案、上线和复盘'], ['数据与技术协作', '与研发、数据、算法协作'], ['客户与项目交付', '调研、方案、推进、验收']], 3, 1.62, 1.86);
  t(s, '不是六项必须均衡，而是找出自己最有证据的组合。', 1.3, 5.83, 10.8, 0.4, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 32);
}

// 33 First three assets
{
  const s = base(33, '前三类资产：你理解谁、什么行业、哪条业务');
  const a = [
    ['行业资产', '不是“我在这个行业待过”', '而是能讲清结构、规则、角色与变化'],
    ['用户资产', '不是“我做过用户工作”', '而是长期观察同类用户的问题与行为'],
    ['业务资产', '不是“我熟悉公司业务”', '而是理解目标、流程、指标与关键约束']
  ];
  a.forEach((d, i) => {
    const x = 0.82 + i * 4.12;
    card(s, x, 1.72, 3.55, 4.35, d[0], d[1] + '\n\n→\n\n' + d[2], [C.red, C.teal, C.orange][i], { headSize: 18, bodySize: 13, bodyColor: C.ink });
  });
  note(s, 33);
}

// 34 Last three assets
{
  const s = base(34, '后三类资产：你真正负责、协作和交付过什么');
  const a = [
    ['产品资产', '不是“参与过产品”', '而是能说明需求、方案、取舍、上线和结果'],
    ['数据与技术协作', '不是“和研发关系好”', '而是处理过接口、数据、权限或技术边界'],
    ['客户与项目交付', '不是“见过客户”', '而是承担过调研、方案、推进、验收或采用']
  ];
  a.forEach((d, i) => {
    const x = 0.82 + i * 4.12;
    card(s, x, 1.72, 3.55, 4.35, d[0], d[1] + '\n\n→\n\n' + d[2], [C.purple, C.blue, C.green][i], { headSize: 17, bodySize: 12.8, bodyColor: C.ink });
  });
  note(s, 34);
}

// 35 Evidence levels
{
  const s = base(35, '不给自己打分，先看证据等级');
  const levels = [['强证据', '承担关键责任，能讲清行动、取舍和结果', C.green, 8.0], ['中证据', '深度参与并理解过程，但不是关键负责人', C.orange, 6.6], ['弱证据', '只有学习、体验、观察、兴趣或自我评价', C.red, 5.2]];
  levels.forEach((d, i) => {
    const x = 0.82 + i * 0.7, y = 1.72 + i * 1.2;
    shape(s, S.roundRect, x, y, d[3], 0.9, [C.softGreen, C.softOrange, C.softRed][i], C.line, { rectRadius: 0.04 });
    t(s, d[0], x + 0.25, y + 0.25, 1.35, 0.3, { fontSize: 15, bold: true, color: d[2] });
    t(s, d[1], x + 1.75, y + 0.24, d[3] - 2.0, 0.34, { fontSize: 12, color: C.text });
  });
  quote(s, '一条证据的基本结构：场景 → 责任 → 行动 → 结果', 1.1, 5.6, 10.8, 0.72, C.red, 17);
  note(s, 35);
}

function profileSlide(n, titleText, bg, evidence, direction, conclusion, accent) {
  const s = base(n, titleText);
  label(s, bg, 0.82, 1.65, 4.6, accent, accent === C.teal ? C.softTeal : C.softPurple);
  card(s, 0.82, 2.25, 5.35, 3.75, '现有证据', evidence, accent, { shadow: false, bodySize: 12.5 });
  card(s, 6.47, 2.25, 5.75, 3.75, '方向假设', direction, C.orange, { shadow: false, bodySize: 12.5 });
  t(s, conclusion, 1.0, 6.25, 11.2, 0.34, { fontSize: 12.5, bold: true, color: C.red, align: 'center' });
  note(s, n);
}
profileSlide(36, '示例：企业后台产品，为什么先验证To B？', '3年企业后台产品经验｜没有AI项目经历', '产品｜强：流程、后台、上线\n技术协作｜强：权限、接口、多系统\n业务｜中：理解一条核心流程\nAI项目｜弱：只有工具使用', '优先验证｜To B AI应用\n备选验证｜AI平台／能力\n最大缺口｜AI边界、效果评估、业务AI化项目\n\n平台方向仍待真实JD验证', '不是“后台产品＝To B”，而是证据组合使To B更值得优先验证。', C.teal);
profileSlide(37, '示例：内容运营，为什么先验证To C？', '3年内容运营和用户增长｜没有正式产品岗位经历', '用户｜强：长期服务特定内容用户\n业务｜强：内容供给、活动、增长复盘\n数据｜中：观察内容与增长指标\n产品｜弱：缺少方案和研发推进', '优先验证｜To C AI应用\n备选验证｜内容营销类To B应用\n最大缺口｜需求分析、方案、原型和项目推进证据', '不是“运营＝To C”，而是用户、内容和增长证据使To C更值得优先验证。', C.purple);

// 38 V0.1 form
{
  const s = base(38, '现在填写你的方向判断卡 V0.1');
  const a = [['背景事实', '岗位／行业／用户／流程'], ['经验资产', '最强两项证据'], ['方向假设', '优先方向＋备选方向'], ['判断依据', '为什么这样判断'], ['最大缺口', '当前只选一个'], ['个人定位句', '证据→方向→补缺口']];
  grid(s, a, 3, 1.6, 1.82);
  pill(s, '03:00', 11.15, 0.82, 1.07, C.red);
  t(s, '先写能够被事实证明的内容；不要求现在得到最终答案。', 1.0, 5.85, 11.2, 0.4, { fontSize: 15.5, bold: true, color: C.red, align: 'center' });
  note(s, 38);
}

// 39 Zero-basis comparison
{
  const s = base(39, '两个“产品零基础”，应该走同一条路吗？');
  card(s, 0.82, 1.72, 5.25, 4.5, '人物 A', '5年制造业大客户销售\n无产品岗位经历\n\n熟悉企业客户、销售流程和客户决策\n参与过方案沟通与项目推进', C.teal, { headSize: 20, bodySize: 14 });
  card(s, 6.25, 1.72, 5.25, 4.5, '人物 B', '刚毕业，没有正式工作经验\n无产品岗位经历\n\n使用过一些AI工具\n暂时没有稳定的行业、用户和项目证据', C.orange, { headSize: 20, bodySize: 14 });
  pill(s, '先让学员回答', 10.45, 0.82, 1.78, C.red);
  t(s, '他们都叫“产品零基础”，第一步应该完全相同吗？', 1.3, 6.3, 10.6, 0.38, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  note(s, 39, '参考收口：不相同。A已有行业、业务、客户与交付资产，重点补产品与AI项目证据；B先选择可接触用户的场景，创造第一份真实证据。');
}

// 40 Three starting points
{
  const s = base(40, '零基础的三种主要起点');
  const a = [['有产品，无AI', '保留产品基本功\n补AI边界、效果与项目闭环'], ['无产品，有行业／职能', '保留行业和用户资产\n补产品流程与推进证据'], ['三类积累都弱', '先选可接触用户的场景\n创造第一份真实证据']];
  a.forEach((d, i) => card(s, 0.82 + i * 4.12, 1.75, 3.55, 4.15, d[0], d[1], [C.teal, C.orange, C.red][i], { headSize: 18, bodySize: 14 }));
  t(s, '看主要缺口，不是给人贴标签｜本课程教学分型', 1.2, 6.18, 10.8, 0.35, { fontSize: 14, bold: true, color: C.red, align: 'center' });
  note(s, 40);
}

// 41 Two paths
{
  const s = base(41, '前两类：保留已有资产，补不同缺口');
  card(s, 0.78, 1.68, 5.65, 4.65, '路径 A｜有产品，无AI', '已有｜需求、方案、协作、上线、复盘\n\n重点补｜AI边界、人机分工、效果标准、Bad Case、持续优化\n\n本周动作｜选一条熟悉流程，收集10条脱敏问题并写出AI任务与效果标准', C.teal, { headSize: 18, bodySize: 12.5 });
  card(s, 6.72, 1.68, 5.65, 4.65, '路径 B｜无产品，有行业／职能', '已有｜行业、用户、业务、客户或交付经验\n\n重点补｜问题定义、产品流程、最小可行方案、原型、指标、迭代\n\n本周动作｜画一条熟悉流程，选一个节点并访谈3名真实从业者', C.orange, { headSize: 18, bodySize: 12.5 });
  note(s, 41);
}

// 42 Seven-step evidence path
{
  const s = base(42, '三类积累都弱：先创造第一份真实证据');
  const a = [['熟悉场景', ''], ['找到3—5人', '启动样本'], ['画出现状流程', ''], ['找到真实问题', ''], ['补最低AI基础', ''], ['做小Demo', ''], ['反馈＋JD验证', '']];
  flow(s, a, 1.85, 0.42, 12.48, [C.red, C.orange, C.teal, C.purple, C.blue, C.green, C.red]);
  t(s, '3—5人只是启动门槛，不是充分用户研究样本。', 1.0, 4.75, 11.2, 0.4, { fontSize: 16, bold: true, color: C.red, align: 'center' });
  quote(s, '第一目标不是“选对终身方向”，而是创造一份可被追问和修正的真实证据。', 1.0, 5.35, 11.0, 0.72, C.teal, 15.5);
  note(s, 42);
}

// 43 Break
{
  const s = base(43, '', '', { noHeader: true, noFooter: true, bg: C.dark });
  t(s, '休息 5 分钟', 0.82, 0.92, 5.5, 0.65, { fontSize: 34, bold: true, color: C.white });
  t(s, '回来后：3名同学现场方向诊断', 0.86, 1.85, 7.5, 0.4, { fontSize: 17, color: 'D9D4CC' });
  t(s, '05:00', 8.75, 0.9, 3.5, 1.0, { fontSize: 55, bold: true, color: 'E09A78', align: 'right' });
  const rules = [['不讲公司机密', '只使用明确授权和可公开信息'], ['不做就业承诺', '输出方向假设，不预测录用结果'], ['其他人同步填写', '把追问应用到自己的V0.2']];
  rules.forEach((d, i) => {
    rule(s, 0.85, 3.05 + i * 1.02, 11.4, '4A4B48');
    t(s, String(i + 1).padStart(2, '0'), 0.88, 3.28 + i * 1.02, 0.55, 0.22, { fontSize: 10, bold: true, color: 'E09A78' });
    t(s, d[0], 1.7, 3.22 + i * 1.02, 2.1, 0.32, { fontSize: 15, bold: true, color: C.white });
    t(s, d[1], 4.18, 3.25 + i * 1.02, 7.6, 0.3, { fontSize: 12.2, color: 'C7C4BE' });
  });
  note(s, 43);
}

function diagnosticSlide(n) {
  const s = base(n, `现场诊断 ${n - 43}：从事实到验证动作`);
  pill(s, '08:00', 11.15, 0.82, 1.07, C.red);
  const a = [['背景事实', '当前岗位／行业／服务对象'], ['关键证据', '只写两条可复述事实'], ['主要资产', '最强的证据组合'], ['方向假设', '优先／备选／暂不确定'], ['最大缺口', '当前只选一个'], ['双验证任务', 'JD任务＋项目／学习任务']];
  grid(s, a, 3, 1.6, 1.83, [C.red, C.teal, C.orange, C.purple, C.green, C.blue]);
  t(s, '事实 → 资产证据 → 方向假设 → 最大缺口 → JD＋项目验证', 1.0, 5.92, 11.2, 0.38, { fontSize: 14.5, bold: true, color: C.red, align: 'center' });
  note(s, n, '现场填写：不预先写死方向结论；若信息不足，填写“暂不确定＋需要补什么证据”。');
}
diagnosticSlide(44); diagnosticSlide(45); diagnosticSlide(46);

// 47 Common process
{
  const s = base(47, '三个人答案不同，判断过程相同');
  flow(s, [['背景事实', ''], ['资产证据', ''], ['方向假设', ''], ['最大缺口', ''], ['JD＋项目验证', '']], 1.85, 0.62, 12.0, [C.red, C.teal, C.orange, C.purple, C.green]);
  card(s, 0.9, 4.28, 3.6, 1.55, '方法提醒 01', '不从职位名称直接下结论', C.red, { shadow: false, headSize: 12.5, bodySize: 11 });
  card(s, 4.86, 4.28, 3.6, 1.55, '方法提醒 02', '不把兴趣或自评当成强证据', C.orange, { shadow: false, headSize: 12.5, bodySize: 11 });
  card(s, 8.82, 4.28, 3.6, 1.55, '方法提醒 03', '信息不足时不强行判断', C.teal, { shadow: false, headSize: 12.5, bodySize: 11 });
  t(s, 'V0.2补充｜1条关键追问 · 1条证据修正 · JD任务＋项目任务', 1.0, 6.15, 11.2, 0.35, { fontSize: 14, bold: true, color: C.red, align: 'center' });
  note(s, 47);
}

// 48 Conclusions
{
  const s = base(48, '今天带走三个结论');
  const a = [['工具不定义岗位', 'AI产品经理对问题、方案、效果和价值负责'], ['同名岗位工作不同', '要看服务对象、目标、形态和交付方式'], ['方向要通过行动验证', '从资产证据出发，用JD和项目持续修正']];
  a.forEach((d, i) => card(s, 0.82 + i * 4.12, 1.75, 3.55, 4.15, d[0], d[1], [C.red, C.teal, C.orange][i], { headSize: 18, bodySize: 13.5 }));
  note(s, 48);
}

// 49 Homework JD
{
  const s = base(49, '课后任务：用真实JD验证方向');
  const a = [['收集岗位', '普通作业5份\n求职学员建议10份'], ['逐份提取', '对象／目标／AI任务／职责／经历要求／能力要求／匹配证据／缺口'], ['总结共性', '哪些要求反复出现？\n我的假设需要怎样修正？']];
  a.forEach((d, i) => { card(s, 0.82 + i * 4.12, 1.72, 3.55, 4.25, d[0], d[1], [C.red, C.teal, C.orange][i], { headSize: 18, bodySize: i === 1 ? 11.4 : 13.2 }); if (i < 2) arrow(s, 4.48 + i * 4.12, 3.32, 0.35); });
  t(s, '不要用一份JD代表整个方向。', 1.4, 6.17, 10.5, 0.35, { fontSize: 15, bold: true, color: C.red, align: 'center' });
  note(s, 49);
}

// 50 V1.0
{
  const s = base(50, '你的最终提交：方向判断卡 V1.0');
  flow(s, [['V0.1', '课堂初判'], ['V0.2', '现场诊断修正'], ['V1.0', '真实JD验证']], 1.72, 2.0, 9.3, [C.red, C.orange, C.green]);
  const fields = ['背景与资产证据', '优先／备选方向', 'JD共性要求', '当前最大缺口', '下一步行动', '个人定位句'];
  fields.forEach((x, i) => label(s, x, 0.92 + (i % 3) * 4.05, 4.25 + Math.floor(i / 3) * 0.85, 3.55, [C.red, C.teal, C.orange, C.purple, C.green, C.blue][i], [C.softRed, C.softTeal, C.softOrange, C.softPurple, C.softGreen, C.softBlue][i]));
  t(s, 'V1.0仍然允许在后续项目和求职反馈中继续更新。', 1.0, 6.17, 11.2, 0.35, { fontSize: 14.5, bold: true, color: C.red, align: 'center' });
  note(s, 50);
}

// 51 Closing
{
  const s = base(51, '', '', { noHeader: true, noFooter: true, bg: C.bg });
  shape(s, S.rect, 0, 0, 0.22, 7.5, C.red, C.red);
  t(s, '方向不是一次选出来的，\n而是在行动中验证出来的。', 0.9, 1.25, 10.8, 1.25, { fontSize: 34, bold: true, color: C.ink, breakLine: false });
  rule(s, 0.92, 3.0, 3.0, C.red, 3);
  t(s, '暂时无法选方向？', 0.92, 3.55, 3.0, 0.32, { fontSize: 15, bold: true, color: C.red });
  t(s, '延续第42页最低验证路径，并分析3份相关JD。', 0.92, 4.05, 6.5, 0.42, { fontSize: 14, color: C.text });
  t(s, '下一课', 8.1, 3.55, 3.8, 0.3, { fontSize: 11, bold: true, color: C.muted, align: 'right', charSpacing: 1.2 });
  t(s, '大模型基础、Prompt\n与上下文设计', 7.0, 4.05, 4.9, 1.0, { fontSize: 21, bold: true, color: C.red, align: 'right', breakLine: false });
  pill(s, '116—120分钟：机动与集中答疑', 8.65, 6.42, 3.25, C.red);
  t(s, 'AI 产品经理系统课程 · 第一阶段', 0.92, 6.65, 4.2, 0.2, { fontSize: 9, color: C.muted });
  note(s, 51);
}

if (pptx._slides.length !== 51) throw new Error(`Expected 51 slides, got ${pptx._slides.length}`);
pptx.writeFile({ fileName: OUT });

