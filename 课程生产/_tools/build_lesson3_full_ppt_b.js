const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第03课：AI产品形态分类与场景判断';
pptx.title = 'AI产品形态分类与场景判断';
pptx.company = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };

const S = pptx.ShapeType;
const OUT = '课程生产/第一阶段/第03课_AI产品形态分类与场景判断_重做工作区/10_正式PPT_第03课_AI产品形态分类与场景判断.pptx';
const ROOT = '课程生产/第一阶段/第03课_AI产品形态分类与场景判断_重做工作区';
const C = {
  bg: 'F4F0E8', paper: 'FBF9F4', ink: '20211F', text: '292A27', muted: '77766F', line: 'D8D0C2', pale: 'ECE7DE', white: 'FFFFFF',
  red: 'C35331', teal: '387B77', purple: '74546A', orange: 'D18A38', green: '527553', blue: '405D78',
  softRed: 'F2E2DB', softTeal: 'DDEBE8', softPurple: 'E8E0E7', softOrange: 'F2E7D5', softGreen: 'E2E9DE', softBlue: 'DDE5EB'
};
const accentList = [C.red, C.teal, C.orange, C.purple, C.green, C.blue];
const softList = [C.softRed, C.softTeal, C.softOrange, C.softPurple, C.softGreen, C.softBlue];

const moduleOf = n => n <= 5 ? '课程任务' : n <= 12 ? '01 先判断是否需要AI' : n <= 20 ? '02 五维场景判断' : n <= 24 ? '03 两层产品形态图' : n === 25 ? '短休息' : n <= 30 ? '04 用户体验层' : n <= 40 ? '05 任务完成层' : n <= 43 ? '06 产品形态选择路径' : n <= 49 ? '07 综合案例' : '08 总结与课后作业';
const kicker = n => moduleOf(n).replace(/^\d+\s*/, '');

function t(slide, text, x, y, w, h, o = {}) {
  slide.addText(text, { x, y, w, h, margin: 0, fontFace: 'PingFang SC', fontSize: 16, color: C.text, fit: 'shrink', valign: 'top', breakLine: false, ...o });
}
function shape(slide, type, x, y, w, h, fill, line = fill, o = {}) {
  slide.addShape(type, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 1 }, ...o });
}
function line(slide, x, y, w, color = C.line, width = 1) {
  slide.addShape(S.line, { x, y, w, h: 0, line: { color, width } });
}
function pill(slide, text, x, y, w, color = C.red, fill = color, textColor = C.white) {
  shape(slide, S.roundRect, x, y, w, 0.36, fill, fill, { rectRadius: 0.08 });
  t(slide, text, x + 0.06, y + 0.08, w - 0.12, 0.18, { fontSize: 10, bold: true, color: textColor, align: 'center' });
}
function base(n, title, opts = {}) {
  const s = pptx.addSlide();
  s.background = { color: opts.bg || C.bg };
  if (!opts.noHeader) {
    pill(s, String(n).padStart(2, '0'), 0.62, 0.3, 0.62, opts.accent || C.red);
    t(s, kicker(n), 1.42, 0.38, 5.5, 0.18, { fontSize: 10, bold: true, color: C.muted, charSpacing: 1 });
    t(s, title, 0.62, 0.78, 12.0, 0.54, { fontSize: 25, bold: true, color: C.ink });
    line(s, 0.62, 1.42, 12.05);
  }
  if (!opts.noFooter) {
    t(s, 'AI 产品经理系统课程 · 第一阶段', 0.62, 7.06, 4.4, 0.16, { fontSize: 8, color: C.muted });
    t(s, moduleOf(n), 7.25, 7.06, 4.7, 0.16, { fontSize: 8, color: C.muted, align: 'right' });
    t(s, String(n).padStart(2, '0'), 12.1, 7.04, 0.55, 0.18, { fontSize: 8.5, color: C.muted, align: 'right' });
  }
  s.addNotes(`第${n}页｜${moduleOf(n)}\n逐页内容：${ROOT}/09_PPT结构与逐页内容\n老师讲解：${ROOT}/05_老师版讲义`);
  return s;
}
function card(slide, x, y, w, h, head, body = '', accent = C.red, o = {}) {
  shape(slide, S.roundRect, x, y, w, h, o.fill || C.paper, o.line || C.line, { rectRadius: 0.05, shadow: { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: 0.06 } });
  shape(slide, S.rect, x, y, 0.07, h, accent, accent);
  t(slide, head, x + 0.25, y + 0.22, w - 0.48, 0.42, { fontSize: o.headSize || 17, bold: true, color: C.ink, align: o.align || 'left' });
  if (body) t(slide, body, x + 0.25, y + 0.82, w - 0.5, h - 1.03, { fontSize: o.bodySize || 13.5, color: o.bodyColor || C.muted, valign: o.valign || 'top', paraSpaceAfterPt: 7, breakLine: false, align: o.bodyAlign || 'left' });
}
function quote(slide, text, y = 2.25, accent = C.red, size = 27) {
  shape(slide, S.rect, 1.0, y, 0.08, 1.25, accent, accent);
  t(slide, text, 1.4, y + 0.03, 10.8, 1.18, { fontSize: size, bold: true, color: C.ink, valign: 'mid', align: 'center' });
}
function cards(slide, items, cols = 3, y = 1.75, h = 1.75) {
  const gap = 0.24, x0 = 0.72, total = 11.9, w = (total - gap * (cols - 1)) / cols;
  items.forEach((it, i) => {
    const x = x0 + (i % cols) * (w + gap), yy = y + Math.floor(i / cols) * (h + 0.24);
    card(slide, x, yy, w, h, it[0], it[1] || '', it[2] || accentList[i % 6], { headSize: it[3] || 16, bodySize: it[4] || 13, bodyAlign: it[5] || 'left' });
  });
}
function flow(slide, items, y = 2.0, o = {}) {
  const x0 = o.x || 0.66, total = o.w || 12.0, gap = o.gap || 0.18, h = o.h || 1.65, w = (total - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const x = x0 + i * (w + gap), a = it[2] || accentList[i % 6];
    shape(slide, S.roundRect, x, y, w, h, C.paper, C.line, { rectRadius: 0.05 });
    t(slide, String(i + 1).padStart(2, '0'), x + 0.15, y + 0.17, 0.45, 0.2, { fontSize: 10, bold: true, color: a });
    t(slide, it[0], x + 0.14, y + 0.55, w - 0.28, 0.44, { fontSize: o.headSize || 15, bold: true, color: C.ink, align: 'center', valign: 'mid' });
    if (it[1]) t(slide, it[1], x + 0.15, y + 1.1, w - 0.3, h - 1.23, { fontSize: o.bodySize || 10.5, color: C.muted, align: 'center', valign: 'mid' });
  });
}
function twoCol(slide, left, right, opts = {}) {
  card(slide, 0.78, 1.72, 5.78, opts.h || 4.75, left[0], left[1], left[2] || C.teal, { headSize: left[3] || 19, bodySize: left[4] || 14, valign: opts.valign || 'top' });
  card(slide, 6.78, 1.72, 5.78, opts.h || 4.75, right[0], right[1], right[2] || C.orange, { headSize: right[3] || 19, bodySize: right[4] || 14, valign: opts.valign || 'top' });
}
function banner(slide, text, y = 5.82, accent = C.red) {
  shape(slide, S.roundRect, 1.05, y, 11.2, 0.58, softList[accentList.indexOf(accent)] || C.softRed, softList[accentList.indexOf(accent)] || C.softRed, { rectRadius: 0.05 });
  t(slide, text, 1.25, y + 0.14, 10.8, 0.28, { fontSize: 15, bold: true, color: accent, align: 'center' });
}
function bullets(lines) { return lines.map(x => `• ${x}`).join('\n'); }

// 01 封面
{
  const s = base(1, '', { noHeader: true, noFooter: true });
  shape(s, S.rect, 0, 0, 0.22, 7.5, C.red, C.red);
  t(s, '03', 0.88, 0.72, 1.1, 0.66, { fontSize: 35, bold: true, color: C.red });
  t(s, 'AI产品形态分类与\n场景判断', 0.88, 1.58, 8.3, 1.55, { fontSize: 38, bold: true, color: C.ink });
  t(s, '先判断是否需要AI，再选择能够落地的产品组合', 0.92, 3.52, 8.4, 0.42, { fontSize: 17, color: C.muted });
  line(s, 0.92, 4.18, 3.0, C.red, 3);
  ['是否AI', '体验层', '完成层', '验证'].forEach((x, i) => { pill(s, x, 8.8, 1.15 + i * 1.05, 2.4, accentList[i], softList[i], accentList[i]); });
  pill(s, '第一阶段 · 第03课', 9.66, 6.43, 2.55, C.red);
}

// 02—05 开场
{
  const s = base(2, '今天只解决四个问题');
  cards(s, [['什么任务适合AI？', '先与规则、系统和人工比较'], ['怎样判断场景成立？', '价值、任务、数据、风险、成本'], ['为什么不是五选一？', '体验层与完成层分开'], ['怎样写出专业方案？', '选择＋不选＋边界＋验证']], 2, 1.72, 2.05);
  banner(s, '不讲“哪种技术最先进”，只讲“为什么这样选”。');
}
{
  const s = base(3, '你会带走一张选择卡');
  twoCol(s, ['V0.1｜先判断', bullets(['五维场景判断', '适合／局部适合／暂不适合', '最大待验证项']), C.blue], ['V0.2｜再选方案', bullets(['用户体验层', '任务完成层', '不选方案与人工边界']), C.purple]);
  banner(s, '课后换成自己的工作任务，完成V1.0。', 6.05, C.red);
}
{
  const s = base(4, '需求来了，你先选哪个？');
  t(s, '每天把最新活动规则同步给商家，并提醒需要调整的门店。', 1.0, 1.72, 11.3, 0.6, { fontSize: 20, bold: true, color: C.ink, align: 'center' });
  cards(s, [['A｜做Chatbot', '让商家自己问'], ['B｜做Agent', '自动处理一切'], ['C｜先拆任务', '拆数据、动作和风险，再决定组合']], 3, 2.65, 2.4);
  banner(s, '先举手投票｜本页不公布“标准技术答案”', 5.55, C.orange);
}
{
  const s = base(5, '第2课解决Prompt，本课解决方案选择');
  twoCol(s, ['第2课', '怎样把一个任务说清楚\n怎样验证模型输出\n怎样识别Prompt边界', C.blue], ['第3课', '这个任务是否值得用AI\n用户应该怎样使用\n系统依靠什么完成', C.red]);
  banner(s, 'Prompt是一块积木，不是整个产品方案。');
}

// 06—12 是否需要AI
{
  const s = base(6, '能生成，不等于值得做成产品');
  twoCol(s, ['Demo问题', '模型能不能给出一个结果？', C.teal, 21, 20], ['产品问题', '价值、数据、风险、成本是否成立？', C.red, 21, 20], { valign: 'mid' });
  banner(s, '始终与规则系统、数据库、原有功能和人工方案比较。');
}
{
  const s = base(7, '大模型更擅长写、读、判、答');
  cards(s, [['写', '生成｜改写', C.red, 24, 17, 'center'], ['读', '总结｜抽取', C.teal, 24, 17, 'center'], ['判', '分类｜匹配', C.orange, 24, 17, 'center'], ['答', '依据上下文回答', C.purple, 24, 17, 'center']], 4, 2.0, 2.6);
  banner(s, '这是能力范围，不是立项理由。', 5.35, C.red);
}
{
  const s = base(8, '三类常见机会特征');
  cards(s, [['重认知', '需要理解大量非结构化信息'], ['可标准化', '输出不唯一，但有基本标准'], ['信息过载', '人工处理速度跟不上增长']], 3, 1.9, 3.0);
  banner(s, '有价值、可验证，才值得继续。', 5.45, C.red);
}
{
  const s = base(9, '固定规则：先用确定性系统');
  quote(s, '金额 × 固定费率，超过上限则取上限', 1.85, C.blue, 25);
  cards(s, [['规则系统', '准确｜便宜｜可审计', C.teal], ['生成式AI', '可能增加不必要的不确定性', C.orange]], 2, 3.75, 1.65);
  banner(s, '能用明确公式稳定解决，不必为了AI而AI。', 5.75, C.red);
}
{
  const s = base(10, '实时事实：先查数据库');
  flow(s, [['用户问题', '昨天GMV？'], ['意图识别', '确定查询条件'], ['数据库', '返回真实数据'], ['模型解释', '组织成人话']], 2.0, { headSize: 16 });
  banner(s, '模型可以解释结果，但不能根据经验生成一个数字。', 4.9, C.red);
}
{
  const s = base(11, '一个流程可以只在局部使用AI');
  flow(s, [['AI理解', '咨询意图'], ['系统查询', '规则／数据'], ['AI组织', '回复草稿'], ['人工确认', '高风险项'], ['系统执行', '确定动作']], 1.92, { headSize: 14, bodySize: 10 });
  banner(s, 'AI是流程中的角色，不必吞掉整个流程。', 4.85, C.red);
}
{
  const s = base(12, '六张任务卡：AI、系统、人怎么分工？');
  cards(s, [['申诉分类', '自然语言→问题队列'], ['费率计算', '明确公式，必须精确'], ['规则问答', '依据更新规则回答'], ['昨日GMV', '查询真实经营数据'], ['活动文案', '生成多个可编辑初稿'], ['高额退款', '资金风险，缺稳定验收']], 3, 1.65, 1.65);
  banner(s, '30秒讨论｜不要只回答“用AI／不用AI”', 5.42, C.orange);
}

// 13—20 五维判断
{
  const s = base(13, '五维不是打分，是找成立条件');
  cards(s, [['业务价值', '谁的问题？'], ['任务匹配', '模型是否擅长？'], ['数据条件', '材料是否可用？'], ['风险与验收', '错了怎么办？'], ['成本收益', '值得投入吗？']], 5, 2.1, 2.4);
  banner(s, '任何一个关键条件不成立，都可能改变方案。', 5.2, C.red);
}
{
  const s = base(14, '业务价值：谁的问题，值多少钱');
  cards(s, [['谁的问题？', '明确用户'], ['多高频？', '不是偶发想象'], ['损失什么？', '时间／成本／机会'], ['看什么指标？', '改善必须可衡量']], 4, 1.95, 2.55);
  banner(s, '反例：“大家都在做，所以我们也需要。”', 5.3, C.orange);
}
{
  const s = base(15, '任务匹配：输入、输出、评价能否说清');
  flow(s, [['真实输入', '用户实际提供什么'], ['处理动作', '理解／生成／判断'], ['可检查输出', '怎样算完成']], 2.1, { x: 1.0, w: 11.3, h: 2.05, headSize: 18, bodySize: 12 });
  banner(s, '“提升商家经营效率”不是一个可验证任务。', 5.15, C.orange);
}
{
  const s = base(16, '数据条件：知识、事实、权限、更新');
  cards(s, [['输入材料', '用户当次提供'], ['规则知识', '制度／手册／口径'], ['实时事实', '数据库／业务系统'], ['系统接口', '查询或执行动作']], 4, 1.9, 2.55);
  banner(s, '每一类都追问：在哪、谁有权限、多久更新、质量怎样？', 5.28, C.red);
}
{
  const s = base(17, '风险与验收：错了能否发现和恢复');
  flow(s, [['发现', '怎样识别错误'], ['复核', '谁有能力确认'], ['恢复', '能否撤回或重试'], ['留痕', '证据与动作可追踪']], 2.0, { headSize: 17 });
  banner(s, '风险越高，越需要证据、权限、确认和兜底。', 5.0, C.red);
}
{
  const s = base(18, '成本收益：模型费只是小部分');
  shape(s, S.trapezoid, 4.05, 1.75, 5.2, 1.25, C.softBlue, C.blue);
  t(s, '模型调用费', 4.25, 2.13, 4.8, 0.34, { fontSize: 20, bold: true, color: C.blue, align: 'center' });
  shape(s, S.trapezoid, 2.3, 3.18, 8.7, 2.35, C.softRed, C.red);
  t(s, '数据整理｜系统接入｜评测｜运营维护｜人工复核', 2.7, 4.05, 7.9, 0.54, { fontSize: 19, bold: true, color: C.red, align: 'center' });
  banner(s, '与当前人工或系统成本比较，而不是只算Token。', 5.85, C.red);
}
{
  const s = base(19, '三类结论：适合、局部适合、暂不适合');
  cards(s, [['适合AI', '关键条件基本成立', C.green], ['局部适合AI', '只让AI承担可验证的一段', C.orange], ['暂不适合AI', '仍有关键价值、数据或风险缺口', C.red]], 3, 1.95, 3.0);
  banner(s, '“暂不适合”也是专业结论。', 5.48, C.red);
}
{
  const s = base(20, 'V0.1：先判断，不选技术');
  quote(s, '根据商家经营情况给出活动建议，并在确认后创建活动草稿。', 1.72, C.blue, 23);
  cards(s, [['5分钟', '每个维度写一条已知、一条待验证', C.orange], ['只写判断', '适合／局部适合／暂不适合', C.teal], ['最后一句', '主要依据＋最大待验证项', C.purple]], 3, 3.55, 1.95);
  banner(s, '此时不要写Chatbot、RAG或Agent。', 5.8, C.red);
}

// 21—25 两层与休息
{
  const s = base(21, '“Chatbot还是RAG”为什么问错了');
  quote(s, '这类似于问：“门店还是库存系统？”', 1.75, C.red, 27);
  twoCol(s, ['Chatbot', '回答：用户怎样使用', C.teal, 21, 18], ['RAG', '回答：系统怎样找到知识', C.purple, 21, 18], { h: 2.25, valign: 'mid' });
  banner(s, '两者可以组合，不是同层二选一。', 5.75, C.red);
}
{
  const s = base(22, '两层产品形态图');
  t(s, '用户体验层｜用户怎样使用', 0.82, 1.68, 2.8, 0.35, { fontSize: 16, bold: true, color: C.teal });
  cards(s, [['Chatbot', '主动提问'], ['Copilot', '嵌入原页面'], ['后台处理', '事件／队列触发']], 3, 2.08, 1.35);
  t(s, '任务完成层｜系统怎样完成', 0.82, 3.86, 3.2, 0.35, { fontSize: 16, bold: true, color: C.purple });
  cards(s, [['Prompt', '直接生成'], ['RAG', '特定知识'], ['数据／工具', '事实与动作'], ['Workflow', '固定路径'], ['Agent', '动态路径']], 5, 4.28, 1.35);
  banner(s, '先选用户入口，再按任务组合后台积木。', 5.95, C.red);
}
{
  const s = base(23, '四种常见组合');
  cards(s, [['规则咨询', 'Chatbot ＋ RAG'], ['文案创作', 'Copilot ＋ Prompt'], ['材料初审', '后台处理 ＋ Workflow'], ['异常排查', 'Copilot ＋ 数据工具 ＋ 受控流程']], 2, 1.68, 2.1);
}
{
  const s = base(24, '同一入口，不等于同一后台机制');
  card(s, 0.72, 2.25, 2.45, 2.2, '商家经营助手', '一个统一入口', C.red, { headSize: 19, bodySize: 16, bodyAlign: 'center', valign: 'mid' });
  const ys = [1.55, 2.55, 3.55, 4.55];
  [['规则知识', 'RAG'], ['实时数据', '数据库API'], ['文案生成', 'Prompt'], ['复杂排查', 'Workflow／Agent验证']].forEach((d, i) => {
    shape(s, S.chevron, 3.48, ys[i] + 0.3, 0.55, 0.48, C.line, C.line);
    shape(s, S.roundRect, 4.25, ys[i], 7.8, 0.9, C.paper, C.line, { rectRadius: 0.05, shadow: { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: 0.05 } });
    shape(s, S.rect, 4.25, ys[i], 0.07, 0.9, accentList[i], accentList[i]);
    t(s, d[0], 4.5, ys[i] + 0.25, 2.2, 0.34, { fontSize: 16, bold: true, color: C.ink });
    t(s, d[1], 7.2, ys[i] + 0.25, 4.45, 0.34, { fontSize: 14, bold: true, color: accentList[i], align: 'right' });
  });
  banner(s, '入口可以统一，任务机制必须按问题选择。', 5.88, C.red);
}
{
  const s = base(25, '', { noHeader: true, noFooter: true, bg: C.softBlue });
  t(s, '休息 5 分钟', 1.0, 1.55, 11.3, 0.75, { fontSize: 38, bold: true, color: C.blue, align: 'center' });
  t(s, '回来后，先选体验，再选机制', 1.0, 2.72, 11.3, 0.55, { fontSize: 24, bold: true, color: C.ink, align: 'center' });
  pill(s, '保存 V0.1', 4.1, 4.15, 2.2, C.blue);
  pill(s, '打开案例任务卡', 6.75, 4.15, 2.45, C.teal);
  t(s, '05:00', 1.0, 5.65, 11.3, 0.6, { fontSize: 30, bold: true, color: C.muted, align: 'center' });
}

// 26—30 体验层
{
  const s = base(26, 'Chatbot：适合主动提问与澄清');
  twoCol(s, ['适用条件', bullets(['用户有明确问题', '表达不固定', '追问澄清有价值']), C.teal], ['最小闭环', '提问 → 必要澄清 → 有依据回答 → 无依据转人工', C.blue, 19, 16]);
  banner(s, '例：商家询问当前有效活动规则。');
}
{
  const s = base(27, 'Chatbot不是所有任务的默认入口');
  twoCol(s, ['适合对话', '规则咨询\n开放探索\n需要多轮澄清', C.green], ['强行对话会增加步骤', '批量材料审核\n固定表单处理\n后台队列任务', C.red]);
  banner(s, '对话如果增加步骤，就不是更自然的体验。', 6.0, C.red);
}
{
  const s = base(28, 'Copilot：在原工作流里帮一把');
  card(s, 0.78, 1.72, 7.1, 4.72, '活动创建页', '门店特色：社区亲子餐厅\n活动目标：拉新\n\nAI建议：生成标题、卖点与短文案\n\n用户操作：编辑｜采用｜拒绝｜提交', C.blue, { headSize: 20, bodySize: 15 });
  card(s, 8.18, 1.72, 4.2, 4.72, 'Copilot关键', bullets(['利用当前页面上下文', '在需要时出现', '不夺走用户控制权']), C.teal, { headSize: 19, bodySize: 15 });
}
{
  const s = base(29, '后台自动处理：用户不用一直对话');
  flow(s, [['事件触发', '材料提交'], ['进入队列', '等待处理'], ['解析检查', 'AI＋规则'], ['结果分流', '成功／待人工'], ['失败处理', '重试／留痕']], 1.95, { headSize: 14, bodySize: 10 });
  banner(s, '看不见AI，不代表不需要状态、审计和兜底。', 4.85, C.red);
}
{
  const s = base(30, '体验层选择：看用户当时在哪里');
  cards(s, [['会主动提问？', '是 → Chatbot'], ['已在工作页面？', '是 → Copilot'], ['无需持续参与？', '是 → 后台自动处理']], 3, 1.95, 3.0);
  banner(s, '可以组合，但必须有明确触发时点。', 5.45, C.red);
}

// 31—40 完成层
{
  const s = base(31, '系统完成任务的五类积木');
  cards(s, [['Prompt', '直接生成'], ['RAG', '特定知识'], ['数据／工具', '事实与动作'], ['Workflow', '固定路径'], ['Agent', '动态路径']], 5, 2.0, 2.7);
  banner(s, '不是能力排行榜，而是按任务组合的积木。', 5.35, C.red);
}
{
  const s = base(32, 'Prompt：输入基本齐全时直接完成');
  flow(s, [['输入', '门店特色＋目标'], ['模型', '依据指令生成'], ['输出', '三个可编辑文案']], 2.0, { x: 1.0, w: 11.3, h: 2.0, headSize: 18, bodySize: 13 });
  banner(s, 'Prompt不能补齐缺失的实时数据、内部知识和系统权限。', 5.08, C.red);
}
{
  const s = base(33, 'RAG：答案必须依据特定知识');
  flow(s, [['问题', '自然语言咨询'], ['检索', '找到有效规则'], ['证据', '交给模型'], ['回答', '依据证据组织']], 1.95, { headSize: 17 });
  banner(s, '无证据、版本冲突或高风险解释时转人工。', 4.95, C.red);
}
{
  const s = base(34, '实时数据与动作：数据库、API、工具');
  twoCol(s, ['事实查询', '订单｜GMV｜库存\n由业务系统返回真实数据', C.blue], ['受控动作', '创建草稿｜发起工单｜发送通知\n接口校验权限并执行', C.orange]);
  banner(s, '模型理解意图和解释结果；系统提供事实和执行动作。');
}
{
  const s = base(35, '知识和事实不能混为一谈');
  twoCol(s, ['规则知识', '“活动门槛是什么？”\n\nRAG检索当前有效规则', C.purple], ['业务事实', '“我的门店达到门槛了吗？”\n\n数据库／API查询真实数据', C.blue]);
  banner(s, '模型结合两条通道解释｜不要把实时数据当普通知识文档。', 6.0, C.red);
}
{
  const s = base(36, 'Workflow：路径固定，步骤可预设');
  flow(s, [['接收材料', ''], ['AI提取', '字段信息'], ['规则校验', '固定条件'], ['待补清单', ''], ['人工复核', '']], 1.92, { headSize: 14, bodySize: 10 });
  banner(s, '可追踪、可重试、可插人工、容易验收。', 4.85, C.green);
}
{
  const s = base(37, 'Agent：路径会随结果动态变化');
  card(s, 0.78, 1.72, 3.1, 4.75, '目标', '排查活动效果异常', C.red, { headSize: 20, bodySize: 19, bodyAlign: 'center', valign: 'mid' });
  [['查数据', '指标异常在哪里'], ['查规则', '活动条件是否变化'], ['查工单', '是否有服务问题'], ['继续追问', '补充缺失上下文']].forEach((d, i) => {
    const x = 4.18 + (i % 2) * 4.1, y = 1.72 + Math.floor(i / 2) * 2.36;
    card(s, x, y, 3.82, 2.12, d[0], d[1], accentList[i], { headSize: 18, bodySize: 14 });
  });
  banner(s, '灵活性更高，也更难预测、测试和治理。', 6.05, C.red);
}
{
  const s = base(38, 'Workflow还是Agent？先看路径是否固定');
  twoCol(s, ['Workflow', bullets(['步骤预先设计', '确定性较高', '容易测试和审计', '固定任务第一版优先']), C.teal], ['Agent', bullets(['步骤动态选择', '确定性较低', '测试治理更复杂', '动态多工具才验证']), C.orange]);
}
{
  const s = base(39, 'Multi-Agent不是复杂任务的默认答案');
  quote(s, '只有职责确实需要独立分工、隔离或并行协作时再讨论。', 1.82, C.purple, 24);
  cards(s, [['协调成本', '谁把任务交给谁'], ['状态一致性', '信息是否同步'], ['错误定位', '哪一环出了问题'], ['时延与费用', '调用链更长']], 4, 3.75, 1.65);
  banner(s, '本课只掌握边界，不展开架构实现。', 5.78, C.red);
}
{
  const s = base(40, '五个常见误判');
  cards(s, [['01 数据都放进RAG', '事实与知识混淆'], ['02 Agent一定更先进', '忽略确定性与成本'], ['03 Chatbot承载所有任务', '忽略用户工作位置'], ['04 Demo跑通即可上线', '忽略验收与治理'], ['05 模型直接执行高风险动作', '缺权限与人工确认']], 3, 1.58, 1.65);
  banner(s, '逐项回答：错在哪里？应该补什么？', 5.5, C.orange);
}

// 41—43 选择路径
{
  const s = base(41, '六问把场景推导成产品组合');
  flow(s, [['需要AI？', '与基线比较'], ['怎样使用？', '体验层'], ['特定知识？', 'RAG'], ['事实／动作？', '数据与工具'], ['固定／动态？', 'Workflow／Agent'], ['错了怎么办？', '确认与兜底']], 1.9, { x: 0.42, w: 12.5, h: 2.0, headSize: 13.5, bodySize: 9.5 });
  banner(s, '输出：体验层＋完成层＋人工边界＋验证计划。', 4.95, C.red);
}
{
  const s = base(42, '一个专业方案要说完整六件事');
  cards(s, [['任务是什么', '具体到用户与时点'], ['AI适配度', '适合／局部／暂不'], ['体验层', '用户怎样使用'], ['完成层', '系统依靠什么'], ['不选什么', '理由是什么'], ['风险与验证', '确认、兜底、指标']], 3, 1.65, 1.65);
  banner(s, '专业感来自每个选择都能回到任务证据。', 5.45, C.red);
}
{
  const s = base(43, '微练习：材料初审为什么不先做Agent');
  quote(s, '字段固定｜审核规则可列｜最终审核需要人工', 1.72, C.blue, 25);
  cards(s, [['体验层？', '后台自动处理'], ['完成层？', 'Workflow：AI提取＋规则判断'], ['不选什么？', '暂不先做Agent']], 3, 3.45, 2.0);
  banner(s, '异常和最终审核转人工。', 5.82, C.red);
}

// 44—49 案例
{
  const s = base(44, '邻里增长平台：项目背景');
  card(s, 0.78, 1.72, 4.2, 4.7, '业务', '为本地生活商家提供\n活动｜素材｜经营数据｜服务工单', C.blue, { headSize: 20, bodySize: 17, bodyAlign: 'center', valign: 'mid' });
  card(s, 5.28, 1.72, 7.1, 2.15, '当前问题', '规则多、页面多、数据难理解；平台运营重复处理量大。', C.orange, { headSize: 19, bodySize: 16 });
  card(s, 5.28, 4.15, 7.1, 2.27, '开放问题', '如果增加AI，应该从哪些任务开始？', C.red, { headSize: 19, bodySize: 20, bodyAlign: 'center', valign: 'mid' });
  pill(s, '教学虚构｜不对应真实公司或项目', 8.52, 0.82, 3.5, C.purple, C.softPurple, C.purple);
}
{
  const s = base(45, '现有系统与约束');
  cards(s, [['规则文档', '每周可能更新'], ['经营数据', '保存在业务数据库'], ['活动创建', '已有接口，发布需确认'], ['入驻审核', '规则明确，部分需人工'], ['历史样本', '可申请脱敏记录']], 5, 1.95, 2.65);
  banner(s, '哪些是知识？哪些是事实？哪些是动作？', 5.25, C.orange);
}
{
  const s = base(46, '五个任务，不会是同一种方案');
  cards(s, [['A 规则问答', '依据当前有效规则'], ['B 活动文案', '生成可编辑初稿'], ['C 材料初审', '提取、校验、待补'], ['D 数据查询解释', '先取真实数据'], ['E 异常排查', '跨页面动态排查']], 5, 1.95, 2.7);
  banner(s, '先独立写初判｜不要提前翻到答案页', 5.35, C.orange);
}
{
  const s = base(47, 'V0.2：选一个任务完成组合判断');
  cards(s, [['01 AI适配度', '适合／局部／暂不'], ['02 体验层', '用户怎样使用'], ['03 完成层', '系统依靠什么'], ['04 不选什么', '写出理由'], ['05 边界与验证', '谁确认、先验证什么']], 5, 1.75, 2.35);
  banner(s, '8分钟｜抽2—3组｜每组60秒', 4.85, C.orange);
  t(s, '汇报模板：任务是______，当前属于______；选择______＋______；暂不选______；最大风险______。', 1.0, 5.72, 11.3, 0.55, { fontSize: 14, bold: true, color: C.ink, align: 'center' });
}
{
  const s = base(48, '讨论后揭示：A—C参考拆解');
  cards(s, [['A 规则问答', '适合\nChatbot＋RAG\n无证据转人工', C.purple], ['B 活动文案', '适合辅助\nCopilot＋Prompt\n商家编辑确认', C.teal], ['C 材料初审', '局部适合\n后台处理＋Workflow\n异常与终审转人工', C.orange]], 3, 1.85, 3.35);
  banner(s, '体验由用户工作位置决定，机制由任务条件决定。', 5.55, C.red);
}
{
  const s = base(49, '讨论后揭示：D—E参考拆解');
  twoCol(s, ['D 数据查询解释', '局部适合\nChatbot／Copilot\n＋数据库API＋分析模板\n\n推测必须标记，重要判断确认', C.blue], ['E 复杂异常排查', '待验证\nCopilot＋受控Workflow\n路径确实动态再验证Agent\n\n预算、活动修改与对外动作前确认', C.red]);
  banner(s, '不是看到“复杂”就直接选择Agent。', 6.02, C.red);
}

// 50 总结
{
  const s = base(50, '四句话总结＋课后V1.0');
  card(s, 0.72, 1.62, 7.15, 4.95, '四句话', '1  能力成立，不等于产品价值成立\n\n2  先判断是否需要AI，再选择产品形态\n\n3  体验层与任务完成层必须分开\n\n4  写清不选什么、谁确认、怎样验证', C.red, { headSize: 20, bodySize: 16 });
  card(s, 8.17, 1.62, 4.45, 4.95, '课后V1.0', '从自己的工作中选择一个具体任务\n\n必须包含：\n基线比较｜五维证据\n产品组合｜人工边界\n最小验证\n\n不提交个人信息、公司机密或未脱敏数据', C.blue, { headSize: 20, bodySize: 14.5 });
}

if (pptx._slides.length !== 50) throw new Error(`Expected 50 slides, got ${pptx._slides.length}`);
pptx.writeFile({ fileName: OUT });
