const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第03课：Workflow类AI产品设计与实操';
pptx.title = 'Workflow类AI产品设计：从一次模型调用到可运行流程';
pptx.company = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };

const S = pptx.ShapeType;
const ROOT = '课程生产/第一阶段/第03课_Workflow类AI产品设计与实操_重做工作区';
const OUT = `${ROOT}/10_正式PPT_第03课_Workflow类AI产品设计与实操.pptx`;
const C = {
  bg: 'F4F0E8', paper: 'FBF9F4', ink: '20211F', text: '292A27', muted: '77766F', line: 'D8D0C2', white: 'FFFFFF',
  red: 'C35331', teal: '387B77', orange: 'D18A38', purple: '74546A', green: '527553', blue: '405D78',
  softRed: 'F2E2DB', softTeal: 'DDEBE8', softOrange: 'F2E7D5', softPurple: 'E8E0E7', softGreen: 'E2E9DE', softBlue: 'DDE5EB'
};
const accents = [C.red, C.teal, C.orange, C.purple, C.green, C.blue];
const softs = [C.softRed, C.softTeal, C.softOrange, C.softPurple, C.softGreen, C.softBlue];

function tx(slide, text, x, y, w, h, o = {}) {
  slide.addText(text, { x, y, w, h, margin: 0, fontFace: 'PingFang SC', fontSize: 16, color: C.text, fit: 'shrink', valign: 'top', breakLine: false, ...o });
}
function sh(slide, type, x, y, w, h, fill, line = fill, o = {}) {
  slide.addShape(type, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 1 }, ...o });
}
function ln(slide, x, y, w, color = C.line, width = 1) { slide.addShape(S.line, { x, y, w, h: 0, line: { color, width } }); }
function pill(slide, text, x, y, w, color = C.red, fill = color, textColor = C.white) {
  sh(slide, S.roundRect, x, y, w, 0.36, fill, fill, { rectRadius: 0.06 });
  tx(slide, text, x + 0.05, y + 0.08, w - 0.1, 0.18, { fontSize: 10, bold: true, color: textColor, align: 'center' });
}
function moduleOf(n) {
  if (n <= 5) return '课程任务'; if (n <= 10) return '01 Workflow是什么'; if (n <= 15) return '02 从SOP到流程';
  if (n <= 20) return '03 五块积木'; if (n <= 23) return '04 纸面练习'; if (n <= 28) return '05 单路V0.2';
  if (n <= 37) return '06 四路V0.3'; if (n <= 41) return '07 测试与排错'; return '08 总结与作业';
}
function base(n, title, o = {}) {
  const s = pptx.addSlide(); s.background = { color: o.bg || C.bg };
  if (!o.noHeader) {
    pill(s, String(n).padStart(2, '0'), 0.62, 0.3, 0.62, o.accent || C.red);
    tx(s, moduleOf(n).replace(/^\d+\s*/, ''), 1.42, 0.38, 5.8, 0.18, { fontSize: 10, bold: true, color: C.muted, charSpacing: 1 });
    tx(s, title, 0.62, 0.78, 12.0, 0.54, { fontSize: 25, bold: true, color: C.ink });
    ln(s, 0.62, 1.42, 12.05);
  }
  if (!o.noFooter) {
    tx(s, 'AI 产品经理系统课程 · 第一阶段', 0.62, 7.06, 4.5, 0.16, { fontSize: 8, color: C.muted });
    tx(s, moduleOf(n), 8.0, 7.06, 3.9, 0.16, { fontSize: 8, color: C.muted, align: 'right' });
    tx(s, String(n).padStart(2, '0'), 12.1, 7.04, 0.55, 0.18, { fontSize: 8.5, color: C.muted, align: 'right' });
  }
  s.addNotes(`第${n}页｜${moduleOf(n)}\n逐页内容：${ROOT}/09_PPT结构与逐页内容\n老师讲解：${ROOT}/05_老师版讲义`);
  return s;
}
function card(slide, x, y, w, h, head, body = '', accent = C.red, o = {}) {
  sh(slide, S.roundRect, x, y, w, h, o.fill || C.paper, o.line || C.line, { rectRadius: 0.05, shadow: { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: 0.05 } });
  sh(slide, S.rect, x, y, 0.07, h, accent, accent);
  tx(slide, head, x + 0.24, y + 0.2, w - 0.46, 0.44, { fontSize: o.headSize || 17, bold: true, color: C.ink, align: o.align || 'left' });
  if (body) tx(slide, body, x + 0.24, y + 0.78, w - 0.48, h - 0.97, { fontSize: o.bodySize || 13.5, color: o.bodyColor || C.muted, valign: o.valign || 'top', paraSpaceAfterPt: 7, align: o.bodyAlign || 'left' });
}
function banner(slide, text, y = 5.85, accent = C.red) {
  const idx = Math.max(0, accents.indexOf(accent));
  sh(slide, S.roundRect, 1.02, y, 11.28, 0.58, softs[idx], softs[idx], { rectRadius: 0.05 });
  tx(slide, text, 1.25, y + 0.14, 10.82, 0.28, { fontSize: 15, bold: true, color: accent, align: 'center' });
}
function cards(slide, items, cols = 3, y = 1.78, h = 1.8) {
  const gap = 0.24, x0 = 0.72, total = 11.9, w = (total - gap * (cols - 1)) / cols;
  items.forEach((it, i) => card(slide, x0 + (i % cols) * (w + gap), y + Math.floor(i / cols) * (h + 0.24), w, h, it[0], it[1] || '', it[2] || accents[i % 6], it[3] || {}));
}
function flow(slide, items, y = 2.0, o = {}) {
  const x0 = o.x || 0.66, total = o.w || 12.0, gap = o.gap || 0.18, h = o.h || 1.72, w = (total - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const x = x0 + i * (w + gap), a = it[2] || accents[i % 6];
    card(slide, x, y, w, h, it[0], it[1] || '', a, { headSize: o.headSize || 15, bodySize: o.bodySize || 11, bodyAlign: 'center', align: 'center', valign: 'mid' });
    if (i < items.length - 1) sh(slide, S.chevron, x + w + 0.03, y + 0.61, gap - 0.06, 0.45, C.line, C.line);
  });
}
function two(slide, a, b, y = 1.75, h = 4.72) {
  card(slide, 0.78, y, 5.78, h, a[0], a[1], a[2] || C.teal, a[3] || { headSize: 20, bodySize: 15 });
  card(slide, 6.78, y, 5.78, h, b[0], b[1], b[2] || C.orange, b[3] || { headSize: 20, bodySize: 15 });
}
function quote(slide, text, sub = '', accent = C.red) {
  sh(slide, S.rect, 1.05, 2.05, 0.08, 1.35, accent, accent);
  tx(slide, text, 1.45, 2.05, 10.55, 1.2, { fontSize: 28, bold: true, color: C.ink, align: 'center', valign: 'mid' });
  if (sub) tx(slide, sub, 1.5, 3.62, 10.45, 0.5, { fontSize: 16, color: C.muted, align: 'center' });
}

// 01 cover
{
  const s = base(1, '', { noHeader: true, noFooter: true });
  sh(s, S.rect, 0, 0, 0.22, 7.5, C.red, C.red);
  tx(s, '03', 0.88, 0.72, 1.1, 0.66, { fontSize: 35, bold: true, color: C.red });
  tx(s, 'Workflow类AI产品设计', 0.88, 1.58, 8.8, 0.72, { fontSize: 38, bold: true, color: C.ink });
  tx(s, '从一次模型调用，到可运行、可分支、可测试的业务流程', 0.92, 2.62, 9.2, 0.5, { fontSize: 19, color: C.muted });
  ln(s, 0.92, 3.35, 3.1, C.red, 3);
  ['输入', '理解', '分支', '安全输出'].forEach((x, i) => pill(s, x, 9.45, 1.15 + i * 1.05, 2.15, accents[i], softs[i], accents[i]));
  pill(s, '第一阶段 · 第03课', 9.66, 6.43, 2.55, C.red);
}
{
  const s = base(2, '今天会完成什么');
  cards(s, [['能解释', '说清Prompt与Workflow的差别'], ['能画', '把人工SOP画成四路流程'], ['能搭', '跑通最小Workflow'], ['能测', '用测试集定位具体节点']], 4, 2.0, 2.65);
  banner(s, '最后交付：一张流程图＋一条可运行Workflow＋五条测试记录。', 5.3);
}
{
  const s = base(3, '三版递进作品');
  flow(s, [['V0.1', '纸面四路图'], ['V0.2', '开始→LLM→结束'], ['V0.3', '分类→分支→安全输出'], ['V1.0', '自己的工作场景']], 2.05, { headSize: 19, bodySize: 13 });
  banner(s, '一次只增加一种复杂度，出错时才知道查哪里。', 5.0, C.teal);
}
{
  const s = base(4, '第02课停在哪里');
  two(s, ['一次Prompt', '输入一次\n↓\n模型处理一次\n↓\n输出一次', C.blue, { headSize: 21, bodySize: 21, bodyAlign: 'center', valign: 'mid' }], ['Workflow', '输入\n↓\n理解与判断\n↓\n不同处理路径\n↓\n输出', C.red, { headSize: 21, bodySize: 20, bodyAlign: 'center', valign: 'mid' }]);
}
{
  const s = base(5, '四个问题，本来就不该走同一条路');
  cards(s, [['规则', '满减能和会员券叠加吗？'], ['数据', '昨天门店GMV是多少？'], ['投诉', '平台乱扣费，我要投诉。'], ['模糊', '活动怎么回事？']], 4, 1.9, 2.85);
  banner(s, '模型都能生成文字，但业务处理必须不同。', 5.32);
}
{
  const s = base(6, 'Workflow是什么');
  quote(s, '把任务拆成步骤，再通过变量和连线，按预设顺序或条件重复执行。', '先把它理解为：可运行的流程图。', C.red);
}
{
  const s = base(7, '流程图与Workflow');
  two(s, ['流程图', '描述应该怎样走\n\n静态表达业务逻辑', C.teal, { headSize: 22, bodySize: 19, bodyAlign: 'center', valign: 'mid' }], ['Workflow', '接收输入｜传递变量\n调用节点｜产生输出\n保留运行日志', C.orange, { headSize: 22, bodySize: 17, bodyAlign: 'center', valign: 'mid' }]);
  banner(s, '先用流程图想清业务，再让Workflow执行。', 6.0);
}
{
  const s = base(8, '三个适用信号');
  cards(s, [['多步骤', '后一步依赖前一步结果'], ['有分支', '不同输入走不同路径'], ['有边界', '需要兜底或人工确认']], 3, 1.95, 3.05);
  banner(s, '满足两个，就值得考虑Workflow。', 5.45, C.teal);
}
{
  const s = base(9, '不是所有任务都需要Workflow');
  cards(s, [['标题润色', '一次生成即可'], ['单段总结', '输入和输出明确'], ['头脑风暴', '一次性开放探索']], 3, 1.9, 3.0);
  banner(s, '节点越多，不等于产品越智能。', 5.42, C.orange);
}
{
  const s = base(10, 'Workflow与Agent：一句区分');
  two(s, ['Workflow', '产品预设主要路径\n强调可控与可复现', C.teal, { headSize: 22, bodySize: 19, bodyAlign: 'center', valign: 'mid' }], ['Agent', '模型更自主决定下一步\n强调灵活，但更难控制', C.purple, { headSize: 22, bodySize: 18, bodyAlign: 'center', valign: 'mid' }]);
  banner(s, '本课只学Workflow，不展开Agent。', 6.0);
}
{
  const s = base(11, '项目背景｜商家咨询分流助手');
  card(s, 0.78, 1.72, 4.1, 4.7, '邻里增长平台', '虚构的本地商家服务平台\n\n商家运营每天处理：\n活动规则｜经营数据\n扣费投诉｜模糊问题', C.blue, { headSize: 21, bodySize: 17, bodyAlign: 'center', valign: 'mid' });
  card(s, 5.18, 1.72, 7.2, 2.05, '当前问题', '人工先读、再判断、再转交，重复分流成本高。', C.orange, { headSize: 19, bodySize: 18, bodyAlign: 'center' });
  card(s, 5.18, 4.05, 7.2, 2.37, '第一步要做什么？', '先做“理解＋分流＋安全回复”，不做全自动客服。', C.red, { headSize: 19, bodySize: 19, bodyAlign: 'center', valign: 'mid' });
  pill(s, '教学虚构｜不对应真实公司项目', 8.62, 0.82, 3.35, C.purple, C.softPurple, C.purple);
}
{
  const s = base(12, '第一版目标与红线');
  two(s, ['目标', '识别问题类型\n进入正确路径\n给出安全回复', C.green, { headSize: 22, bodySize: 19, bodyAlign: 'center', valign: 'mid' }], ['红线', '不编实时数据\n不承诺退款赔偿\n信息不足不强猜', C.red, { headSize: 22, bodySize: 19, bodyAlign: 'center', valign: 'mid' }]);
}
{
  const s = base(13, '先定义输出｜结果必须可使用、可检查');
  cards(s, [['issue_type', '问题类型'], ['handling_method', '处理方式'], ['reference_reply', '参考回复'], ['need_human', '是否人工'], ['reason', '内部原因']], 5, 2.0, 2.65);
  banner(s, '输出不是一段“好看的话”，而是业务可使用的结果。', 5.3);
}
{
  const s = base(14, '再定义输入｜第一版刻意控制复杂度');
  quote(s, 'user_query', '商家提交的原始问题｜文本｜必填', C.blue);
  banner(s, '商家ID、门店ID和历史对话可以后加，今天先不加。', 5.18, C.blue);
}
{
  const s = base(15, '把人工SOP翻译成流程');
  flow(s, [['读问题', '接收原始输入'], ['判断类型', '理解用户意图'], ['选择路径', '执行分支条件'], ['安全处理', '按边界行动'], ['输出结果', '用户话术＋内部字段']], 1.95, { headSize: 15, bodySize: 10.5 });
  banner(s, '逻辑没画清楚，打开平台只会得到一张复杂画布。', 4.95, C.orange);
}
{
  const s = base(16, 'Workflow五块积木');
  cards(s, [['节点', '一项动作'], ['连线', '顺序与流向'], ['变量', '传递的信息'], ['分支', '选择路径'], ['输出', '返回结果']], 5, 2.0, 2.7);
  banner(s, '平台会变，五块积木基本不变。', 5.35, C.teal);
}
{
  const s = base(17, '节点与连线');
  flow(s, [['开始', '接收问题'], ['判断咨询类型', 'LLM节点'], ['选择处理路径', '条件节点'], ['返回结果', '结束节点']], 2.05, { headSize: 16, bodySize: 11 });
  banner(s, '用业务动作命名节点，不要只叫“LLM1”“处理2”。', 5.1);
}
{
  const s = base(18, '变量是节点之间的接力棒');
  flow(s, [['user_query', '原始问题'], ['issue_type', '分类标签'], ['final_reply', '最终话术']], 2.1, { x: 1.0, w: 11.3, h: 2.05, headSize: 19, bodySize: 14 });
  banner(s, '画了线，不等于变量自动传对。', 5.2, C.orange);
}
{
  const s = base(19, '分类与分支不是一回事');
  two(s, ['分类LLM', '理解语言\n回答：它是什么？\n\n输出 issue_type', C.purple, { headSize: 22, bodySize: 18, bodyAlign: 'center', valign: 'mid' }], ['条件节点', '执行硬规则\n回答：如果是它，走哪里？\n\n精确匹配标签', C.teal, { headSize: 22, bodySize: 18, bodyAlign: 'center', valign: 'mid' }]);
}
{
  const s = base(20, '确定性与概率性');
  two(s, ['更偏概率', 'LLM分类｜LLM生成\n结果可能波动', C.orange, { headSize: 22, bodySize: 19, bodyAlign: 'center', valign: 'mid' }], ['更偏确定', '条件匹配｜固定话术\n相同规则稳定执行', C.green, { headSize: 22, bodySize: 19, bodyAlign: 'center', valign: 'mid' }]);
  banner(s, '需要理解时用模型，需要守规则时用结构。', 6.0);
}
{
  const s = base(21, 'V0.1纸面任务｜先不要打开工具');
  cards(s, [['4分钟', '个人画出最小流程'], ['2分钟', '两人互查路径'], ['必须包含', '1输入｜1分类｜4路径｜1输出'], ['必须写出', '2条系统红线']], 4, 1.9, 2.85);
  banner(s, '同桌必须能只看图讲清每一种问题会去哪。', 5.3, C.orange);
}
{
  const s = base(22, 'V0.1参考流程');
  // central flow + branch cards
  card(s, 0.62, 2.65, 1.65, 1.2, '开始', 'user_query', C.blue, { headSize: 16, bodySize: 12, align: 'center', bodyAlign: 'center' });
  sh(s, S.chevron, 2.36, 3.0, 0.45, 0.45, C.line, C.line);
  card(s, 2.9, 2.65, 1.85, 1.2, 'AI分类', 'issue_type', C.purple, { headSize: 16, bodySize: 12, align: 'center', bodyAlign: 'center' });
  sh(s, S.chevron, 4.83, 3.0, 0.45, 0.45, C.line, C.line);
  card(s, 5.36, 2.65, 1.75, 1.2, '条件', '四路分流', C.teal, { headSize: 16, bodySize: 12, align: 'center', bodyAlign: 'center' });
  [['RULE', '规则卡回答'], ['DATA', '需查系统'], ['COMPLAINT', '转人工'], ['OTHER', '追问澄清']].forEach((d, i) => card(s, 7.48 + (i % 2) * 2.47, 1.72 + Math.floor(i / 2) * 2.15, 2.2, 1.7, d[0], d[1], accents[i], { headSize: 15, bodySize: 12, align: 'center', bodyAlign: 'center', valign: 'mid' }));
  banner(s, 'DATA不编数｜COMPLAINT不承诺｜OTHER不强猜', 5.75);
}
{
  const s = base(23, '', { noHeader: true, noFooter: true, bg: C.softBlue });
  tx(s, '休息 5 分钟', 1.0, 1.55, 11.3, 0.75, { fontSize: 38, bold: true, color: C.blue, align: 'center' });
  tx(s, '保存V0.1｜老师检查演示环境', 1.0, 2.75, 11.3, 0.55, { fontSize: 23, bold: true, color: C.ink, align: 'center' });
  tx(s, '05:00', 1.0, 5.3, 11.3, 0.65, { fontSize: 31, bold: true, color: C.muted, align: 'center' });
}
{
  const s = base(24, '平台演示路线｜一次只增加一种复杂度');
  flow(s, [['跑单路', '开始→LLM→结束'], ['看日志', '逐节点输入输出'], ['加分类', '只输出标签'], ['加分支', '四种处理'], ['做测试', '制造失败并修正']], 1.95, { headSize: 14, bodySize: 10 });
  banner(s, '界面名称会变，按节点功能找同类项。', 4.95, C.blue);
}
{
  const s = base(25, '开始节点｜创建唯一必填输入');
  cards(s, [['变量名', 'user_query'], ['类型', '文本 String'], ['要求', '必填'], ['示例', '满减能和会员券一起用吗？']], 4, 2.0, 2.65);
  banner(s, '开始节点决定流程真正拿到了什么。', 5.3, C.blue);
}
{
  const s = base(26, 'LLM节点｜Prompt必须插入真实变量');
  card(s, 0.9, 1.72, 11.55, 4.55, 'V0.2 Prompt骨架', '你是商家咨询处理助手。\n请阅读用户问题，说明正式处理前需要确认什么信息。\n不得编造实时业务数据，不得承诺退款或赔偿。\n\n用户问题：〔从变量选择器插入 user_query〕', C.purple, { headSize: 20, bodySize: 18 });
  banner(s, '不要只手敲占位符后假设平台一定识别。', 6.05, C.orange);
}
{
  const s = base(27, '结束节点与首次运行');
  flow(s, [['开始输入', 'user_query有值？'], ['LLM输入', '看到真实问题？'], ['LLM输出', '有处理建议？'], ['结束输出', 'final_reply引用正确？']], 2.0, { headSize: 15, bodySize: 11 });
  banner(s, '逐节点读日志：输入是什么，输出是什么。', 5.0, C.teal);
}
{
  const s = base(28, '跑通不等于可用');
  quote(s, 'V0.2会用同一种方式处理所有问题。', '它只证明链路通了，还没有证明业务逻辑正确。', C.orange);
}
{
  const s = base(29, '升级目标｜四路V0.3');
  flow(s, [['开始', 'user_query'], ['分类LLM', 'issue_type'], ['条件分支', '精确匹配'], ['四种处理', '守住边界'], ['结束', '结构化结果']], 1.95, { headSize: 15, bodySize: 10 });
  banner(s, '先测试分类，再接条件，最后逐条接处理。', 4.95, C.teal);
}
{
  const s = base(30, '四个标签必须有进入标准');
  cards(s, [['RULE', '平台/活动/经营规则'], ['DATA', '实时订单/GMV/结算/状态'], ['COMPLAINT', '投诉/扣费/退款赔付/高风险'], ['OTHER', '信息不足或未覆盖']], 4, 1.9, 2.9);
  banner(s, '冲突时高风险优先；无法确定进入OTHER。', 5.35);
}
{
  const s = base(31, '分类Prompt｜四个组成部分');
  cards(s, [['标签定义', '什么情况进入哪类'], ['冲突优先级', '投诉与扣费优先COMPLAINT'], ['严格输出', '只输出一个大写标签'], ['边界示例', '典型＋模糊问题']], 4, 1.85, 2.9);
  banner(s, '完整可复制Prompt见课堂材料04。', 5.3, C.purple);
}
{
  const s = base(32, '条件分支｜自然语言输出要接确定性接口');
  flow(s, [['分类输出', 'issue_type'], ['精确条件', '== RULE / DATA / ...'], ['业务路径', '进入对应处理']], 2.05, { x: 1.0, w: 11.3, h: 2.05, headSize: 18, bodySize: 13 });
  banner(s, '优先枚举/结构化输出；必须保留OTHER兜底。', 5.15, C.teal);
}
{
  const s = base(33, 'RULE路径｜只能依据课堂规则卡');
  two(s, ['课堂规则', '平台满减与平台新客券不可叠加\n商家券看活动配置页', C.purple], ['输出要求', '结论 → 依据 → 不确定性\n未覆盖：核实当前有效规则', C.teal]);
  banner(s, '这不是RAG：规则材料是固定粘贴进去的。', 6.0, C.orange);
}
{
  const s = base(34, 'DATA路径｜没有数据源，就不要生成数字');
  quote(s, '该问题需要查询业务系统中的实时数据。', '本课堂版本未连接订单或经营系统，因此不会生成具体数值。', C.blue);
  banner(s, '模型可以解释数据，但真实数据必须由业务系统提供。', 5.2, C.blue);
}
{
  const s = base(35, 'COMPLAINT路径｜转人工不是失败');
  quote(s, '记录原始问题，并转交人工处理。', '当前不自动承诺退款、赔偿或处理结果。', C.red);
  banner(s, '质量可控优先于表面上的“全自动”。', 5.2);
}
{
  const s = base(36, 'OTHER路径｜只追问一个最必要问题');
  card(s, 1.05, 1.88, 11.2, 3.85, '示例', '“活动怎么回事？”\n\n为了准确处理，请补充具体活动名称；你想了解的是参与条件、叠加规则，还是生效时间？', C.orange, { headSize: 22, bodySize: 21, bodyAlign: 'center', valign: 'mid' });
  banner(s, 'OTHER是安全兜底，不是垃圾桶。', 6.0, C.orange);
}
{
  const s = base(37, '完整Workflow｜四路分流与后续接口');
  card(s, 0.42, 2.75, 1.25, 1.05, '开始', 'query', C.blue, { headSize: 14, bodySize: 10, align: 'center', bodyAlign: 'center' });
  card(s, 1.95, 2.75, 1.35, 1.05, '分类', 'type', C.purple, { headSize: 14, bodySize: 10, align: 'center', bodyAlign: 'center' });
  card(s, 3.58, 2.75, 1.35, 1.05, '条件', '四路', C.teal, { headSize: 14, bodySize: 10, align: 'center', bodyAlign: 'center' });
  ln(s, 1.67, 3.28, 0.28, C.line, 2); ln(s, 3.30, 3.28, 0.28, C.line, 2); ln(s, 4.93, 3.28, 0.27, C.line, 2);
  s.addShape(S.line, { x: 5.2, y: 1.91, w: 0, h: 2.73, line: { color: C.line, width: 2 } });
  s.addShape(S.line, { x: 9.38, y: 1.91, w: 0, h: 2.73, line: { color: C.line, width: 2 } });
  const branch = (y, label, body, color) => {
    ln(s, 5.2, y + 0.35, 0.25, C.line, 2);
    sh(s, S.roundRect, 5.45, y, 3.68, 0.7, C.paper, C.line, { rectRadius: 0.04, shadow: { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: 0.05 } });
    sh(s, S.rect, 5.45, y, 0.07, 0.7, color, color);
    tx(s, label, 5.68, y + 0.17, 1.35, 0.26, { fontSize: 14, bold: true, color: C.ink });
    tx(s, body, 7.05, y + 0.18, 1.78, 0.25, { fontSize: 12, color: C.muted, align: 'right' });
    ln(s, 9.13, y + 0.35, 0.25, C.line, 2);
  };
  branch(1.56, 'RULE', '固定规则卡', C.red);
  branch(2.47, 'DATA', '需查系统', C.teal);
  branch(3.38, 'COMPLAINT', '转人工', C.orange);
  branch(4.29, 'OTHER', '追问', C.purple);
  ln(s, 9.38, 3.28, 0.42, C.line, 2);
  card(s, 9.8, 2.75, 1.45, 1.05, '结束', '结果', C.green, { headSize: 14, bodySize: 10, align: 'center', bodyAlign: 'center' });
  banner(s, '下一课：用RAG替换RULE分支的固定材料入口。', 5.85, C.purple);
}
{
  const s = base(38, '五条测试｜先写预期，再运行');
  cards(s, [['T1', '满减能和会员券叠加吗？'], ['T2', '帮我查昨天门店GMV。'], ['T3', '平台乱扣费，我要投诉。'], ['T4', '帮我写七夕活动标题。'], ['T5', '活动怎么回事？']], 5, 1.85, 2.85);
  banner(s, '不要看到结果后再反推“预期”。', 5.3, C.orange);
}
{
  const s = base(39, '预期路径与安全处理');
  cards(s, [['RULE', '按规则卡回答'], ['DATA', '不输出数值'], ['COMPLAINT', '转人工不承诺'], ['OTHER', '未覆盖内容生成'], ['OTHER', '追问活动信息']], 5, 1.9, 2.75);
  banner(s, '典型、模糊和高风险输入都要覆盖。', 5.25, C.teal);
}
{
  const s = base(40, '故意制造一个Bad Case');
  two(s, ['分类节点实际输出', '“RULE，因为这是活动规则问题”', C.orange, { headSize: 20, bodySize: 19, bodyAlign: 'center', valign: 'mid' }], ['条件节点接受值', '精确等于：RULE\n\n结果：匹配失败或走兜底', C.red, { headSize: 20, bodySize: 18, bodyAlign: 'center', valign: 'mid' }]);
  banner(s, '先换模型，还是先检查输出接口和条件？', 6.0, C.orange);
}
{
  const s = base(41, '四步排查法');
  flow(s, [['01 看输入', '变量有没有值'], ['02 看输出', '实际标签是什么'], ['03 看条件', '引用和值对不对'], ['04 再优化', 'Prompt或模型']], 2.0, { headSize: 17, bodySize: 12 });
  banner(s, '不要把所有失败都归咎于模型。', 5.05);
}
{
  const s = base(42, '四句总结＋课后V1.0');
  card(s, 0.72, 1.62, 7.15, 4.95, '四句话', '1  Prompt管一次处理，Workflow管多步与分支\n\n2  先画业务路径，再拖平台节点\n\n3  模型负责理解，条件与固定输出守边界\n\n4  能跑只是起点，测试决定能否成为产品', C.red, { headSize: 20, bodySize: 16 });
  card(s, 8.17, 1.62, 4.45, 4.95, '课后V1.0', '自选一个小场景\n\n至少三条路径\n一个兜底\n五条测试\n一次失败修正\n\n下一课：RAG升级RULE分支', C.blue, { headSize: 20, bodySize: 15.5, bodyAlign: 'center' });
}

if (pptx._slides.length !== 42) throw new Error(`Expected 42 slides, got ${pptx._slides.length}`);
pptx.writeFile({ fileName: OUT });
