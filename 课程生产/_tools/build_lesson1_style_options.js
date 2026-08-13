const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第01课PPT版式候选';
pptx.title = '第01课PPT版式候选对比';
pptx.company = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };

const S = pptx.ShapeType;
const OUT = '课程生产/第一阶段/第01课_AI产品经理岗位认知与转型路径_重做工作区/19_PPT版式候选_3套对比.pptx';

const A = {
  name: 'A｜体系课程蓝', bg: 'F5F8FC', paper: 'FFFFFF', navy: '122A49', ink: '203047', muted: '64748B',
  main: '2563EB', cyan: '22B8CF', purple: '7C3AED', orange: 'F59E0B', green: '16A34A', red: 'E34D59', line: 'DCE5F0'
};
const B = {
  name: 'B｜战略编辑风', bg: 'F4F0E8', paper: 'FBF9F4', navy: '20211F', ink: '292A27', muted: '77766F',
  main: 'C35331', cyan: '387B77', purple: '74546A', orange: 'D18A38', green: '527553', red: 'B6483F', line: 'D8D0C2'
};
const C = {
  name: 'C｜AI未来感', bg: '081426', paper: '10213A', navy: 'F7FAFF', ink: 'E7EEF9', muted: '9DB0C9',
  main: '4E7CFF', cyan: '27D3EE', purple: '9A6BFF', orange: 'FF9C5B', green: '38D996', red: 'FF6685', line: '28405E'
};

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h, margin: 0, fontFace: 'PingFang SC', fontSize: 14,
    color: opts.color || '203047', bold: opts.bold || false,
    align: opts.align || 'left', valign: opts.valign || 'top',
    breakLine: false, fit: 'shrink', ...opts
  });
}

function line(slide, x, y, w, color, width = 1) {
  slide.addShape(S.line, { x, y, w, h: 0, line: { color, width } });
}

function pill(slide, text, x, y, w, color, bg, dark = false) {
  slide.addShape(S.roundRect, { x, y, w, h: 0.34, rectRadius: 0.12, fill: { color: bg }, line: { color: bg } });
  addText(slide, text, x + 0.08, y + 0.075, w - 0.16, 0.16, { fontSize: 9, bold: true, color: dark ? 'FFFFFF' : color, align: 'center' });
}

function optionMark(slide, theme, option, pageType) {
  pill(slide, option, 0.62, 0.31, 1.48, theme.main, theme.main, true);
  addText(slide, pageType, 2.25, 0.39, 2.2, 0.18, { fontSize: 9.5, bold: true, color: theme.muted, charSpacing: 1.2 });
}

function footer(slide, theme, n) {
  addText(slide, 'AI 产品经理系统课程 · 第01课 · 版式候选', 0.62, 7.06, 4.5, 0.16, { fontSize: 8, color: theme.muted });
  addText(slide, String(n).padStart(2, '0'), 12.1, 7.04, 0.55, 0.18, { fontSize: 8.5, color: theme.muted, align: 'right' });
}

function title(slide, theme, text, kicker, n) {
  optionMark(slide, theme, theme.name, kicker);
  addText(slide, text, 0.62, 0.78, 11.8, 0.52, { fontSize: 25, bold: true, color: theme.navy });
  line(slide, 0.62, 1.42, 12.05, theme.line, 1);
  footer(slide, theme, n);
}

function card(slide, theme, x, y, w, h, head, body, accent, opts = {}) {
  slide.addShape(S.roundRect, {
    x, y, w, h, rectRadius: opts.radius || 0.06,
    fill: { color: opts.fill || theme.paper }, line: { color: opts.line || theme.line, width: 1 },
    shadow: opts.shadow === false ? undefined : { type: 'outer', color: '000000', blur: 1, angle: 45, distance: 1, opacity: theme === C ? 0.24 : 0.10 }
  });
  if (opts.topLine) slide.addShape(S.rect, { x, y, w, h: 0.08, fill: { color: accent }, line: { color: accent } });
  else slide.addShape(S.rect, { x, y, w: 0.07, h, fill: { color: accent }, line: { color: accent } });
  addText(slide, head, x + 0.28, y + 0.28, w - 0.52, 0.3, { fontSize: opts.headSize || 15, bold: true, color: opts.headColor || theme.navy });
  addText(slide, body, x + 0.28, y + 0.82, w - 0.55, h - 1.06, { fontSize: opts.bodySize || 11.5, color: opts.bodyColor || theme.muted, breakLine: false, valign: 'top', paraSpaceAfterPt: 7 });
}

function numberDot(slide, theme, x, y, n, color) {
  slide.addShape(S.ellipse, { x, y, w: 0.42, h: 0.42, fill: { color }, line: { color } });
  addText(slide, String(n), x, y + 0.11, 0.42, 0.14, { fontSize: 9, bold: true, color: 'FFFFFF', align: 'center' });
}

// 00 selection guide
{
  const s = pptx.addSlide(); s.background = { color: 'EEF3F8' };
  addText(s, '第01课 PPT 版式候选', 0.72, 0.62, 8.6, 0.65, { fontSize: 31, bold: true, color: '122A49' });
  addText(s, '请重点比较：专业感、教学亲和度、信息密度和与第0课的连续性', 0.75, 1.43, 10.8, 0.38, { fontSize: 15, color: '64748B' });
  const options = [
    [A, '最稳妥', '延续第0课的课程体系感\n清晰、克制、容易长时间观看'],
    [B, '最专业', '更像咨询报告与商业课程\n强调判断、案例和职业成熟度'],
    [C, '最吸睛', '更强的AI主题和视觉记忆点\n适合宣传，但课堂长时间观看更挑制作']
  ];
  options.forEach((d, i) => {
    const [t, tag, desc] = d, x = 0.72 + i * 4.18;
    s.addShape(S.roundRect, { x, y: 2.25, w: 3.72, h: 3.5, rectRadius: 0.08, fill: { color: t.bg }, line: { color: t.line, width: 1 } });
    s.addShape(S.rect, { x, y: 2.25, w: 3.72, h: 0.12, fill: { color: t.main }, line: { color: t.main } });
    pill(s, tag, x + 0.28, 2.65, 1.05, t.main, t.main, true);
    addText(s, t.name, x + 0.28, 3.27, 3.15, 0.36, { fontSize: 19, bold: true, color: t.navy });
    addText(s, desc, x + 0.28, 4.05, 3.1, 1.05, { fontSize: 12.5, color: t.muted, breakLine: false, paraSpaceAfterPt: 8 });
  });
  addText(s, '每套均展示：封面｜能力框架｜Bad Case｜课堂练习', 0.75, 6.38, 11.7, 0.34, { fontSize: 13, bold: true, color: '203047', align: 'center' });
  footer(s, A, 1);
}

function optionA() {
  // cover
  {
    const s = pptx.addSlide(); s.background = { color: A.navy };
    s.addShape(S.ellipse, { x: 9.0, y: -1.6, w: 5.8, h: 5.8, fill: { color: A.main, transparency: 25 }, line: { color: A.main, transparency: 100 } });
    s.addShape(S.ellipse, { x: 10.6, y: 4.1, w: 3.4, h: 3.4, fill: { color: A.cyan, transparency: 52 }, line: { color: A.cyan, transparency: 100 } });
    pill(s, '方案 A｜体系课程蓝', 0.82, 0.78, 2.2, A.cyan, A.cyan, true);
    addText(s, 'AI产品经理岗位认知\n与转型路径', 0.82, 1.62, 8.4, 1.45, { fontSize: 36, bold: true, color: 'FFFFFF', breakLine: false, paraSpaceAfterPt: 10 });
    line(s, 0.84, 3.48, 5.8, A.cyan, 3);
    addText(s, '看懂岗位 · 判断方向 · 找到第一步', 0.84, 3.85, 7.2, 0.4, { fontSize: 17, bold: true, color: 'CFE0F6' });
    addText(s, '第一阶段 · 第01课', 0.84, 6.62, 3.5, 0.25, { fontSize: 10, color: '9FB6D1' });
  }
  // framework
  {
    const s = pptx.addSlide(); s.background = { color: A.bg }; title(s, A, 'AI产品经理的五类核心能力', '能力框架页', 3);
    const data = [
      ['问题与业务判断', '先判断问题是否真实、AI是否必要', A.main],
      ['AI边界与可行性', '理解能力上限、数据条件与风险', A.purple],
      ['产品与人机协作', '设计流程、交互与人工兜底', A.cyan],
      ['效果与Bad Case', '定义标准、评测并定位失败原因', A.orange],
      ['落地与持续优化', '跨团队推进上线、运营和迭代', A.green]
    ];
    data.forEach((d, i) => {
      const x = 0.64 + i * 2.52;
      card(s, A, x, 1.83, 2.25, 3.85, d[0], d[1], d[2], { topLine: true, headSize: 14.2, bodySize: 11.2 });
      numberDot(s, A, x + 0.25, 5.08, i + 1, d[2]);
    });
    addText(s, '它不是行业官方评分模型，而是从真实工作链路归纳出的课堂地图。', 1.15, 6.08, 11, 0.34, { fontSize: 12.5, bold: true, color: A.navy, align: 'center' });
  }
  // case
  {
    const s = pptx.addSlide(); s.background = { color: A.bg }; title(s, A, '一条错误回答，问题一定出在模型吗？', '案例分析页', 4);
    card(s, A, 0.72, 1.72, 4.2, 4.65, '失败案例（Bad Case）', '员工问：“这张优惠券今天还能用吗？”\n\nAI回答：“可以使用。”\n\n实际：引用了上月已失效的活动规则', A.red, { headSize: 17, bodySize: 13 });
    const causes = [
      ['资料', '旧规则没有下线', A.main], ['检索', '没有找到有效版本', A.purple], ['模型', '结论与依据不一致', A.orange],
      ['交互', '没有确认适用范围', A.cyan], ['流程', '高风险承诺未转人工', A.green]
    ];
    causes.forEach((d, i) => {
      const x = 5.25 + (i % 2) * 3.65, y = 1.72 + Math.floor(i / 2) * 1.46;
      card(s, A, x, y, i === 4 ? 7.05 : 3.38, 1.12, d[0], d[1], d[2], { headSize: 13, bodySize: 10.5, shadow: false });
    });
    addText(s, '目标不是统一归因于“模型不好”，而是找到正确的优化对象。', 5.28, 6.22, 7.0, 0.3, { fontSize: 12.5, bold: true, color: A.navy });
  }
  // exercise
  {
    const s = pptx.addSlide(); s.background = { color: A.bg }; title(s, A, '现在填写你的方向判断卡 V0.1', '课堂练习页', 5);
    const fields = [['背景事实', '岗位／行业／用户／流程'], ['经验资产', '最强两项证据'], ['方向假设', '优先方向＋备选方向'], ['最大缺口', '当前只选一个'], ['验证动作', 'JD任务＋项目任务'], ['定位句', '证据→方向→缺口']];
    fields.forEach((d, i) => card(s, A, 0.72 + (i % 3) * 4.17, 1.65 + Math.floor(i / 3) * 2.1, 3.75, 1.68, d[0], d[1], [A.main, A.purple, A.cyan, A.orange, A.green, A.red][i], { shadow: false }));
    pill(s, '3:00', 11.18, 0.82, 1.08, A.red, A.red, true);
    addText(s, '先写“可以被事实证明的内容”，不要求现在就得到最终答案。', 1.2, 6.17, 10.9, 0.34, { fontSize: 13, bold: true, color: A.navy, align: 'center' });
  }
}

function optionB() {
  // cover
  {
    const s = pptx.addSlide(); s.background = { color: B.bg };
    s.addShape(S.rect, { x: 0, y: 0, w: 0.22, h: 7.5, fill: { color: B.main }, line: { color: B.main } });
    addText(s, '01', 0.88, 0.72, 1.1, 0.66, { fontSize: 35, bold: true, color: B.main });
    addText(s, 'AI产品经理岗位认知\n与转型路径', 0.88, 1.68, 8.2, 1.42, { fontSize: 37, bold: true, color: B.navy, breakLine: false });
    addText(s, '看懂岗位／判断方向／建立验证计划', 0.92, 3.48, 7.3, 0.4, { fontSize: 16, color: B.muted });
    line(s, 0.92, 4.15, 3.0, B.main, 3);
    addText(s, 'AI PRODUCT MANAGEMENT', 9.22, 1.05, 3.15, 0.25, { fontSize: 10, bold: true, color: B.muted, charSpacing: 2.2, align: 'right' });
    addText(s, '职业判断不是选热门方向，\n而是用证据形成可验证的假设。', 8.35, 4.72, 3.95, 1.0, { fontSize: 17, bold: true, color: B.main, breakLine: false, align: 'right' });
    pill(s, '方案 B｜战略编辑风', 9.55, 6.44, 2.7, B.main, B.main, true);
  }
  // framework
  {
    const s = pptx.addSlide(); s.background = { color: B.bg }; title(s, B, '五类能力，不是五个孤立知识点', '能力框架页', 7);
    const data = [
      ['01', '判断问题', '业务价值与AI必要性'], ['02', '判断边界', '能力、数据与风险'], ['03', '设计产品', '流程、交互与人机协作'],
      ['04', '验证效果', '指标、评测与失败案例'], ['05', '推动落地', '协作、上线与持续优化']
    ];
    data.forEach((d, i) => {
      const y = 1.7 + i * 0.92;
      addText(s, d[0], 0.83, y + 0.05, 0.65, 0.3, { fontSize: 15, bold: true, color: B.main });
      addText(s, d[1], 1.67, y, 2.15, 0.35, { fontSize: 18, bold: true, color: B.navy });
      addText(s, d[2], 4.15, y + 0.04, 6.7, 0.3, { fontSize: 13, color: B.muted });
      line(s, 0.82, y + 0.62, 11.5, B.line, 1);
    });
    addText(s, '工作链路', 10.7, 0.84, 1.5, 0.26, { fontSize: 11, bold: true, color: B.main, align: 'right' });
  }
  // case
  {
    const s = pptx.addSlide(); s.background = { color: B.bg }; title(s, B, '失败案例不是“模型打分题”', '案例分析页', 8);
    addText(s, '“这张优惠券今天还能用吗？”', 0.82, 1.72, 5.1, 0.5, { fontSize: 23, bold: true, color: B.navy });
    addText(s, 'AI：可以使用。', 0.82, 2.52, 3.5, 0.38, { fontSize: 17, color: B.main, bold: true });
    addText(s, '事实：引用了上月已失效的活动规则。', 0.82, 3.17, 5.1, 0.38, { fontSize: 15, color: B.muted });
    line(s, 6.23, 1.65, 0, B.line, 1);
    const causes = [['资料', '旧规则未下线'], ['检索', '未命中有效版本'], ['模型', '结论与依据不一致'], ['交互', '未确认适用范围'], ['流程', '高风险问题未转人工']];
    causes.forEach((d, i) => {
      const y = 1.65 + i * 0.83;
      addText(s, String(i + 1).padStart(2, '0'), 6.55, y + 0.04, 0.55, 0.24, { fontSize: 11, bold: true, color: B.main });
      addText(s, d[0], 7.28, y, 1.1, 0.32, { fontSize: 15, bold: true, color: B.navy });
      addText(s, d[1], 8.65, y + 0.03, 3.4, 0.3, { fontSize: 12.5, color: B.muted });
      line(s, 6.55, y + 0.52, 5.55, B.line, 1);
    });
    addText(s, '正确归因，才能找到正确的优化对象。', 0.82, 5.55, 5.2, 0.4, { fontSize: 17, bold: true, color: B.main });
  }
  // exercise
  {
    const s = pptx.addSlide(); s.background = { color: B.bg }; title(s, B, '方向判断卡：先证据，后结论', '课堂练习页', 9);
    const left = [['背景事实', '我长期面对谁、负责什么'], ['关键证据', '两项可复述的经历'], ['方向假设', '优先＋备选']];
    const right = [['最大缺口', '只选择当前最关键的一项'], ['验证动作', 'JD任务＋项目任务'], ['个人定位句', '证据→方向→缺口']];
    [left, right].forEach((col, c) => col.forEach((d, i) => {
      const x = 0.82 + c * 6.05, y = 1.72 + i * 1.35;
      addText(s, d[0], x, y, 1.4, 0.3, { fontSize: 15, bold: true, color: B.main });
      addText(s, d[1], x + 1.6, y, 3.85, 0.33, { fontSize: 13, color: B.navy });
      line(s, x, y + 0.6, 5.4, B.line, 1);
    }));
    pill(s, '3分钟独立填写', 10.28, 0.82, 1.95, B.main, B.main, true);
    addText(s, '今天输出的是可修正的方向假设，不是职业测评结果。', 0.85, 6.12, 11.2, 0.35, { fontSize: 14, bold: true, color: B.navy });
  }
}

function optionC() {
  // cover
  {
    const s = pptx.addSlide(); s.background = { color: C.bg };
    s.addShape(S.ellipse, { x: 8.9, y: -1.5, w: 5.7, h: 5.7, fill: { color: C.main, transparency: 38 }, line: { color: C.main, transparency: 100 } });
    s.addShape(S.ellipse, { x: 10.0, y: 3.4, w: 4.5, h: 4.5, fill: { color: C.purple, transparency: 54 }, line: { color: C.purple, transparency: 100 } });
    s.addShape(S.ellipse, { x: 9.55, y: 1.4, w: 2.1, h: 2.1, fill: { color: C.cyan, transparency: 54 }, line: { color: C.cyan, transparency: 100 } });
    pill(s, '方案 C｜AI未来感', 0.82, 0.76, 2.15, C.cyan, C.cyan, true);
    addText(s, '从“会用AI”\n到“会做AI产品”', 0.82, 1.62, 8.7, 1.5, { fontSize: 39, bold: true, color: C.navy, breakLine: false });
    addText(s, 'AI产品经理岗位认知与转型路径', 0.86, 3.52, 7.7, 0.42, { fontSize: 17, bold: true, color: C.cyan });
    addText(s, 'ROLE  ×  DIRECTION  ×  EVIDENCE', 0.86, 4.28, 6.8, 0.27, { fontSize: 11, color: C.muted, charSpacing: 2 });
    addText(s, '第01课', 0.86, 6.62, 2.0, 0.25, { fontSize: 10, color: C.muted });
  }
  // framework
  {
    const s = pptx.addSlide(); s.background = { color: C.bg }; title(s, C, '五类能力构成一条产品闭环', '能力框架页', 11);
    const data = [
      ['问题判断', 'WHY', C.main], ['边界判断', 'CAN', C.purple], ['产品设计', 'HOW', C.cyan], ['效果验证', 'WORKS', C.orange], ['落地优化', 'SCALE', C.green]
    ];
    data.forEach((d, i) => {
      const x = 0.55 + i * 2.54;
      s.addShape(S.roundRect, { x, y: 1.85, w: 2.2, h: 3.65, rectRadius: 0.08, fill: { color: C.paper }, line: { color: d[2], transparency: 25, width: 1.4 } });
      addText(s, d[1], x + 0.2, 2.22, 1.8, 0.27, { fontSize: 10, bold: true, color: d[2], charSpacing: 1.5, align: 'center' });
      addText(s, d[0], x + 0.2, 2.85, 1.8, 0.38, { fontSize: 17, bold: true, color: C.navy, align: 'center' });
      s.addShape(S.ellipse, { x: x + 0.78, y: 3.67, w: 0.64, h: 0.64, fill: { color: d[2], transparency: 8 }, line: { color: d[2] } });
      addText(s, String(i + 1), x + 0.78, 3.87, 0.64, 0.2, { fontSize: 12, bold: true, color: 'FFFFFF', align: 'center' });
      if (i < 4) s.addShape(S.chevron, { x: x + 2.23, y: 3.63, w: 0.28, h: 0.68, fill: { color: C.line }, line: { color: C.line } });
    });
    addText(s, '判断问题 → 判断边界 → 设计产品 → 验证效果 → 持续优化', 1.15, 6.03, 11, 0.36, { fontSize: 14, bold: true, color: C.cyan, align: 'center' });
  }
  // case
  {
    const s = pptx.addSlide(); s.background = { color: C.bg }; title(s, C, 'BAD CASE / 错误答案从哪里产生？', '案例分析页', 12);
    s.addShape(S.roundRect, { x: 0.72, y: 1.72, w: 4.35, h: 4.55, rectRadius: 0.08, fill: { color: C.paper }, line: { color: C.red, transparency: 15, width: 1.5 } });
    pill(s, 'FAILURE', 1.02, 2.02, 1.1, C.red, C.red, true);
    addText(s, '“优惠券今天还能用吗？”', 1.02, 2.78, 3.75, 0.48, { fontSize: 21, bold: true, color: C.navy });
    addText(s, 'AI：可以使用。', 1.02, 3.62, 3.3, 0.38, { fontSize: 16, bold: true, color: C.cyan });
    addText(s, '实际：引用了上月失效规则', 1.02, 4.42, 3.5, 0.34, { fontSize: 13, color: C.muted });
    const causes = [['DATA', '资料过期', C.main], ['RETRIEVE', '检索失准', C.purple], ['MODEL', '依据不一致', C.orange], ['UX', '范围未确认', C.cyan], ['PROCESS', '未转人工', C.green]];
    causes.forEach((d, i) => {
      const x = 5.45 + (i % 2) * 3.45, y = 1.72 + Math.floor(i / 2) * 1.48;
      card(s, C, x, y, i === 4 ? 6.62 : 3.18, 1.16, d[0], d[1], d[2], { headSize: 11, bodySize: 12, shadow: false, line: C.line });
    });
    addText(s, 'TRACE THE FAILURE → FIX THE RIGHT LAYER', 5.48, 6.22, 6.5, 0.25, { fontSize: 10, bold: true, color: C.cyan, charSpacing: 1.2 });
  }
  // exercise
  {
    const s = pptx.addSlide(); s.background = { color: C.bg }; title(s, C, 'BUILD YOUR DIRECTION HYPOTHESIS', '课堂练习页', 13);
    const fields = [['01', 'BACKGROUND', '背景事实'], ['02', 'ASSETS', '关键证据'], ['03', 'DIRECTION', '方向假设'], ['04', 'GAP', '最大缺口'], ['05', 'VALIDATION', '验证动作'], ['06', 'POSITIONING', '定位句']];
    fields.forEach((d, i) => {
      const x = 0.72 + (i % 3) * 4.16, y = 1.62 + Math.floor(i / 3) * 2.15, color = [C.main, C.purple, C.cyan, C.orange, C.green, C.red][i];
      s.addShape(S.roundRect, { x, y, w: 3.72, h: 1.72, rectRadius: 0.08, fill: { color: C.paper }, line: { color, transparency: 20, width: 1.3 } });
      addText(s, d[0], x + 0.24, y + 0.22, 0.5, 0.25, { fontSize: 11, bold: true, color });
      addText(s, d[1], x + 0.83, y + 0.22, 2.35, 0.25, { fontSize: 9.5, bold: true, color, charSpacing: 1.1 });
      addText(s, d[2], x + 0.24, y + 0.88, 3.1, 0.34, { fontSize: 17, bold: true, color: C.navy });
    });
    pill(s, '03:00', 11.15, 0.82, 1.08, C.red, C.red, true);
    addText(s, 'HYPOTHESIS, NOT A FINAL ANSWER', 3.9, 6.17, 5.5, 0.25, { fontSize: 10.5, bold: true, color: C.cyan, charSpacing: 1.4, align: 'center' });
  }
}

optionA();
optionB();
optionC();

pptx.writeFile({ fileName: OUT });
