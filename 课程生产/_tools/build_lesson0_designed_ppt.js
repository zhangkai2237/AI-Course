const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第 00 课：课程说明会、学习地图与第一次 AI 场景思考';
pptx.title = '课程说明会与学习地图';
pptx.company = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN'
};
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: 'F6F8FC' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.09, fill: { color: '36C5F0' }, line: { color: '36C5F0' } } },
    { text: { text: 'AI 产品经理系统课程 · 第一阶段', options: { x: 0.55, y: 7.08, w: 4.5, h: 0.18, fontFace: 'PingFang SC', fontSize: 8.5, color: '718096', margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.02, w: 0.45, h: 0.2, color: '718096', fontFace: 'Aptos', fontSize: 9, align: 'right', margin: 0 }
});

const C = { navy:'142B4A', blue:'246BFD', cyan:'36C5F0', orange:'FF9F43', green:'27AE60', red:'E85D75', ink:'1F2937', gray:'667085', light:'EAF0F8', white:'FFFFFF', bg:'F6F8FC', line:'D9E2EF', purple:'7655D4' };
const S = pptx.ShapeType;

function addTitle(slide, title, kicker='') {
  if (kicker) slide.addText(kicker, {x:0.62,y:0.34,w:4.2,h:0.22,fontSize:9.5,bold:true,color:C.blue,charSpacing:1.2,margin:0});
  slide.addText(title, {x:0.62,y:0.62,w:11.8,h:0.52,fontSize:24,bold:true,color:C.navy,margin:0,breakLine:false,fit:'shrink'});
  slide.addShape(S.line,{x:0.62,y:1.28,w:12.05,h:0,line:{color:C.line,width:1}});
}
function addPill(slide, text, x, y, color=C.blue, w=1.3) {
  slide.addShape(S.roundRect,{x,y,w,h:0.34,rectRadius:0.08,fill:{color,transparency:88},line:{color,transparency:100}});
  slide.addText(text,{x:x+0.05,y:y+0.075,w:w-0.1,h:0.14,fontSize:9,bold:true,color,align:'center',margin:0,fit:'shrink'});
}
function addCard(slide, {x,y,w,h,title,body,color=C.blue,num}) {
  slide.addShape(S.roundRect,{x,y,w,h,rectRadius:0.06,fill:{color:C.white},line:{color:C.line,width:1},shadow:{type:'outer',color:'AAB6C5',blur:1,angle:45,distance:1,opacity:0.16}});
  slide.addShape(S.rect,{x,y,w:0.08,h,fill:{color},line:{color}});
  if (num) {
    slide.addShape(S.ellipse,{x:x+0.26,y:y+0.25,w:0.46,h:0.46,fill:{color},line:{color}});
    slide.addText(String(num),{x:x+0.26,y:y+0.345,w:0.46,h:0.15,fontSize:11,bold:true,color:C.white,align:'center',margin:0});
  }
  const tx=x+(num?0.86:0.28);
  slide.addText(title,{x:tx,y:y+0.24,w:w-(tx-x)-0.25,h:0.32,fontSize:15.5,bold:true,color:C.navy,margin:0,fit:'shrink'});
  slide.addText(body,{x:x+0.28,y:y+0.86,w:w-0.55,h:h-1.08,fontSize:11.5,color:C.gray,breakLine:false,margin:0.02,fit:'shrink',valign:'top',bullet:undefined,paraSpaceAfterPt:8});
}
function addBulletList(slide, items, x=0.9,y=1.65,w=11.4,h=4.8,color=C.blue) {
  items.forEach((it,i)=>{
    const yy=y+i*(h/items.length);
    slide.addShape(S.ellipse,{x, y:yy+0.06,w:0.28,h:0.28,fill:{color},line:{color}});
    slide.addText(String(i+1),{x,y:yy+0.115,w:0.28,h:0.12,fontSize:8.5,bold:true,color:C.white,align:'center',margin:0});
    slide.addText(it.title,{x:x+0.48,y:yy,w:3.3,h:0.32,fontSize:15,bold:true,color:C.navy,margin:0,fit:'shrink'});
    slide.addText(it.body,{x:x+3.75,y:yy,w:w-3.75,h:0.56,fontSize:11.5,color:C.gray,margin:0,fit:'shrink'});
  });
}
function sectionSlide(no,title,subtitle,color=C.blue) {
  const s=pptx.addSlide(); s.background={color:C.navy};
  s.addShape(S.ellipse,{x:9.8,y:-1.1,w:4.7,h:4.7,fill:{color,transparency:42},line:{color,transparency:100}});
  s.addShape(S.ellipse,{x:10.8,y:4.7,w:3.2,h:3.2,fill:{color:C.cyan,transparency:76},line:{color:C.cyan,transparency:100}});
  s.addText(String(no).padStart(2,'0'),{x:0.75,y:0.8,w:1.5,h:1,fontSize:50,bold:true,color:C.cyan,margin:0});
  s.addText(title,{x:0.78,y:2.1,w:9.5,h:0.82,fontSize:34,bold:true,color:C.white,margin:0,fit:'shrink'});
  s.addText(subtitle,{x:0.82,y:3.25,w:8.8,h:0.7,fontSize:16,color:'C9D7EB',margin:0,fit:'shrink'});
  s.addText('AI 产品经理系统课程 · 第 00 课',{x:0.82,y:6.72,w:4,h:0.2,fontSize:9,color:'91A4BE',margin:0});
  return s;
}
function arrow(slide,x,y,w=0.55,color=C.blue){ slide.addShape(S.chevron,{x,y,w,h:0.42,fill:{color,transparency:12},line:{color,transparency:100}}); }

// 01 cover
{
  const s=pptx.addSlide(); s.background={color:C.navy};
  s.addShape(S.ellipse,{x:8.7,y:-1.5,w:6.2,h:6.2,fill:{color:C.blue,transparency:35},line:{color:C.blue,transparency:100}});
  s.addShape(S.ellipse,{x:10.6,y:3.6,w:4.2,h:4.2,fill:{color:C.cyan,transparency:58},line:{color:C.cyan,transparency:100}});
  addPill(s,'第 00 课',0.78,0.82,C.cyan,1.2);
  s.addText('课程说明会与学习地图',{x:0.78,y:1.6,w:9.8,h:0.88,fontSize:38,bold:true,color:C.white,margin:0,fit:'shrink'});
  s.addText('先看清地图，再开始学习',{x:0.82,y:2.8,w:7.5,h:0.45,fontSize:20,color:'C9D7EB',margin:0});
  s.addShape(S.line,{x:0.82,y:3.55,w:5.8,h:0,line:{color:C.cyan,width:3}});
  s.addText('系统知识  ×  项目实践  ×  求职转化',{x:0.82,y:3.9,w:7.4,h:0.36,fontSize:15,bold:true,color:C.cyan,margin:0});
  s.addText('AI 产品经理系统课程 · 第一阶段',{x:0.82,y:6.65,w:4.5,h:0.22,fontSize:10,color:'91A4BE',margin:0});
}
// 02 diagnostic
{
  const s=pptx.addSlide('MASTER'); addTitle(s,'你现在最大的困惑是什么？','开场诊断');
  const data=[['方向不清','不知道 AI 产品经理具体做什么',C.blue],['技术焦虑','名词很多，担心自己学不会',C.purple],['工具碎片','会用一些工具，但没有体系',C.cyan],['没有项目','学过概念，却没有完整产出',C.orange],['不会表达','经历和项目无法转成求职证据',C.green]];
  data.forEach((d,i)=>addCard(s,{x:0.7+(i%3)*4.15,y:1.62+Math.floor(i/3)*2.25,w:3.75,h:1.75,title:d[0],body:d[1],color:d[2],num:i+1}));
  s.addText('请选择最接近你当前状态的一项，并用 30 秒介绍自己的背景。',{x:5.0,y:6.1,w:7.2,h:0.35,fontSize:12,bold:true,color:C.navy,align:'right',margin:0});
}
// 03 goals
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'今天完成四件事','本课目标');
 const a=[['讲清初心','为什么要做这门课'],['看懂路径','三个阶段怎样递进'],['建立约定','如何更新、纠错与保证质量'],['开始行动','完成第一次 AI 场景思考']];
 a.forEach((d,i)=>addCard(s,{x:0.72+i*3.12,y:1.75,w:2.78,h:3.75,title:d[0],body:d[1]+'\n\n不是背术语，而是建立接下来学习的坐标。',color:[C.blue,C.purple,C.orange,C.green][i],num:i+1}));
}
// 04 section
sectionSlide(1,'为什么要做这门课','不是再教一套工具，而是补上一条适合产品经理的完整学习路径');
// 05 two biases
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'我调研到的课程，常见两种偏向','课程初心');
 addCard(s,{x:0.8,y:1.65,w:5.75,h:3.8,title:'偏技术原理',body:'模型、训练、微调、向量、召回讲得很深。\n\n产品经理听懂概念后，仍可能不知道怎样转化为需求、PRD、交互和项目决策。',color:C.purple,num:1});
 addCard(s,{x:6.78,y:1.65,w:5.75,h:3.8,title:'偏工具操作',body:'搭 Bot、做工作流、生成网页和原型。\n\n学员可能做出 Demo，却回答不了为什么要做、效果怎样验证、出了错谁负责。',color:C.orange,num:2});
 s.addText('说明：这是基于有限课程样本的个人观察，不代表对整个市场的绝对判断。',{x:1.0,y:5.9,w:11.3,h:0.35,fontSize:10.5,color:C.gray,italic:true,align:'center',margin:0});
}
// 06 not tool PM
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'AI 产品经理 ≠ AI 工具产品经理','核心立场');
 s.addShape(S.roundRect,{x:0.85,y:1.65,w:5.3,h:4.35,rectRadius:0.06,fill:{color:'EDF4FF'},line:{color:'BBD2FF'}});
 s.addText('会使用 AI 工具',{x:1.2,y:2.0,w:4.6,h:0.42,fontSize:22,bold:true,color:C.blue,align:'center',margin:0});
 s.addText('研究与写作\n搭建 Bot / Workflow\n生成原型与 Demo\n提高个人工作效率',{x:1.45,y:2.75,w:4.1,h:2.2,fontSize:16,color:C.ink,breakLine:false,align:'center',valign:'mid',margin:0.04,fit:'shrink'});
 s.addShape(S.roundRect,{x:7.15,y:1.65,w:5.3,h:4.35,rectRadius:0.06,fill:{color:'F0F9F4'},line:{color:'BFE6CE'}});
 s.addText('会设计 AI 产品',{x:7.5,y:2.0,w:4.6,h:0.42,fontSize:22,bold:true,color:C.green,align:'center',margin:0});
 s.addText('识别真实业务问题\n判断 AI 是否必要\n设计系统与异常路径\n对效果、风险和价值负责',{x:7.75,y:2.75,w:4.1,h:2.2,fontSize:16,color:C.ink,align:'center',valign:'mid',margin:0.04,fit:'shrink'});
 arrow(s,6.38,3.35,0.55,C.cyan);
}
// 07 tool value
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'工具仍然重要，但它是手段','课程初心');
 const a=[['研究','快速获得信息和线索'],['表达','把方案变成文档与图示'],['验证','低成本做原型和 Demo'],['提效','缩短重复劳动时间']];
 a.forEach((d,i)=>addCard(s,{x:0.75+i*3.12,y:1.8,w:2.75,h:3.5,title:d[0],body:d[1]+'\n\n关键追问：它帮助我们解决了什么产品问题？',color:[C.blue,C.purple,C.orange,C.green][i],num:i+1}));
 s.addText('课程不会排斥工具，也不会把工具熟练度当成岗位价值的全部。',{x:1.3,y:5.75,w:10.8,h:0.38,fontSize:15,bold:true,color:C.navy,align:'center',margin:0});
}
// 08 connect
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'课程试图连接技术与产品','课程初心');
 addCard(s,{x:0.8,y:1.7,w:3.5,h:3.8,title:'理解够用的技术',body:'知道模型、Prompt、Workflow、RAG、Agent 的作用、边界和基本机制。',color:C.purple,num:1});
 arrow(s,4.55,3.15,0.65,C.cyan);
 addCard(s,{x:5.2,y:1.7,w:3.5,h:3.8,title:'形成产品判断',body:'从用户、业务、流程、风险和价值出发，选择合适方案。',color:C.blue,num:2});
 arrow(s,8.95,3.15,0.65,C.cyan);
 addCard(s,{x:9.6,y:1.7,w:2.9,h:3.8,title:'推动真实落地',body:'用测试、反馈、成本与协作，把 Demo 推向可用产品。',color:C.green,num:3});
}
// 09 section
sectionSlide(2,'三阶段怎样设计','知识不是终点：最终要进入项目，再进入求职表达');
// 10 three phase overview
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'三阶段是一条成果链','课程设计');
 const a=[['第一阶段','系统学习','建立 AI 产品完整地图',C.blue],['第二阶段','项目实践','把知识变成可展示产出',C.orange],['第三阶段','求职转化','让市场看懂并相信能力',C.green]];
 a.forEach((d,i)=>{addCard(s,{x:0.85+i*4.15,y:1.75,w:3.55,h:3.8,title:d[0]+'｜'+d[1],body:d[2]+'\n\n核心判断：'+['你懂不懂？','你做没做？','你讲得清吗？'][i],color:d[3],num:i+1}); if(i<2)arrow(s,4.45+i*4.15,3.25,0.55,C.cyan)});
}
// 11 phase1
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'第一阶段｜建立体系化认知','系统学习');
 const items=[['基础','大模型、Prompt 与上下文'],['判断','场景与产品形态选择'],['方案','PRD、Workflow、RAG、Agent'],['表达','交互、原型与 Demo'],['验证','测试、Bad Case、风险治理'],['经营','模型选择、成本、ROI 与立项']];
 items.forEach((d,i)=>addCard(s,{x:0.72+(i%3)*4.18,y:1.55+Math.floor(i/3)*2.45,w:3.78,h:1.95,title:d[0],body:d[1],color:[C.blue,C.purple,C.orange,C.cyan,C.red,C.green][i]}));
}
// 12 hot map
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'热门内容很多，但不是名词堆砌','系统学习');
 const tags=['Prompt','Workflow','RAG','Agent','MCP','Function Calling','Vibe Coding','Bad Case','Token 成本','ROI'];
 tags.forEach((t,i)=>{const x=0.9+(i%5)*2.43,y=1.75+Math.floor(i/5)*1.05; s.addShape(S.roundRect,{x,y,w:2.05,h:0.64,rectRadius:0.12,fill:{color:[C.blue,C.purple,C.orange,C.green,C.cyan][i%5],transparency:8},line:{color:C.white,transparency:100}}); s.addText(t,{x:x+0.1,y:y+0.19,w:1.85,h:0.2,fontSize:12,bold:true,color:C.white,align:'center',margin:0,fit:'shrink'});});
 s.addShape(S.roundRect,{x:1.1,y:4.35,w:11.1,h:1.3,rectRadius:0.06,fill:{color:C.white},line:{color:C.line}});
 s.addText('学习目标',{x:1.45,y:4.7,w:1.4,h:0.28,fontSize:16,bold:true,color:C.navy,margin:0});
 s.addText('知道每个概念解决什么问题、彼此是什么关系、何时适合使用，以及怎样进入产品方案。',{x:3.0,y:4.67,w:8.6,h:0.42,fontSize:14,color:C.gray,margin:0,fit:'shrink'});
}
// 13 phase2
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'第二阶段｜用项目把知识变成能力','项目实践');
 const a=[['公共项目','所有人走通 0→1 完整闭环'],['方向项目 A','Workflow / RAG / Agent / 复杂 B 端'],['方向项目 B','选择与背景、目标岗位匹配的方向'],['主项目深化','补齐指标、风险、成本和追问']];
 a.forEach((d,i)=>{addCard(s,{x:0.7+i*3.14,y:1.75,w:2.78,h:3.75,title:d[0],body:d[1],color:[C.blue,C.purple,C.orange,C.green][i],num:i+1}); if(i<3)arrow(s,3.5+i*3.14,3.35,0.45,C.cyan)});
 s.addText('最终形成：3 个完整项目 + 1 个重点主项目',{x:2.2,y:5.85,w:8.9,h:0.42,fontSize:18,bold:true,color:C.navy,align:'center',margin:0});
}
// 14 group
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'为什么采用小组和不同项目？','项目实践');
 addBulletList(s,[
  {title:'避免同质化',body:'不是全班复制同一个知识库机器人，每个人保留独立选题和产出。'},
  {title:'扩大案例面',body:'不同背景带来不同业务、流程、用户和系统限制。'},
  {title:'共享 Bad Case',body:'同学之间不仅分享结果，也分享失败、修改和验证过程。'},
  {title:'训练协作表达',body:'展示、互评和追问，让方案从“自己觉得对”变成“别人能理解”。'}
 ],0.85,1.55,11.4,4.8,C.orange);
}
//15 truth
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'项目必须真实标注','项目实践');
 const a=[['课程项目','基于课堂题目完成，明确学习目的'],['个人项目','由个人发现问题并独立设计验证'],['业务改造方案','基于熟悉业务提出的方案，不冒充上线'],['真实工作项目','只写本人真实参与范围和可公开结果']];
 a.forEach((d,i)=>addCard(s,{x:0.72+i*3.12,y:1.72,w:2.76,h:3.65,title:d[0],body:d[1],color:[C.blue,C.purple,C.orange,C.green][i],num:i+1}));
 s.addShape(S.roundRect,{x:1.35,y:5.75,w:10.65,h:0.68,rectRadius:0.05,fill:{color:'FFF1F3'},line:{color:'FFD0D8'}});
 s.addText('底线：没有上线不写上线，没有真实数据不编数据，不把团队成果全部写成个人贡献。',{x:1.65,y:5.97,w:10.0,h:0.22,fontSize:12.5,bold:true,color:C.red,align:'center',margin:0,fit:'shrink'});
}
// 16 phase3
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'第三阶段｜把项目转化成求职证据','求职转化');
 const flow=[['定位','目标岗位与 JD'],['简历','经历与关键词证据'],['项目','主项目深度表达'],['作品集','辅助材料与图示'],['面试','追问与模拟反馈'],['投递','版本管理与复盘']];
 flow.forEach((d,i)=>{const x=0.55+i*2.12; s.addShape(S.ellipse,{x:x+0.5,y:2.05,w:0.72,h:0.72,fill:{color:[C.blue,C.purple,C.orange,C.cyan,C.green,C.red][i]},line:{color:C.white,width:2}}); s.addText(String(i+1),{x:x+0.5,y:2.27,w:0.72,h:0.2,fontSize:12,bold:true,color:C.white,align:'center',margin:0}); s.addText(d[0],{x:x,y:3.05,w:1.72,h:0.3,fontSize:15,bold:true,color:C.navy,align:'center',margin:0}); s.addText(d[1],{x:x,y:3.55,w:1.72,h:0.66,fontSize:10.5,color:C.gray,align:'center',valign:'mid',margin:0.02,fit:'shrink'}); if(i<5)arrow(s,x+1.75,2.2,0.35,C.line)});
 s.addText('目标不是“包装得像做过”，而是让真实能力更容易被看见、理解和追问。',{x:1.4,y:5.4,w:10.5,h:0.45,fontSize:16,bold:true,color:C.navy,align:'center',margin:0});
}
// 17 summary stages
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'三个阶段，分别回答三个问题','课程设计');
 const a=[['你懂不懂？','第一阶段建立体系与判断'],['你做没做？','第二阶段形成项目与证据'],['你讲得清吗？','第三阶段完成求职转化']];
 a.forEach((d,i)=>addCard(s,{x:0.85+i*4.15,y:1.8,w:3.55,h:3.7,title:d[0],body:d[1]+'\n\n三个阶段可以独立购买，但成果彼此衔接。',color:[C.blue,C.orange,C.green][i],num:i+1}));
}
// 18 section
sectionSlide(3,'一份透明的课程说明','AI 会变化，课程会更新；坦诚边界，但不降低质量标准');
// 19 stable/change
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'什么保持稳定，什么可能变化？','课程约定');
 addCard(s,{x:0.85,y:1.65,w:5.65,h:4.2,title:'保持稳定',body:'阶段解决的问题\n核心能力主线\n对学员承诺的交付\n已购买阶段的完整权益',color:C.green,num:'✓'});
 addCard(s,{x:6.82,y:1.65,w:5.65,h:4.2,title:'可能变化',body:'工具与模型版本\n案例与行业素材\n部分课次顺序\n正课与加餐的分配',color:C.orange,num:'↻'});
}
// 20 update mechanism
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'更新不是随意改课','课程约定');
 const flow=[['发现变化','工具、案例或方法失效'],['判断影响','是否影响核心能力和交付'],['说明调整','原因、范围和新旧关系'],['同步材料','版本号、勘误和新文件'],['收集反馈','验证调整是否有效']];
 flow.forEach((d,i)=>{const x=0.55+i*2.52; addCard(s,{x,y:1.85,w:2.15,h:3.5,title:d[0],body:d[1],color:[C.blue,C.purple,C.orange,C.green,C.cyan][i],num:i+1}); if(i<4)arrow(s,x+2.16,3.28,0.32,C.line)});
}
// 21 quality
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'坦诚能力边界，同时建立质量机制','课程约定');
 const a=[['来源追溯','核心知识点记录参考文件、页码或段落'],['实际验证','能实测的功能、Prompt 和流程尽量实测'],['专业审核','技术性强的内容接受研发/算法视角检查'],['公开勘误','事实、推断和观点分开；发现问题及时修订'],['持续更新','保留版本和修改原因，不用“变化快”掩盖错误']];
 a.forEach((d,i)=>addCard(s,{x:0.68+(i%3)*4.2,y:1.55+Math.floor(i/3)*2.45,w:3.8,h:1.95,title:d[0],body:d[1],color:[C.blue,C.purple,C.orange,C.green,C.cyan][i],num:i+1}));
}
//22 learning loop
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'学习不是听完，而是完成一个闭环','学习机制');
 const a=[['课前','带着真实问题'],['课中','使用模板判断'],['课后','在自己的场景重做'],['反馈','接受老师和同学质疑'],['修订','保留版本与修改原因']];
 a.forEach((d,i)=>{const x=0.55+i*2.52; s.addShape(S.ellipse,{x:x+0.55,y:1.85,w:0.95,h:0.95,fill:{color:[C.blue,C.purple,C.orange,C.green,C.cyan][i]},line:{color:C.white,width:2}}); s.addText(String(i+1),{x:x+0.55,y:2.15,w:0.95,h:0.2,fontSize:14,bold:true,color:C.white,align:'center',margin:0}); s.addText(d[0],{x,y:3.08,w:2.05,h:0.36,fontSize:17,bold:true,color:C.navy,align:'center',margin:0}); s.addText(d[1],{x,y:3.65,w:2.05,h:0.65,fontSize:11,color:C.gray,align:'center',valign:'mid',margin:0,fit:'shrink'}); if(i<4)arrow(s,x+2.08,2.1,0.38,C.line)});
 s.addText('眼睛会了、脑子理解了，不等于手能够完成。',{x:1.8,y:5.3,w:9.8,h:0.5,fontSize:20,bold:true,color:C.navy,align:'center',margin:0});
}
//23 template
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'模板是一组防遗漏问题','学习机制');
 const a=[['第一遍｜完整填写','先保证分析不漏项'],['第二遍｜按项目调整','删掉无关项，补充特殊约束'],['第三遍｜形成方法','沉淀自己的判断顺序和表达方式']];
 a.forEach((d,i)=>{addCard(s,{x:0.9+i*4.15,y:1.85,w:3.55,h:3.6,title:d[0],body:d[1]+'\n\n模板不是标准答案，最终要服务真实项目。',color:[C.blue,C.orange,C.green][i],num:i+1}); if(i<2)arrow(s,4.5+i*4.15,3.25,0.45,C.cyan)});
}
//24 compare expression
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'同一个助手，两种表达证明不同能力','岗位理解');
 addCard(s,{x:0.8,y:1.6,w:5.7,h:4.25,title:'工具表达',body:'“我用某个平台上传了文档，搭了一个客服机器人，还生成了一个页面。”\n\n证明：会操作工具、能够快速搭建。\n缺少：问题、方案选择、验证和结果。',color:C.orange,num:1});
 addCard(s,{x:6.83,y:1.6,w:5.7,h:4.25,title:'产品表达',body:'“我针对查询慢和口径不一，设计了资料来源、候选回答、人工确认和效果验证流程。”\n\n证明：从问题出发，能设计并验证产品。',color:C.green,num:2});
}
//25 value
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'AI 产品经理真正负责什么？','岗位理解');
 addBulletList(s,[
   {title:'判断问题',body:'用户和业务问题是否真实，AI 是否比普通方案更合适。'},
   {title:'设计方案',body:'组合模型、知识、规则、工具、流程、交互和人工。'},
   {title:'验证效果',body:'用样本、指标和 Bad Case 决定是否可用。'},
   {title:'推动落地',body:'处理风险、成本、协作、上线和持续运营。'}
 ],0.85,1.55,11.4,4.75,C.blue);
}
//26 section
sectionSlide(4,'全课程学习地图','今天只建立概念关系；后续课程逐个拆解、练习和验证');
//27 map
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'从基础能力到产品立项','学习地图');
 const a=[['基础','大模型\nPrompt'],['判断','场景\n产品形态'],['方案','PRD\nWorkflow/RAG/Agent'],['表达','交互\n原型/Demo'],['验证','测试\n风险治理'],['经营','成本\nROI/立项']];
 a.forEach((d,i)=>{const x=0.48+i*2.14; s.addShape(S.roundRect,{x,y:1.85,w:1.85,h:2.65,rectRadius:0.08,fill:{color:[C.blue,C.purple,C.orange,C.cyan,C.red,C.green][i]},line:{color:C.white,transparency:100}}); s.addText(d[0],{x:x+0.18,y:2.18,w:1.49,h:0.3,fontSize:17,bold:true,color:C.white,align:'center',margin:0}); s.addShape(S.line,{x:x+0.38,y:2.75,w:1.08,h:0,line:{color:C.white,transparency:35,width:1}}); s.addText(d[1],{x:x+0.18,y:3.08,w:1.49,h:0.8,fontSize:12.5,color:C.white,align:'center',valign:'mid',margin:0,fit:'shrink'}); if(i<5)arrow(s,x+1.88,2.95,0.28,C.line)});
 s.addText('每一层都要回到同一个问题：它帮助谁解决了什么真实问题？',{x:1.6,y:5.35,w:10.2,h:0.44,fontSize:17,bold:true,color:C.navy,align:'center',margin:0});
}
//28 definitions
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'四个热门概念，先记住一句话','学习地图');
 const a=[['Prompt','管理输入','把任务、背景、规则和输出要求告诉模型',C.blue],['Workflow','管理流程','把稳定业务拆成节点、分支和人工审核',C.purple],['RAG','管理知识','回答前检索企业或行业知识，并提供引用',C.orange],['Agent','管理任务执行','围绕目标使用记忆和工具完成多步骤任务',C.green]];
 a.forEach((d,i)=>addCard(s,{x:0.72+i*3.12,y:1.65,w:2.78,h:4.15,title:d[0]+'｜'+d[1],body:d[2]+'\n\n后续课程会学习适用边界、产品要求和常见失败。',color:d[3],num:i+1}));
}
//29 case
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'贯穿案例：本地生活商家客服与知识助手','案例预告');
 addPill(s,'公开业务信息 + 教学假设',9.75,0.72,C.orange,2.45);
 addCard(s,{x:0.78,y:1.55,w:3.62,h:4.15,title:'业务现状',body:'活动、商品、履约、售后规则较多\n资料分散在文档与消息中\n商家客服查询慢、回答口径不一',color:C.red,num:1});
 arrow(s,4.52,3.25,0.52,C.cyan);
 addCard(s,{x:5.05,y:1.55,w:3.3,h:4.15,title:'AI 可能介入',body:'理解问题意图\n查找相关规则与知识\n生成候选答案并附来源\n高风险问题转人工确认',color:C.blue,num:2});
 arrow(s,8.48,3.25,0.52,C.cyan);
 addCard(s,{x:9.0,y:1.55,w:3.52,h:4.15,title:'需要验证',body:'查询时长是否下降？\n答案是否被采纳？\n修改率和错误数怎样？\n维护知识的成本多高？',color:C.green,num:3});
}
//30 exercise
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'课堂练习：第一次 AI 场景思考','个人练习 · 12 分钟');
 const a=[['谁遇到什么问题？','限定一个角色、一次任务和一个主要结果'],['现在怎样解决？','写出现有流程、工具和最痛的环节'],['AI 帮助什么？','只选择一个具体介入点，并保留人工边界'],['如何验证？','定义一个效果指标、一个业务/体验指标和最小试点']];
 a.forEach((d,i)=>addCard(s,{x:0.72+i*3.12,y:1.55,w:2.78,h:4.3,title:d[0],body:d[1]+'\n\n不知道的信息请标记“待验证”。',color:[C.blue,C.purple,C.orange,C.green][i],num:i+1}));
}
//31 peer review
{
 const s=pptx.addSlide('MASTER'); addTitle(s,'互评只问三个问题','课堂练习');
 const a=[['场景具体吗？','是否明确到一个角色、一次任务和主要问题'],['AI 必要吗？','普通产品优化能否更低成本地解决'],['结果可验证吗？','“提高效率”能否转化为可观察变化']];
 a.forEach((d,i)=>addCard(s,{x:0.85+i*4.15,y:1.8,w:3.55,h:3.75,title:d[0],body:d[1],color:[C.blue,C.orange,C.green][i],num:i+1}));
 s.addText('点评重点不是方案听起来多高级，而是问题是否真实、选择是否有理由。',{x:1.25,y:5.85,w:10.8,h:0.4,fontSize:15,bold:true,color:C.navy,align:'center',margin:0});
}
//32 final
{
 const s=pptx.addSlide(); s.background={color:C.navy};
 s.addShape(S.ellipse,{x:9.6,y:-1.2,w:4.8,h:4.8,fill:{color:C.blue,transparency:42},line:{color:C.blue,transparency:100}});
 s.addText('今天带走四个结论',{x:0.82,y:0.7,w:7.8,h:0.65,fontSize:30,bold:true,color:C.white,margin:0});
 const a=['工具重要，但不等于产品判断','系统学习、项目实践、求职转化逐级递进','课程会更新，但来源、边界和质量机制必须透明','从一个真实、具体、可验证的问题开始'];
 a.forEach((t,i)=>{s.addShape(S.ellipse,{x:0.9,y:1.75+i*0.95,w:0.5,h:0.5,fill:{color:[C.blue,C.orange,C.green,C.cyan][i]},line:{color:C.white,width:1.5}}); s.addText(String(i+1),{x:0.9,y:1.9+i*0.95,w:0.5,h:0.14,fontSize:10,bold:true,color:C.white,align:'center',margin:0}); s.addText(t,{x:1.65,y:1.78+i*0.95,w:9.9,h:0.44,fontSize:16,color:C.white,margin:0,fit:'shrink'});});
 s.addShape(S.roundRect,{x:0.86,y:5.75,w:11.3,h:0.76,rectRadius:0.04,fill:{color:C.white,transparency:90},line:{color:C.white,transparency:100}});
 s.addText('课后作业：个人能力与学习目标表 + 第一次 AI 场景思考卡',{x:1.15,y:5.98,w:10.7,h:0.25,fontSize:13,bold:true,color:C.cyan,align:'center',margin:0});
 s.addText('下一课：AI 产品经理岗位认知与转型路径',{x:0.9,y:6.75,w:5.5,h:0.22,fontSize:10,color:'9DB0C9',margin:0});
}

pptx.writeFile({ fileName: process.argv[2] || 'lesson0_designed.pptx' });
