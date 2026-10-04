import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const workspace = "/Users/keivn/Project/AI-Course";
const outDir = path.join(workspace, "课程生产/第二阶段/专题课_PPT_便宜的答案昂贵的判断_完整版");
const buildDir = path.join(outDir, ".build");
const sourceDeck = path.join(workspace, "课程生产/第二阶段/第二课_PPT_电商智能客服项目实战_完整版/第二课_电商智能客服项目实战_完整PPT_v1.1.pptx");
const lecture = path.join(workspace, "课程生产/第二阶段/07_专题课_便宜的答案昂贵的判断_前五节老师讲义_v1.0.md");
const candidate = path.join(outDir, ".candidate.pptx");
const finalDeck = path.join(outDir, "专题课_便宜的答案昂贵的判断_完整PPT_v1.0.pptx");

const C = {
  bg: "031016", panel: "102B37", panel2: "133843", line: "416A78",
  teal: "36D6C2", coral: "FF7657", text: "F5F2E9", mute: "A7B5BC",
  yellow: "F0C565", green: "68D391", red: "EF6969",
};
const IN = 914400;
const W = 12191695;
const H = 6858000;
const FONT = "Hiragino Sans GB";
let shapeId = 2;

function emu(v) { return Math.round(v * IN); }
function esc(v) {
  return String(v ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}
function paragraphs(text, size, color, bold=false, align="l", linePct=112000) {
  const lines = String(text ?? "").split("\n");
  return lines.map(line => {
    if (!line) return `<a:p><a:pPr algn="${align}"/><a:endParaRPr lang="zh-CN" sz="${Math.round(size*100)}"/></a:p>`;
    return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="${linePct}"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr><a:r><a:rPr lang="zh-CN" sz="${Math.round(size*100)}" b="${bold?1:0}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${FONT}"/><a:ea typeface="${FONT}"/></a:rPr><a:t>${esc(line)}</a:t></a:r></a:p>`;
  }).join("");
}
function textbox(x,y,w,h,text,size=18,color=C.text,bold=false,align="l",anchor="t",margin=0.03) {
  const id=shapeId++;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${anchor}" lIns="${emu(margin)}" rIns="${emu(margin)}" tIns="${emu(margin)}" bIns="${emu(margin)}"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paragraphs(text,size,color,bold,align)}</p:txBody></p:sp>`;
}
function rect(x,y,w,h,fill=C.panel,line=C.line,lineWidth=0.8,round=true) {
  const id=shapeId++;
  const geom=round?"roundRect":"rect";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="${Math.round(lineWidth*12700)}">${line?`<a:solidFill><a:srgbClr val="${line}"/></a:solidFill>`:"<a:noFill/>"}</a:ln><a:effectLst/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}
function line(x,y,w,h,color=C.line,width=0.8) {
  const id=shapeId++;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Line ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="${Math.round(width*12700)}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></p:spPr></p:sp>`;
}
function arrow(x,y,w,h,color=C.teal) {
  const id=shapeId++;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Arrow ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}
function card(x,y,w,h,title,body,n,accent=C.teal,titleSize=17,bodySize=13.5) {
  const parts=[rect(x,y,w,h,C.panel,C.line,0.7,true)];
  let top=y+0.20;
  if(n){ parts.push(textbox(x+0.20,top,w-0.40,0.24,n,10.5,accent,true)); top+=0.34; }
  parts.push(textbox(x+0.20,top,w-0.40,0.42,title,titleSize,C.text,true,"l","ctr"));
  if(body) parts.push(textbox(x+0.20,top+0.50,w-0.40,h-(top-y)-0.66,body,bodySize,C.mute,false));
  return parts.join("");
}
function takeaway(label,body,y=6.23,accent=C.coral){
  return rect(0.82,y,11.52,0.61,C.panel2,C.line,0.6,true)+
    textbox(1.02,y+0.13,1.30,0.30,label,11.5,accent,true,"l","ctr")+
    textbox(2.30,y+0.10,9.68,0.36,body,14.8,C.text,true,"l","ctr");
}
function common(page,section,title,subtitle="",accent=C.teal){
  return textbox(0.76,0.20,10.75,0.34,section,14.5,C.teal,true,"l","ctr")+
    textbox(12.00,0.21,0.58,0.30,String(page).padStart(2,"0"),10.5,C.mute,true,"r","ctr")+
    rect(0.76,7.17,11.82,0.014,C.line,null,0,false)+
    rect(0.76,7.11,11.82*page/26,0.04,C.teal,null,0,false)+
    textbox(0.78,0.70,11.70,0.64,title,29,C.text,true,"l","ctr")+
    rect(0.78,1.38,1.46,0.045,accent,null,0,false)+
    (subtitle?textbox(0.78,1.50,11.65,0.38,subtitle,14.5,C.mute,false):"");
}
function slideXml(body){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${C.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
function notesXml(text){
  const ps=String(text).split("\n").map(t=>t?`<a:p><a:r><a:rPr lang="zh-CN"/><a:t>${esc(t)}</a:t></a:r></a:p>`:"<a:p/>").join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="3" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${ps}</p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" idx="5" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}
function quoteBox(x,y,w,h,label,text,accent=C.teal,size=18){
  return rect(x,y,w,h,C.panel,C.line,0.8,true)+textbox(x+0.25,y+0.20,w-0.50,0.28,label,11.5,accent,true)+textbox(x+0.25,y+0.64,w-0.50,h-0.86,text,size,C.text,true,"l","t");
}
function columns(items,y=2.10,h=3.75){
  const gap=0.25,w=(11.52-gap*(items.length-1))/items.length;
  return items.map((it,i)=>card(0.82+i*(w+gap),y,w,h,it[0],it[1],String(i+1).padStart(2,"0"),it[2]||C.teal,it[3]||17,it[4]||13.5)).join("");
}
function flow(items,y=2.55,h=2.35){
  const gap=0.18,w=(11.50-gap*(items.length-1))/items.length;
  let out="";
  items.forEach((it,i)=>{
    const x=0.84+i*(w+gap);
    out+=card(x,y,w,h,it[0],it[1],String(i+1).padStart(2,"0"),it[2]||C.teal,it[3]||15.5,it[4]||12.3);
    if(i<items.length-1) out+=arrow(x+w-0.02,y+h/2-0.08,gap+0.05,0.16,C.teal);
  });
  return out;
}
function row(y,label,body,accent=C.teal){
  return textbox(0.94,y,2.30,0.44,label,17,accent,true,"l","ctr")+
    textbox(3.18,y,8.86,0.50,body,16,C.text,false,"l","ctr")+
    rect(0.94,y+0.61,11.18,0.012,C.line,null,0,false);
}

const slides=[];
function add(body,note){ slides.push({xml:slideXml(body),note:`${note}\n\n讲义来源：${lecture}`}); }

// 01 封面
add(
  rect(0.28,0,0.16,7.50,C.teal,null,0,false)+
  textbox(0.94,0.66,7.5,0.30,"AI PRODUCT MANAGER · PORTFOLIO & INTERVIEW",12,C.teal,true)+
  textbox(0.94,1.48,10.55,1.36,"便宜的答案，昂贵的判断",42,C.text,true)+
  textbox(0.96,3.00,9.20,0.48,"AI 产品经理如何证明自己的专业性",21,C.text,true)+
  rect(0.96,3.72,2.12,0.05,C.coral,null,0,false)+
  textbox(0.96,4.03,9.30,0.60,"从智能客服项目复盘，到作品集打造与面试答辩",17,C.mute,true)+
  textbox(10.25,4.60,1.90,1.15,"03",68,C.teal,true,"c","ctr")+
  textbox(0.96,6.70,5.50,0.26,"第二阶段 · 专题课",11.5,C.mute,true),
  "开场不要先给模板。先抛出核心问题：当 AI 可以快速生成漂亮的作品集时，专业性还体现在哪里？"
);

// 02 A
add(common(2,"第一节｜AI 生成的作品集能说明什么","项目介绍 A","请先判断：它专业吗？")+
  quoteBox(1.14,2.30,11.05,2.90,"版本 A","我使用 Coze 搭建了一个电商智能客服，项目使用了大模型、知识库、Workflow 和 Agent，可以回答用户问题、查询订单、处理售后和推荐商品，提高客服效率和用户体验。",C.teal,23)+
  takeaway("先不解释","请学员只判断：这段介绍说明了什么？还缺少什么？",5.55),
  "只展示版本 A，不提前讲评价标准。让学员说出第一印象。重点观察大家会不会被工具名称和功能数量影响。"
);

// 03 B
add(common(3,"第一节｜AI 生成的作品集能说明什么","项目介绍 B","哪一段更像成熟的 AI 产品经理表达？")+
  quoteBox(1.05,2.16,11.20,3.34,"版本 B","我设计了一款面向电商售后场景的智能客服，主要解决退款政策查询、订单确认和售后申请过程中信息分散、重复沟通的问题。\n\n产品将大模型、知识库、业务流程和人工客服进行分工，帮助用户完成从问题咨询到售后申请预受理的完整流程。",C.coral,20)+
  takeaway("课堂投票","认为 A 更专业的发 A，认为 B 更专业的发 B。",5.75),
  "让学员投票。多数人会选择 B，因为场景、问题和分工更完整。此时先肯定观察，不急着否定 B。"
);

// 04 对比揭示
add(common(4,"第一节｜AI 生成的作品集能说明什么","B 只是几秒钟生成的","真正的问题：哪些来自真实项目，哪些只是合理补全？",C.coral)+
  columns([
    ["项目事实","实际做过，可以查到\n\n例：已经接入订单系统",C.teal,18,15],
    ["产品判断","候选人主动做出的选择\n\n例：模型不直接批准退款",C.coral,18,15],
    ["暂未确认","听起来合理，但还不知道依据\n\n例：显著提升用户满意度",C.yellow,18,15],
  ],2.20,3.50)+takeaway("拆解方法","逐句追问来源、实现范围和运行情况。",5.96),
  "揭示版本 B 来自 AI 优化。不要说 B 是错的，而是带学员逐句判断：属于项目事实、产品判断，还是暂未确认。"
);

// 05 四支点
add(common(5,"第二节｜专业性的四个支点","专业性的四个支点","后续作品集与面试共用同一套判断框架")+
  columns([
    ["看得准","找到具体、真实、值得解决的任务",C.teal,20,15],
    ["选得明","重要设计都能说明为什么",C.coral,20,15],
    ["验得过","失败过、修改过，也重新检查过",C.yellow,20,15],
    ["说得清","准确说明已经上线到什么程度",C.green,20,15],
  ],2.30,3.32)+takeaway("统一框架","表达完整不等于专业。专业来自问题、取舍、验证和边界。",5.86),
  "四个支点是整堂课的主框架。看得准对应问题理解，选得明对应关键判断，验得过对应验证过程，说得清对应边界意识。"
);

// 06 快速练习
add(common(6,"第二节｜专业性的四个支点","四句话快速判断","每句话分别对应哪一个支点？")+
  row(2.10,"句子 1","用户表面问能不能退款，实际任务是确认具体订单是否符合条件。",C.teal)+
  row(2.92,"句子 2","我没有让大模型直接判断退款资格，因为错误判断会影响用户权益。",C.coral)+
  row(3.74,"句子 3","用户更换订单后仍沿用旧状态，我修改后重新测试了三个相关场景。",C.yellow)+
  row(4.56,"句子 4","产品已经上线，但当前样本不足以判断长期成本变化。",C.green)+
  takeaway("参考答案","看得准 · 选得明 · 验得过 · 说得清",5.66),
  "先让学员回答，再公布参考。最后让每个人判断自己的项目目前最弱的是哪一项。"
);

// 07 传统 vs AI
add(common(7,"第三节｜AI 项目作品集的结构与呈现","传统产品作品集与 AI 产品作品集","AI 项目保留产品设计主线，同时补充不确定性处理")+
  rect(0.82,2.05,5.57,3.80,C.panel,C.teal,0.9,true)+textbox(1.10,2.28,5.00,0.42,"传统产品作品集",20,C.teal,true)+
  textbox(1.10,2.96,4.94,2.42,"业务问题与产品目标\n用户流程与功能结构\n页面与交互设计\n系统与接口说明\n上线结果与个人贡献",16,C.text,false)+
  rect(6.70,2.05,5.64,3.80,C.panel,C.coral,0.9,true)+textbox(6.98,2.28,5.02,0.42,"AI 产品作品集",20,C.coral,true)+
  textbox(6.98,2.96,4.98,2.42,"保留传统产品结构\n增加 AI 能力分工\n展示 Bad Case 与修改\n说明低置信度和人工接管\n持续观察模型与流程表现",16,C.text,false)+
  takeaway("核心区别","AI 产品还要说明 AI 负责什么、不负责什么，出现不确定结果时怎么办。",6.08),
  "轻量对比即可。强调 AI 产品作品集没有推翻传统结构，产品形态、流程和页面仍然是主体。"
);

// 08 四问题
add(common(8,"第三节｜AI 项目作品集的结构与呈现","作品集先回答四个问题","面试官按照理解项目的顺序阅读")+
  columns([
    ["这是什么产品","产品定位、目标用户与 MVP",C.teal,18,15],
    ["用户怎样使用","核心流程、关键页面与状态",C.coral,18,15],
    ["为什么这样设计","能力分工与关键产品决策",C.yellow,18,15],
    ["实际运行怎样","修改过程、上线结果与限制",C.green,18,15],
  ],2.38,3.20)+takeaway("阅读逻辑","产品是主体，判断解释设计，结果说明项目实际进展。",5.90),
  "不要先让学员背八部分。先让他们理解，作品集必须回答产品、使用、设计原因和实际运行四个问题。"
);

// 09 八部分
add(common(9,"第三节｜AI 项目作品集的结构与呈现","完整作品集的八个部分","顺序按照面试官理解项目的过程排列")+
  flow([
    ["项目概览","项目全貌",C.teal], ["业务问题","为什么做",C.teal], ["定位与 MVP","第一版边界",C.yellow], ["产品形态","怎样完成任务",C.coral]
  ],2.08,1.55)+
  flow([
    ["页面与状态","产品怎样使用",C.teal], ["能力分工","谁负责什么",C.yellow], ["决策与迭代","为什么这样设计",C.coral], ["结果与贡献","运行情况与个人工作",C.green]
  ],4.16,1.55),
  "用一页建立八部分全貌。第一部分快速了解项目，第二三部分解释为什么做，第四五部分展示产品，第六七部分展开方案和迭代，第八部分说明结果与贡献。"
);

// 10 项目概览
add(common(10,"第三节｜AI 项目作品集的结构与呈现","第一页项目概览","两分钟内让人看懂产品、范围和个人角色")+
  rect(0.82,2.02,7.28,3.95,C.panel,C.teal,0.8,true)+
  textbox(1.10,2.26,6.75,0.38,"电商售后智能服务与业务办理助手",20,C.text,true)+
  textbox(1.10,2.90,6.64,0.82,"面向电商售后用户，通过对话助手、人工客服工作台和订单系统连接，帮助用户完成政策查询、订单确认和售后申请预受理。",16,C.mute,false)+
  textbox(1.10,4.02,6.55,0.34,"核心流程",12,C.coral,true)+
  textbox(1.10,4.48,6.58,0.82,"提出诉求  确认订单  补齐信息  检查条件  用户确认  预受理或转人工",16,C.text,true)+
  rect(8.38,2.02,3.96,3.95,C.panel2,C.line,0.8,true)+
  textbox(8.66,2.28,3.40,0.36,"个人角色",18,C.teal,true)+
  textbox(8.66,2.92,3.34,2.52,"业务流程梳理\nMVP 范围定义\n核心交互设计\nAI 能力分工\n测试用例设计\n上线问题分析",15,C.text,false),
  "项目概览只放最重要的信息。后续页面再展开业务问题、产品设计、能力分工和上线结果。"
);

// 11 业务问题与MVP
add(common(11,"第三节｜AI 项目作品集的结构与呈现","业务问题、产品目标与 MVP","先把为什么做和第一版做到哪里讲清楚")+
  rect(0.82,2.06,5.57,3.88,C.panel,C.teal,0.8,true)+textbox(1.10,2.30,5.02,0.38,"业务问题与产品目标",19,C.teal,true)+
  textbox(1.10,2.96,4.98,2.52,"FAQ 只能解释统一规则，无法根据订单状态继续办理。\n\n人工客服需要重复询问订单、商品和售后原因。\n\n目标：让用户完成低风险售后问题的查询与预受理。",15.5,C.text,false)+
  rect(6.70,2.06,5.64,3.88,C.panel,C.coral,0.8,true)+textbox(6.98,2.30,5.02,0.38,"第一版边界",19,C.coral,true)+
  textbox(6.98,2.96,5.00,2.52,"本期完成\n政策查询、订单确认、信息补齐、预受理\n\n本期暂不处理\n自动批准退款、高额赔偿、复杂投诉判责、特殊承诺",15.5,C.text,false)+
  takeaway("写法原则","业务问题说明为什么做，MVP 说明第一版做到哪一步。",6.13),
  "不要写行业发展史。用具体用户任务和原流程卡点说明问题，同时明确第一版不处理什么。"
);

// 12 产品形态和流程
add(common(12,"第三节｜AI 项目作品集的结构与呈现","产品形态与核心流程","作品集的视觉中心")+
  textbox(0.92,2.06,11.40,0.44,"产品形态：用户端对话助手 ＋ 人工客服工作台 ＋ 订单与售后系统连接",18,C.teal,true,"c","ctr")+
  flow([
    ["提出诉求","政策咨询或业务办理",C.teal,14.5,11.7],
    ["确认订单","选择订单与商品",C.teal,14.5,11.7],
    ["补齐信息","售后原因与必要字段",C.yellow,14.5,11.7],
    ["检查条件","规则与接口返回",C.coral,14.5,11.7],
    ["确认结果","预受理或转人工",C.green,14.5,11.7],
  ],2.90,2.38)+takeaway("展示重点","主流程控制在 6—8 步，复杂异常分支放到后续页面。",5.78),
  "先明确产品形态，再展示一条核心流程。不要只写智能客服系统，也不要在概览页塞入全部异常分支。"
);

// 13 页面和状态
add(common(13,"第三节｜AI 项目作品集的结构与呈现","核心页面与关键状态","既展示用户怎样完成任务，也展示异常时怎样处理")+
  columns([
    ["用户对话页面","当前任务\n已确认信息\n待补充内容\n允许修改选择",C.teal,18,14],
    ["订单确认页面","订单、商品、时间与金额\n用户明确选择\n更换订单后旧状态失效",C.coral,18,14],
    ["人工客服工作台","用户诉求摘要\n订单与售后信息\n转交原因\n人工继续处理",C.yellow,18,14],
  ],2.08,3.30)+
  textbox(0.94,5.67,11.30,0.35,"关键状态：信息不完整 · 用户改口 · 找不到订单 · 接口失败 · 低置信度 · 超权限请求 · 转人工",15,C.green,true,"c","ctr"),
  "推荐展示 2—3 张核心页面。页面旁边标注设计重点，并补充正常流程以外的关键状态。"
);

// 14 能力分工与判断
add(common(14,"第三节｜AI 项目作品集的结构与呈现","AI 方案、能力分工与关键判断","解释页面上的关键动作分别由谁完成")+
  row(2.02,"大模型","理解表达、识别任务、提取信息；不直接批准退款",C.teal)+
  row(2.72,"知识库","提供政策和商品说明；不提供实时订单状态",C.yellow)+
  row(3.42,"Workflow","控制补问、确认、提交和异常分支；不自由生成业务规则",C.coral)+
  row(4.12,"确定性规则","检查时间、金额、字段和资格条件；不理解复杂自然语言",C.green)+
  row(4.82,"人工客服","处理复杂、高风险和超权限问题；不重复询问已确认信息",C.teal)+
  takeaway("关键判断","让能力分工对应到具体页面和流程，而不是堆一张技术架构图。",5.74),
  "能力分工不是术语清单。每一项都要说明负责什么和不负责什么，并能解释为什么这样划分。"
);

// 15 结果、限制、练习
add(common(15,"第三节｜AI 项目作品集的结构与呈现","上线结果、当前限制与个人贡献","准确交代项目进展，不把上线等同于长期价值")+
  columns([
    ["上线结果","已经接入真实订单和售后系统\n上线政策查询、订单确认、信息补齐和预受理",C.teal,18,14],
    ["当前限制","上线周期和样本规模有限\n暂时不能判断长期成本和满意度变化",C.coral,18,14],
    ["个人贡献","流程、MVP、核心交互、能力分工、测试设计与问题分析",C.yellow,18,14],
  ],2.10,3.23)+
  takeaway("课堂练习","按照八部分，每部分先写一句话。",5.72),
  "所有数字使用实际统计。没有统计过就不临时估算。上线结果回答做到了什么，限制说明还有什么没有完成，贡献说明本人承担了什么。"
);

// 16 常见问题
add(common(16,"第四节｜90 秒项目介绍","90 秒介绍的四个常见问题","项目介绍要建立主线，不是把作品集快速念一遍")+
  columns([
    ["从行业背景讲起","讲了很久，仍不知道候选人做了什么",C.teal,17,14],
    ["罗列工具名称","只说明实现方式，没有说明产品任务",C.coral,17,14],
    ["逐项介绍功能","听不出用户怎样完成任务",C.yellow,17,14],
    ["只讲最终结果","范围、取舍和结果口径都不清楚",C.green,17,14],
  ],2.22,3.30)+takeaway("目标","让面试官迅速听懂项目主线，并愿意沿着关键节点继续追问。",5.92),
  "这里不再讲为什么需要 90 秒介绍，直接从常见问题切入。"
);

// 17 五段结构
add(common(17,"第四节｜90 秒项目介绍","五段式表达结构","每一段只保留一个重点")+
  flow([
    ["15 秒","项目与问题\n为什么做",C.teal,15,12],
    ["20 秒","产品形态与流程\n做成什么",C.teal,15,12],
    ["20 秒","关键产品判断\n为什么这样做",C.yellow,15,12],
    ["20 秒","问题与修改\n怎样迭代",C.coral,15,12],
    ["15 秒","上线、贡献与限制\n做到什么程度",C.green,15,12],
  ],2.42,2.72)+takeaway("使用方法","结构帮助检查主线，不要求逐字背诵。",5.66),
  "依次讲项目与问题、产品形态、关键判断、典型问题和上线范围。完整表达控制在 250—320 字。"
);

// 18 示范
add(common(18,"第四节｜90 秒项目介绍","智能客服项目示范","一条清楚的项目故事")+
  row(2.00,"项目与问题","电商售后智能服务，优先解决低风险售后问题。",C.teal)+
  row(2.72,"产品与流程","对话助手连接人工工作台及订单系统，完成预受理或转人工。",C.teal)+
  row(3.44,"关键判断","大模型理解表达，确定性规则判断退款资格。",C.yellow)+
  row(4.16,"问题与修改","更换订单后旧状态未失效，增加状态清空并重新检查。",C.coral)+
  row(4.88,"上线与限制","核心流程已在真实环境运行，长期价值仍需更多样本。",C.green)+
  takeaway("追问入口","为什么这样分工？状态问题怎样定位？上线后观察什么？",5.74),
  "老师完整讲一遍 90 秒示范，再让学员找出产品形态、关键判断和追问入口。逐字稿在老师讲义中。"
);

// 19 保留删掉
add(common(19,"第四节｜90 秒项目介绍","90 秒里保留什么、删掉什么","开场只保留能够建立项目主线的内容")+
  rect(0.82,2.10,5.57,3.75,C.panel,C.teal,0.9,true)+textbox(1.10,2.34,5.02,0.38,"建议保留",19,C.teal,true)+
  textbox(1.10,2.98,4.92,2.48,"具体业务问题\n一条核心用户流程\n一个关键产品判断\n一个典型问题与修改\n当前上线范围与个人贡献",16,C.text,false)+
  rect(6.70,2.10,5.64,3.75,C.panel,C.coral,0.9,true)+textbox(6.98,2.34,5.02,0.38,"删掉或后置",19,C.coral,true)+
  textbox(6.98,2.98,4.96,2.48,"宏观行业背景\n完整功能清单\n所有工具和模型名称\n每一个页面细节\n没有统计支持的效果描述",16,C.text,false)+
  takeaway("判断标准","能够进入后续追问的内容，不需要全部塞进开场。",6.08),
  "90 秒介绍的任务不是把项目讲完。超时后优先删除背景、工具名称和并列功能，不删除关键判断和上线范围。"
);

// 20 练习
add(common(20,"第四节｜90 秒项目介绍","五行提纲与点评标准","先写提纲，再连成完整表达")+
  rect(0.82,2.00,7.15,4.05,C.panel,C.teal,0.8,true)+
  textbox(1.10,2.26,6.58,0.36,"学员先写五行",18,C.teal,true)+
  textbox(1.10,2.90,6.48,2.78,"1  我做的是______，主要解决______\n2  产品由______组成，用户通过______完成任务\n3  我决定______，没有选择______，因为______\n4  我们发现______，后来调整了______\n5  目前已经上线______，当前限制是______",15.3,C.text,false)+
  rect(8.22,2.00,4.12,4.05,C.panel2,C.coral,0.8,true)+
  textbox(8.50,2.26,3.56,0.36,"老师只看五项",18,C.coral,true)+
  textbox(8.50,2.92,3.46,2.62,"问题是否具体\n产品是否可见\n取舍是否明确\n迭代是否真实\n范围是否准确",16,C.text,false),
  "让学员先写五行再计时讲述。点评不要逐句改语言，只看问题、产品、取舍、迭代和范围。"
);

// 21 四入口
add(common(21,"第五节｜项目面试的追问逻辑","项目介绍留下四个追问入口","面试不是随机题库，而是连续核对")+
  columns([
    ["业务问题","这个问题具体怎样发生？",C.teal,18,15],
    ["关键选择","当时为什么这样选择？",C.coral,18,15],
    ["典型问题","修改后怎样确认解决？",C.yellow,18,15],
    ["上线结果","范围、结果和个人贡献是什么？",C.green,18,15],
  ],2.30,3.18)+takeaway("面试官观察","前后信息是否一致，选择是否合理，细节能否继续展开。",5.84),
  "90 秒介绍中的每个重点都可能成为下一轮追问入口。面试官不会因为一句话专业就直接下结论。"
);

// 22 五层追问
add(common(22,"第五节｜项目面试的追问逻辑","一个项目点的五层追问","围绕同一个关键点逐层深入")+
  flow([
    ["事实层","当时具体发生了什么",C.teal,14.5,11.5],
    ["选择层","你最终决定了什么",C.teal,14.5,11.5],
    ["替代层","为什么没选其他方案",C.yellow,14.5,11.5],
    ["验证层","怎样知道方案能工作",C.coral,14.5,11.5],
    ["边界层","什么情况下不成立",C.green,14.5,11.5],
  ],2.48,2.76)+takeaway("准备方法","不要只准备标准答案，要准备一条能够向下展开的判断链。",5.74),
  "五层分别是事实、选择、替代、验证和边界。后一层不能推翻前一层的项目范围。"
);

// 23 连续示范
add(common(23,"第五节｜项目面试的追问逻辑","围绕一个关键选择连续追问","示例：为什么不让大模型直接判断退款资格？")+
  row(1.96,"第一问","为什么不能让大模型判断？",C.teal)+
  row(2.60,"第二问","既然规则更稳定，为什么不全部使用规则？",C.teal)+
  row(3.24,"第三问","规则具体判断哪些条件？",C.yellow)+
  row(3.88,"第四问","怎样测试这种能力分工？",C.coral)+
  row(4.52,"第五问","上线后出现过什么问题？",C.coral)+
  row(5.16,"第六问","这种方案还有什么限制？",C.green)+
  takeaway("一致性","所有回答来自同一套产品设计，角色分工和上线范围不能前后变化。",5.90),
  "逐问展开参考回答。重点不是背答案，而是看同一套设计能否从原因讲到实现、测试、运行和限制。"
);

// 24 五步回答
add(common(24,"第五节｜项目面试的追问逻辑","回答“为什么这样做”的五步结构","先给结论，再还原当时的选择过程")+
  flow([
    ["先给结论","最终选择是什么",C.teal,14.5,11.5],
    ["说明约束","用户风险、规则与范围",C.teal,14.5,11.5],
    ["比较方案","还考虑过什么",C.yellow,14.5,11.5],
    ["说明落地","怎样进入流程和页面",C.coral,14.5,11.5],
    ["结果与限制","做到什么程度",C.green,14.5,11.5],
  ],2.45,2.78)+takeaway("使用边界","不是每个问题都机械回答五步，结构只负责检查是否缺少关键环节。",5.76),
  "回答必须回到具体产品流程。不要讲两分钟背景让面试官猜结论，也不要只有设计没有范围。"
);

// 25 普通 vs 专业
add(common(25,"第五节｜项目面试的追问逻辑","普通回答与专业回答","问题：为什么退款资格不用大模型判断？")+
  rect(0.82,2.08,4.20,3.90,C.panel,C.coral,0.9,true)+textbox(1.10,2.34,3.65,0.36,"普通回答",19,C.coral,true)+
  textbox(1.10,3.10,3.56,1.10,"因为大模型可能产生幻觉，所以重要业务不能使用大模型。",19,C.text,true)+
  textbox(1.10,4.62,3.60,0.84,"只有正确观点\n没有项目约束和落地方式",14,C.mute,false)+
  rect(5.30,2.08,7.04,3.90,C.panel,C.teal,0.9,true)+textbox(5.58,2.34,6.50,0.36,"更完整的回答",19,C.teal,true)+
  textbox(5.58,2.98,6.36,2.48,"退款影响用户权益，资格条件又能明确表达。\n\n大模型负责理解表达和提取字段，规则检查订单状态、申请时间和商品限制。\n\n这套分工适合低风险售后，复杂投诉和特殊补偿仍转人工。",15.5,C.text,false)+
  takeaway("区别","普通回答停在观点，专业回答还原约束、比较、落地方式和适用范围。",6.13),
  "通过对比说明：成熟回答不仅观点正确，还能还原当时的选择空间和方案边界。"
);

// 26 练习与收束
add(common(26,"第五节｜项目面试的追问逻辑","把一个关键判断准备到足够深","全课练习与收束")+
  rect(0.82,2.05,7.18,3.96,C.panel,C.teal,0.8,true)+
  textbox(1.10,2.30,6.60,0.36,"选择一个项目点，连续回答五行",18,C.teal,true)+
  textbox(1.10,2.92,6.52,2.66,"当时的具体情况是______\n我最终决定______\n还考虑过______\n我通过______进行了检查\n当前方案不适用于______",16,C.text,false)+
  rect(8.24,2.05,4.10,3.96,C.panel2,C.coral,0.8,true)+
  textbox(8.52,2.30,3.54,0.36,"全课结论",18,C.coral,true)+
  textbox(8.52,3.00,3.48,2.10,"AI 可以降低表达成本。\n\n真正有价值的，仍然是你在项目中作出的判断。",20,C.text,true,"l","ctr")+
  takeaway("四个支点","看得准 · 选得明 · 验得过 · 说得清",6.14),
  "全课回扣标题。AI 可以帮助整理结构和语言，但每一个关键判断都要能够从具体问题讲到产品选择、测试修改和结果范围。"
);

if (slides.length !== 26) throw new Error(`Expected 26 slides, got ${slides.length}`);

fs.rmSync(buildDir,{recursive:true,force:true});
fs.mkdirSync(buildDir,{recursive:true});
execFileSync("unzip",["-q",sourceDeck,"-d",buildDir]);

// Keep only the slide and notes parts used by this lesson. The reference deck
// contains additional slides that should not remain hidden in the final file.
for (const relativeDir of [
  "ppt/slides",
  "ppt/slides/_rels",
  "ppt/notesSlides",
  "ppt/notesSlides/_rels",
]) {
  const absoluteDir = path.join(buildDir, relativeDir);
  for (const name of fs.readdirSync(absoluteDir)) {
    const match = name.match(/^(?:slide|notesSlide)(\d+)\.xml(?:\.rels)?$/);
    if (match && Number(match[1]) > slides.length) {
      fs.rmSync(path.join(absoluteDir, name), {force: true});
    }
  }
}

const contentTypesPath = path.join(buildDir, "[Content_Types].xml");
let contentTypes = fs.readFileSync(contentTypesPath, "utf8");
contentTypes = contentTypes.replace(
  /<Override PartName="\/ppt\/(?:slides\/slide|notesSlides\/notesSlide)(\d+)\.xml"[^>]*\/>/g,
  (part, number) => Number(number) > slides.length ? "" : part,
);
fs.writeFileSync(contentTypesPath, contentTypes);

for(let i=1;i<=slides.length;i++){
  fs.writeFileSync(path.join(buildDir,`ppt/slides/slide${i}.xml`),slides[i-1].xml);
  fs.writeFileSync(path.join(buildDir,`ppt/notesSlides/notesSlide${i}.xml`),notesXml(slides[i-1].note));
}

const presPath=path.join(buildDir,"ppt/presentation.xml");
let pres=fs.readFileSync(presPath,"utf8");
const list=pres.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
if(!list) throw new Error("slide list not found");
const ids=[...list[1].matchAll(/<p:sldId[^>]*\/>/g)].map(m=>m[0]).slice(0,26).join("");
pres=pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,`<p:sldIdLst>${ids}</p:sldIdLst>`);
fs.writeFileSync(presPath,pres);

const presRelsPath=path.join(buildDir,"ppt/_rels/presentation.xml.rels");
let presRels=fs.readFileSync(presRelsPath,"utf8");
presRels=presRels.replace(
  /<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*Target="slides\/slide(\d+)\.xml"\s*\/>/g,
  (relationship, number) => Number(number) > slides.length ? "" : relationship,
);
fs.writeFileSync(presRelsPath,presRels);

fs.rmSync(candidate,{force:true});
execFileSync("zip",["-q","-r","-X",candidate,"."],{cwd:buildDir,env:{...process.env,COPYFILE_DISABLE:"1"}});
fs.copyFileSync(candidate,finalDeck);
console.log(finalDeck);
