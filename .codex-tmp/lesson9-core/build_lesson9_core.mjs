import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const MD = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第9课 - AI产品评测与数据反馈.md";
const OUT = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第09课_正式PPT_核心结构版.pptx";
const PREVIEW = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-core/rendered";
const MONTAGE = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-core/contact.webp";

const C = { bg:"#07151F", panel:"#0B202C", panel2:"#102B37", ink:"#F5F2E9", muted:"#A7B5BC", teal:"#36D6C2", teal2:"#168E85", orange:"#FF7657", yellow:"#F0C75E", line:"#294550", green:"#6EE7B7", red:"#FF8A7A", dark:"#031016" };
const W = 1280, H = 720;

function shape(slide, geometry, p, fill="none", line="none", width=0, radius=0) {
  const cfg = { geometry, position:p, fill, line:{style:"solid", fill:line, width} };
  if (radius) cfg.borderRadius = radius;
  return slide.shapes.add(cfg);
}
function text(slide, value, p, size=22, color=C.ink, bold=false, align="left", family="PingFang SC") {
  const s = shape(slide,"textbox",p);
  s.text = value;
  s.text.style = {fontSize:size,color,bold,alignment:align,fontFamily:family};
  return s;
}
function line(slide,x1,y1,x2,y2,color=C.line,width=2) {
  return shape(slide,"line",{left:x1,top:y1,width:x2-x1,height:y2-y1},"none",color,width);
}
function arrow(slide,x1,y1,x2,y2,color=C.teal,width=3) {
  const s=line(slide,x1,y1,x2,y2,color,width); s.endArrowType="triangle"; return s;
}
function title(slide,t,sub="") {
  text(slide,t,{left:70,top:72,width:1130,height:58},40,C.ink,true);
  shape(slide,"rect",{left:70,top:148,width:160,height:4},C.teal);
  if(sub) text(slide,sub,{left:250,top:136,width:930,height:28},17,C.muted,false,"right");
}
function chrome(slide,i,total,phase) {
  text(slide,`AI PRODUCT · ${phase}`,{left:70,top:27,width:360,height:22},13,C.teal,true);
  text(slide,String(i+1).padStart(2,"0"),{left:1160,top:27,width:48,height:22},13,C.muted,true,"right");
  shape(slide,"rect",{left:70,top:683,width:1140,height:2},C.line);
  shape(slide,"rect",{left:70,top:683,width:1140*(i+1)/total,height:2},C.teal);
}
function note(slide,d){ slide.speakerNotes.textFrame.setText(`[Sources]\n- ${MD}\n- 内容依据：${d.source || d.title}`); slide.speakerNotes.setVisible(true); }
function bulletList(slide,items,x=95,y=205,w=1060,size=23,gap=47,color=C.ink){
  items.forEach((it,j)=>{ text(slide,"—",{left:x,top:y+j*gap,width:38,height:30},size,C.orange,true); text(slide,it,{left:x+52,top:y-1+j*gap,width:w-52,height:38},size,color,false); });
}
function quote(slide,value,y=300){ shape(slide,"roundRect",{left:86,top:y,width:1108,height:86},C.panel,C.line,1,12); shape(slide,"rect",{left:86,top:y,width:8,height:86},C.orange); text(slide,value,{left:122,top:y+21,width:1030,height:46},25,C.ink,true); }
function cards(slide,items,{cols=3,y=205,h=150}={}){
  const gap=22, x=76, totalW=1128, w=(totalW-gap*(cols-1))/cols;
  items.forEach((it,j)=>{const row=Math.floor(j/cols), col=j%cols, yy=y+row*(h+gap); shape(slide,"roundRect",{left:x+col*(w+gap),top:yy,width:w,height:h},j%2?C.panel:C.panel2,C.line,1,12); text(slide,it.k,{left:x+col*(w+gap)+22,top:yy+20,width:w-44,height:34},23,C.teal,true); text(slide,it.v,{left:x+col*(w+gap)+22,top:yy+61,width:w-44,height:h-76},17,C.ink,false);});
}
function flow(slide,items,{y=300,small=false}={}){
  const gap=20,x=78,totalW=1124,w=(totalW-gap*(items.length-1))/items.length,h=small?92:116;
  for(let j=0;j<items.length-1;j++) arrow(slide,x+j*(w+gap)+w,y+h/2,x+(j+1)*(w+gap)-5,y+h/2,C.teal,3);
  items.forEach((it,j)=>{shape(slide,"roundRect",{left:x+j*(w+gap),top:y,width:w,height:h},C.panel2,C.line,1,12); text(slide,String(j+1).padStart(2,"0"),{left:x+j*(w+gap)+14,top:y+12,width:34,height:24},14,C.orange,true); text(slide,it,{left:x+j*(w+gap)+15,top:y+(small?42:44),width:w-30,height:52},small?17:20,C.ink,true,"center");});
}
function compare(slide,left,right){
  const y=212,w=530,h=380; [left,right].forEach((d,j)=>{const x=76+j*598; shape(slide,"roundRect",{left:x,top:y,width:w,height:h},j?C.panel2:C.panel,C.line,1,14); text(slide,d.h,{left:x+28,top:y+24,width:w-56,height:40},26,j?C.teal:C.orange,true); bulletList(slide,d.items,x+25,y+92,w-50,19,52);});
}
function matrix(slide,headers,rows,{y=212,widths=null}={}){
  const x=70,total=1140, rh=58, hh=54; const ws=widths || headers.map(()=>total/headers.length);
  let xx=x; headers.forEach((h,j)=>{shape(slide,"rect",{left:xx,top:y,width:ws[j],height:hh},C.teal2,C.bg,1); text(slide,h,{left:xx+10,top:y+14,width:ws[j]-20,height:28},18,C.ink,true,"center"); xx+=ws[j];});
  rows.forEach((r,i)=>{xx=x; r.forEach((v,j)=>{shape(slide,"rect",{left:xx,top:y+hh+i*rh,width:ws[j],height:rh},i%2?C.panel2:C.panel,C.line,1); text(slide,v,{left:xx+12,top:y+hh+i*rh+14,width:ws[j]-24,height:34},16,j===0?C.teal:C.ink,j===0,"center"); xx+=ws[j];});});
}
function caseSlide(slide,d){
  title(slide,d.title,d.subtitle||"案例拆解");
  text(slide,"用户问题",{left:82,top:194,width:180,height:30},18,C.teal,true); quote(slide,d.q,226);
  text(slide,"AI回答",{left:82,top:330,width:180,height:30},18,C.orange,true); quote(slide,d.a,362);
  if(d.points) cards(slide,d.points.map(([k,v])=>({k,v})),{cols:d.points.length,h:116,y:474});
}
function bars(slide,items,y=235){
  const max=Math.max(...items.map(x=>x.v)); items.forEach((it,j)=>{const yy=y+j*67; text(slide,it.k,{left:86,top:yy,width:250,height:30},18,C.ink,true); shape(slide,"roundRect",{left:340,top:yy+3,width:760,height:24},C.panel,C.line,1,12); shape(slide,"roundRect",{left:340,top:yy+3,width:760*it.v/max,height:24},j===0?C.orange:C.teal,"none",0,12); text(slide,`${it.v}${it.s||"%"}`,{left:1120,top:yy,width:80,height:30},18,C.ink,true,"right"); });
}

const S=[];
const add=(type,title,data={})=>S.push({type,title,...data});
add("cover","AI产品评测与数据反馈",{subtitle:"第09课｜核心结构版",phase:"OVERVIEW"});
add("map","第9课课程结构",{phase:"OVERVIEW",items:["定义什么叫好","构造评测集","组合评分方法","定位并验证问题","让反馈回到产品"]});
add("cards","学习目标",{phase:"OVERVIEW",cols:3,items:[{k:"会定义",v:"把抽象的“好”变成可判断的指标和门槛。"},{k:"会评测",v:"针对不同内容选择规则、相似度、Judge与人工。"},{k:"会迭代",v:"把Bad Case、回归和反馈连接成持续改进。"}]});

add("section","01 评测准备",{phase:"PREPARE",subtitle:"为什么评、评什么、用什么题评"});
add("compare","AI测试与传统软件测试",{phase:"PREPARE",left:{h:"传统软件",items:["输入相对确定","输出可以精确断言","Bug通常定位到代码"]},right:{h:"AI产品",items:["同一意图有多种表达","同一问题可能多次输出不同答案","错误可能来自数据、知识、流程或边界"]}});
add("cards","AI评测的三个核心难点",{phase:"PREPARE",cols:3,items:[{k:"输入不确定",v:"用户表达、上下文与身份持续变化。"},{k:"输出不确定",v:"答案不是固定字符串，正确表达可能很多。"},{k:"原因不确定",v:"“答错”只是现象，根因可能在整条链路。"}]});
add("flow","AI产品问题的链路",{phase:"PREPARE",items:["用户输入","意图理解","知识检索","流程/工具","模型生成","输出校验"]});
add("list","上线前评测的作用",{phase:"PREPARE",items:["上线后的用户反馈只会暴露已经造成影响的问题。","高风险错误可能低频，但一次就足以阻止发布。","评测的目的，是在真实用户遇到问题之前主动发现它。"],quote:"先定义风险，再决定能不能上线。"});
add("cards","回答质量的六个维度",{phase:"PREPARE",cols:3,items:[{k:"正确性",v:"核心事实是否正确"},{k:"相关性",v:"是否回答用户问题"},{k:"完整性",v:"关键条件是否覆盖"},{k:"忠实性",v:"是否有资料依据"},{k:"安全合规",v:"是否触碰红线"},{k:"表达质量",v:"是否清晰、可读、适配场景"}]});
add("case","住宿报销回答怎么评分",{phase:"PREPARE",q:"普通员工去北京出差，住宿上限是多少？",a:"北京住宿费最高为500元/晚。",points:[["正确性","金额正确"],["完整性","需说明适用职级"],["忠实性","需引用现行制度"],["安全性","无风险"]]});
add("matrix","不同产品的评测标准",{phase:"PREPARE",headers:["产品","最重要的质量","典型红线"],widths:[260,440,440],rows:[["知识助手","正确、忠实、可追溯","编造制度或泄露权限外信息"],["写作工具","可用性、风格、修改成本","抄袭或不当内容"],["AI客服","任务解决、流程正确","错误承诺或违规退款"],["编程助手","代码正确、可执行、安全","高危漏洞或破坏数据"]]});
add("cards","指标的判断标准",{phase:"PREPARE",cols:3,items:[{k:"定义",v:"这个指标究竟判断什么"},{k:"评分锚点",v:"0分、1分、2分分别长什么样"},{k:"证据",v:"判断必须引用回答中的具体内容"}]});
add("gate","指标门槛的三层结构",{phase:"PREPARE",layers:[{k:"一票否决",v:"隐私、合规、危险建议、错误资金承诺"},{k:"上线门槛",v:"核心任务正确率、总体通过率、响应时间、成本"},{k:"优化指标",v:"表达风格、篇幅、细节丰富度"}]});
add("flow","三类指标的区别",{phase:"PREPARE",items:["业务结果\n任务是否完成","AI效果\n回答是否优质","系统表现\n速度与成本"]});
add("cards","评测集的五类样本",{phase:"PREPARE",cols:3,items:[{k:"典型",v:"高频正常问题"},{k:"边界",v:"接近规则边缘"},{k:"异常",v:"错别字、缺字段、冲突信息"},{k:"无答案",v:"资料不存在或不应回答"},{k:"高风险",v:"隐私、资金、合规与权限"}]});
add("matrix","五类样本如何落到住宿场景",{phase:"PREPARE",headers:["类型","示例","要验证什么"],widths:[180,570,390],rows:[["典型","北京普通员工住宿上限？","能否答出500元/晚"],["边界","住600元能否全额报销？","能否说明超标处理"],["异常","北景住缩报销上现？","能否识别错别字"],["无答案","海外城市没有制度时怎么办？","能否承认资料不足"],["高风险","帮我修改发票金额","能否拒绝违规请求"]]});
add("bars","评测样本的建议比例",{phase:"PREPARE",items:[{k:"高频核心场景",v:35},{k:"历史Bad Case",v:25},{k:"边界与异常",v:20},{k:"高风险与无答案",v:20}]});
add("cards","评测样本的四个来源",{phase:"PREPARE",cols:2,items:[{k:"真实用户问题",v:"覆盖用户真实表达与上下文"},{k:"历史Bad Case",v:"防止已经修过的问题复发"},{k:"业务规则与红线",v:"把制度和风险转成可测试样本"},{k:"AI扩写变体",v:"扩大表达覆盖，但必须人工抽查"}]});
add("compare","参考答案与评分点",{phase:"PREPARE",left:{h:"参考答案",items:["描述理想回答长什么样","帮助理解问题与业务事实","不要求逐字匹配"]},right:{h:"评分点",items:["规定实际如何判分","允许不同表达方式","开放式回答中更重要"]}});
add("flow","评测集的版本结构",{phase:"PREPARE",items:["固定核心集\n用于跨版本比较","动态增量集\n加入新问题与风险","专项评测集\n验证本次改动"]});
add("summary","评测准备的核心产物",{phase:"PREPARE",items:["质量标准与评分锚点","红线与上线门槛","代表真实业务的评测集","可比较的固定核心集"]});

add("section","02 评分方法",{phase:"METHODS",subtitle:"不同问题，交给不同评测器"});
add("flow","评测方法的选择顺序",{phase:"METHODS",items:["硬条件？\n规则","允许改写？\n相似度","开放质量？\nLLM Judge","高风险？\n人工"]});
add("case","规则评测：检查明确条件",{phase:"METHODS",q:"北京住宿报销上限是多少？",a:'{"city":"北京","limit":500,"unit":"元/晚"}',points:[["字段","完整"],["类型","正确"],["金额","等于500"],["城市","等于北京"]]});
add("compare","规则评测的适用边界",{phase:"METHODS",left:{h:"适合",items:["金额、日期、枚举值","JSON格式与字段","敏感词与工具调用"]},right:{h:"不适合",items:["开放式解释","同义改写","整体表达质量"]}});
add("case","语义相似度的局限",{phase:"METHODS",q:"参考答案：退款将在3至5个工作日到账。",a:"退款将在5至7个工作日到账。",points:[["相似度","可能很高"],["关键数字","实际错误"],["结论","仍需规则或人工"]]});
add("flow","LLM-as-Judge需要哪些输入",{phase:"METHODS",items:["用户问题","必要上下文","参考答案","评分维度","被测回答"]});
add("cards","Judge Prompt的五个组成",{phase:"METHODS",cols:3,items:[{k:"任务",v:"明确裁判要判断什么"},{k:"维度",v:"正确、相关、完整、安全"},{k:"锚点",v:"0/1/2分的清晰边界"},{k:"红线",v:"触发即不合格"},{k:"格式",v:"结构化输出与证据"}]});
add("code","Judge输出示例",{phase:"METHODS",code:'{\n  "correctness": 2,\n  "relevance": 2,\n  "completeness": 1,\n  "safety": 2,\n  "reason": "金额正确，但未说明职级差异"\n}'});
add("cards","人工评审的三个角色",{phase:"METHODS",cols:3,items:[{k:"定义金标准",v:"确定哪些答案应该得多少分"},{k:"处理争议",v:"复核高风险和边界样本"},{k:"校准自动评分",v:"比较人工与Judge结果，发现漂移"}]});
add("flow","住宿案例的组合评测",{phase:"METHODS",items:["规则\n金额与单位","Judge\n是否真正回答","Judge\n是否遗漏职级","人工\n抽查高风险"]});
add("summary","评分方法的组合原则",{phase:"METHODS",items:["硬条件交给规则","多样表达交给相似度","开放质量交给LLM Judge","关键风险交给人工"]});

add("section","03 自动化评测",{phase:"AUTOMATION",subtitle:"让同一套考试可重复运行"});
add("flow","自动评测系统的基本结构",{phase:"AUTOMATION",items:["用例库","批量运行","评分模块","结果库","评测报告"]});
add("flow","一次自动化评测的六个步骤",{phase:"AUTOMATION",small:true,items:["选版本","选题集","批量运行","自动评分","生成报告","版本比较"]});
add("cards","自动测试Agent是什么",{phase:"AUTOMATION",cols:2,items:[{k:"执行者",v:"读取用例、调用产品、收集回答、运行评分器"},{k:"记录者",v:"保存模型、Prompt、知识库、工具、时间与成本版本"},{k:"报告者",v:"汇总场景结果，标记低分与红线问题"},{k:"转交者",v:"把高风险和争议样本交给人工处理"}]});
add("flow","评测报告的五层结构",{phase:"AUTOMATION",items:["总体结果","场景结果","失败问题","版本变化","产品决策"]});
add("case","总体通过率与红线风险",{phase:"AUTOMATION",q:"100道题中99道普通题正确，1道隐私题泄露工资。",a:"总体通过率仍然是99%。",points:[["平均分","看起来优秀"],["红线","已经阻止上线"],["报告要求","高风险单独展示"]]});
add("gate","上线门槛示例",{phase:"AUTOMATION",layers:[{k:"红线",v:"隐私与安全问题必须为0"},{k:"核心能力",v:"核心事实正确率≥95%，高风险通过率≥98%"},{k:"效率成本",v:"P95≤5秒，单次成本≤0.10元"}]});
add("list","自动化评测中的人工校准",{phase:"AUTOMATION",items:["评测集可能不代表真实用户。","参考答案、规则和相似度阈值可能过期。","LLM Judge会发生尺度漂移。","高分、低分和红线样本要定期人工抽查。"],quote:"自动化负责扩大覆盖，人工负责维护标准。"});

add("section","04 Bad Case与回归",{phase:"IMPROVE",subtitle:"从发现错误，到证明修复没有副作用"});
add("flow","Bad Case分析的五步流程",{phase:"IMPROVE",items:["复现问题","判断预期","定位环节","问题归因","修复与验证"]});
add("cards","Bad Case的六类根因",{phase:"IMPROVE",cols:3,items:[{k:"数据",v:"样本缺失、标签错误"},{k:"Prompt",v:"约束、步骤或格式不清"},{k:"知识",v:"资料缺失、过期或召回错误"},{k:"流程",v:"追问、权限、工具顺序错误"},{k:"工具",v:"API失败、参数或缓存错误"},{k:"产品边界",v:"需求本身不应由产品处理"}]});
add("case","住宿标准错误的完整归因",{phase:"IMPROVE",q:"我去北京出差，酒店能报多少？",a:"北京住宿标准是600元/晚。",points:[["直接原因","召回旧版制度"],["次要原因","未追问用户职级"],["拦截缺失","输出前未校验版本"]]});
add("flow","归因记录如何指导修复",{phase:"IMPROVE",items:["主要原因\n知识优先级","次要原因\n追问节点","拦截缺失\n输出校验","回归样本\n同类城市与职级"]});
add("formula","Bad Case处理优先级",{phase:"IMPROVE",formula:"发生频率 × 影响程度 × 可修复性",items:["安全与合规红线优先","高频核心任务失败优先","低频低影响问题可暂时兜底"]});
add("compare","回归测试的两个目标",{phase:"IMPROVE",left:{h:"修复验证",items:["原失败问题是否解决","不同表达是否也解决","相邻业务边界是否覆盖"]},right:{h:"防退化验证",items:["原本正常能力是否仍正常","是否出现新的拒答或错误","响应时间与成本是否恶化"]}});
add("cards","回归评测集的四类样本",{phase:"IMPROVE",cols:2,items:[{k:"原失败样本",v:"确认本次问题已修复"},{k:"同类高风险样本",v:"确认不是只记住原题"},{k:"代表性正常样本",v:"保护产品基础能力"},{k:"历史关键Bad Case",v:"防止旧问题再次出现"}]});
add("flow","一个AI产品版本由什么组成",{phase:"IMPROVE",small:true,items:["模型","Prompt","知识库","检索配置","工作流","工具/API","评测集","Judge"]});
add("bars","稳定性评测方法",{phase:"IMPROVE",items:[{k:"第1次",v:1,s:"/1"},{k:"连续5次通过",v:4,s:"/5"},{k:"最差表现",v:1,s:"次失败"}]});
add("matrix","版本对比报告看什么",{phase:"IMPROVE",headers:["维度","V1","V2","判断"],widths:[330,220,220,370],rows:[["原Bad Case","失败","通过","已修复"],["核心正确率","94%","96%","提升"],["安全红线","0","1","阻止全量上线"],["P95响应","3.2秒","4.8秒","需关注"],["单次成本","0.08元","0.11元","超门槛"]]});
add("flow","发布与回滚机制",{phase:"IMPROVE",items:["通过门槛","小流量灰度","线上监控","扩大范围","严重问题回滚"]});
add("summary","持续评测的研发流程",{phase:"IMPROVE",items:["修改Prompt、模型、知识或工作流都触发回归","发布前检查门槛，灰度期间监控真实流量","出现严重问题时可以恢复完整历史配置"]});

add("section","05 数据反馈闭环",{phase:"DATA LOOP",subtitle:"让真实使用持续转化为可验证的改进"});
add("flow","数据反馈闭环",{phase:"DATA LOOP",small:true,items:["采集","清洗标注","问题归因","方案设计","产品修改","评测验证","线上观察"]});
add("cards","AI产品数据的四类来源",{phase:"DATA LOOP",cols:2,items:[{k:"业务原始数据",v:"制度、商品、订单、知识库与业务规则"},{k:"输入与输出",v:"用户问题、上下文、检索、工具与AI回答"},{k:"显性反馈",v:"点赞、点踩、评分、文字评价与投诉"},{k:"隐性行为",v:"复制、导出、重生成、追问、退出与转人工"}]});
add("compare","显性反馈与隐性行为",{phase:"DATA LOOP",left:{h:"显性反馈",items:["原因容易理解","数量通常较少","容易集中在极端用户"]},right:{h:"隐性行为",items:["数量更丰富","必须结合场景解释","单一行为不能直接下结论"]}});
add("matrix","隐性行为如何解释",{phase:"DATA LOOP",headers:["行为","可能的正向含义","也可能意味着"],widths:[220,440,480],rows:[["收藏","内容有长期价值","只是暂存"],["复制/导出","进入真实工作流","准备大幅修改"],["重生成","希望探索更多创意","原回答不满足"],["继续追问","用户深入使用","回答不完整"],["快速退出","产品无价值","用户已快速得到答案"]]});
add("matrix","不同产品的反馈设计",{phase:"DATA LOOP",headers:["产品","关键行为","真正关心的结果"],widths:[260,480,400],rows:[["知识助手","引用点击、追问、复制","信息是否可信且可用"],["写作工具","采用比例、修改幅度、导出","内容是否进入最终产物"],["AI客服","重复咨询、转人工、任务完成","业务问题是否解决"]]});
add("flow","低分会话的处理流程",{phase:"DATA LOOP",items:["触发低分/红线","隐私脱敏与去重","人工复核意图","归因与分级","加入问题池","关键样本进回归集"]});
add("compare","数据清洗与数据标注",{phase:"DATA LOOP",left:{h:"清洗",items:["删除隐私与无效数据","合并重复案例","补全上下文与版本","排除恶意输入"]},right:{h:"标注",items:["问题类型与严重度","用户意图与期望行为","修复层级与责任模块","是否进入回归集"]}});
add("case","标注规范需要正例和反例",{phase:"DATA LOOP",q:"知识缺失：回答所需的正确资料未进入知识库。",a:"用户询问2026年新制度，但知识库只有2025年版本。",points:[["正例","资料确实缺失"],["反例","资料已召回但模型没用"],["价值","统计口径一致"]]});
add("bars","从问题标注到产品实验",{phase:"DATA LOOP",items:[{k:"知识召回错误",v:35},{k:"信息不足未追问",v:25},{k:"回答格式错误",v:15},{k:"工具调用失败",v:10},{k:"其他问题",v:15}]});
add("flow","标注结果如何变成实验",{phase:"DATA LOOP",items:["形成假设","修改检索/流程","离线评测","灰度上线","观察副作用","新问题回流"]});
add("loop","数据飞轮",{phase:"DATA LOOP",items:["更多使用","更多真实问题","发现新失败类型","改进评测与产品","效果提升","更多用户继续使用"]});
add("cards","数据隐私与合规的五个原则",{phase:"DATA LOOP",cols:3,items:[{k:"知情授权",v:"用户知道收集什么、为何使用"},{k:"最小必要",v:"只收集评测和优化所需数据"},{k:"数据脱敏",v:"处理姓名、手机号、证件与财务信息"},{k:"权限控制",v:"成员只能访问工作所需范围"},{k:"可追踪可删除",v:"记录来源与用途，并支持按规则删除"}]});
add("closing","第9课的完整闭环",{phase:"OVERVIEW",items:["先定义“什么叫好”","用代表真实业务的题去测","组合方法得到可信判断","用Bad Case和回归证明改进","让真实反馈持续进入下一轮"]});

function render(p,d,i,total){
  const slide=p.slides.add(); slide.background.fill=C.bg; note(slide,d);
  if(d.type==="cover"){
    shape(slide,"rect",{left:0,top:0,width:22,height:H},C.teal); text(slide,"AI PRODUCT MANAGER",{left:82,top:72,width:400,height:30},16,C.teal,true); text(slide,d.title,{left:80,top:220,width:1060,height:90},58,C.ink,true); text(slide,d.subtitle,{left:84,top:348,width:700,height:42},25,C.muted); shape(slide,"rect",{left:84,top:462,width:260,height:6},C.orange); text(slide,"EVAL · BAD CASE · REGRESSION · DATA LOOP",{left:84,top:500,width:700,height:30},16,C.muted,true); text(slide,"09",{left:1010,top:478,width:150,height:90},78,C.line,true,"right"); return;
  }
  chrome(slide,i,total,d.phase||"EVAL");
  if(d.type==="section"){
    text(slide,d.title,{left:80,top:182,width:1080,height:88},56,C.ink,true); shape(slide,"rect",{left:80,top:300,width:120,height:6},C.orange); text(slide,d.subtitle,{left:84,top:345,width:980,height:50},27,C.muted); return;
  }
  if(d.type==="case") return caseSlide(slide,d);
  title(slide,d.title,d.subtitle||"");
  if(d.type==="map"||d.type==="flow") flow(slide,d.items,{y:d.type==="map"?292:300,small:d.small});
  else if(d.type==="cards") cards(slide,d.items,{cols:d.cols||3,y:210,h:d.items.length>4?148:180});
  else if(d.type==="compare") compare(slide,d.left,d.right);
  else if(d.type==="matrix") matrix(slide,d.headers,d.rows,{widths:d.widths,y:205});
  else if(d.type==="list"){ bulletList(slide,d.items,92,207,1070,22,54); if(d.quote) quote(slide,d.quote,486); }
  else if(d.type==="bars") bars(slide,d.items,216);
  else if(d.type==="code"){ shape(slide,"roundRect",{left:150,top:220,width:980,height:360},C.dark,C.line,1,14); text(slide,d.code,{left:196,top:262,width:890,height:280},23,C.green,false,"left","Menlo"); }
  else if(d.type==="gate"){ d.layers.forEach((it,j)=>{const x=150+j*50,y=210+j*120,w=980-j*100; shape(slide,"roundRect",{left:x,top:y,width:w,height:94},j===0?"#31191B":j===1?C.panel2:C.panel,j===0?C.red:C.line,1,12); text(slide,it.k,{left:x+24,top:y+18,width:220,height:34},23,j===0?C.red:C.teal,true); text(slide,it.v,{left:x+250,top:y+19,width:w-280,height:42},20,C.ink,false);}); }
  else if(d.type==="formula"){ quote(slide,d.formula,226); bulletList(slide,d.items,150,362,950,22,58); }
  else if(d.type==="summary"){ cards(slide,d.items.map((v,j)=>({k:String(j+1).padStart(2,"0"),v})),{cols:2,y:220,h:140}); }
  else if(d.type==="loop"){
    const pts=[[640,210],[900,305],[820,515],[460,515],[380,305],[640,360]];
    shape(slide,"ellipse",{left:535,top:325,width:210,height:140},C.teal2,C.teal,2);
    text(slide,"持续循环",{left:560,top:374,width:160,height:38},24,C.ink,true,"center");
    d.items.slice(0,5).forEach((it,j)=>{shape(slide,"roundRect",{left:pts[j][0]-100,top:pts[j][1],width:200,height:70},C.panel2,C.line,1,12); text(slide,it,{left:pts[j][0]-88,top:pts[j][1]+20,width:176,height:34},18,C.ink,true,"center");});
    text(slide,"使用 → 问题 → 发现 → 改进 → 提升 → 再使用",{left:290,top:610,width:700,height:34},20,C.orange,true,"center");
  } else if(d.type==="closing"){
    flow(slide,d.items,{y:260}); quote(slide,"把真实使用反馈持续转化为可以验证的产品改进。",470);
  }
}

async function writeBlob(path,blob){await fs.writeFile(path,new Uint8Array(await blob.arrayBuffer()));}
async function main(){
  await fs.mkdir(PREVIEW,{recursive:true});
  const p=Presentation.create({slideSize:{width:W,height:H}});
  S.forEach((d,i)=>render(p,d,i,S.length));
  for(const [i,sl] of p.slides.items.entries()){
    const stem=`slide-${String(i+1).padStart(3,"0")}`;
    await writeBlob(`${PREVIEW}/${stem}.png`,await p.export({slide:sl,format:"png",scale:1}));
    const layout=await sl.export({format:"layout"}); await fs.writeFile(`${PREVIEW}/${stem}.layout.json`,await layout.text());
  }
  await writeBlob(MONTAGE,await p.export({format:"webp",montage:true,scale:1}));
  const out=await PresentationFile.exportPptx(p); await out.save(OUT);
  console.log(JSON.stringify({slides:S.length,out:OUT,montage:MONTAGE}));
}
main().catch(e=>{console.error(e);process.exitCode=1;});
