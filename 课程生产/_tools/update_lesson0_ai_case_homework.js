const pptxgen = require('pptxgenjs');
const fs = require('fs');
const JSZip = require('jszip');

const root = '/Users/didi/Project/AI课程/课程生产/第一阶段/第00课_课程说明会学习地图与第一次AI场景思考_完整课程包_v2';
const target = `${root}/06_正式PPT.pptx`;
const generated = '/tmp/lesson0_new_content.pptx';

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };
pptx.defineSlideMaster({
  title: 'MASTER', background: { color: 'F6F8FC' },
  objects: [
    { rect: { x:0,y:0,w:13.333,h:0.09,fill:{color:'36C5F0'},line:{color:'36C5F0'} } },
    { text: { text:'AI 产品经理系统课程 · 第一阶段', options:{x:0.55,y:7.08,w:4.5,h:0.18,fontFace:'PingFang SC',fontSize:8.5,color:'718096',margin:0} } }
  ],
  slideNumber: { x:12.3,y:7.02,w:0.45,h:0.2,color:'718096',fontFace:'Aptos',fontSize:9,align:'right',margin:0 }
});

const C={navy:'142B4A',blue:'246BFD',cyan:'36C5F0',orange:'FF9F43',green:'27AE60',red:'E85D75',ink:'1F2937',gray:'667085',white:'FFFFFF',bg:'F6F8FC',line:'D9E2EF',purple:'7655D4',light:'EAF0F8'};
const S=pptx.ShapeType;
function title(s,t,k=''){
  if(k)s.addText(k,{x:0.62,y:0.34,w:4.4,h:0.22,fontSize:9.5,bold:true,color:C.blue,charSpacing:1.1,margin:0});
  s.addText(t,{x:0.62,y:0.62,w:12,h:0.52,fontSize:24,bold:true,color:C.navy,margin:0,fit:'shrink'});
  s.addShape(S.line,{x:0.62,y:1.28,w:12.05,h:0,line:{color:C.line,width:1}});
}
function card(s,x,y,w,h,head,body,color=C.blue,num=''){
  s.addShape(S.roundRect,{x,y,w,h,rectRadius:0.06,fill:{color:C.white},line:{color:C.line,width:1},shadow:{type:'outer',color:'AAB6C5',blur:1,angle:45,distance:1,opacity:0.14}});
  s.addShape(S.rect,{x,y,w:0.08,h,fill:{color},line:{color}});
  if(num){s.addShape(S.ellipse,{x:x+0.25,y:y+0.24,w:0.43,h:0.43,fill:{color},line:{color}});s.addText(String(num),{x:x+0.25,y:y+0.33,w:0.43,h:0.14,fontSize:9.5,bold:true,color:C.white,align:'center',margin:0});}
  s.addText(head,{x:x+(num?0.82:0.28),y:y+0.23,w:w-(num?1.05:0.55),h:0.34,fontSize:15,bold:true,color:C.navy,margin:0,fit:'shrink'});
  s.addText(body,{x:x+0.28,y:y+0.82,w:w-0.56,h:h-1.04,fontSize:11.2,color:C.gray,margin:0.02,breakLine:false,valign:'top',fit:'shrink',paraSpaceAfterPt:7});
}
function pill(s,text,x,y,color=C.blue,w=2){s.addShape(S.roundRect,{x,y,w,h:0.34,rectRadius:0.1,fill:{color,transparency:88},line:{color,transparency:100}});s.addText(text,{x:x+0.06,y:y+0.075,w:w-0.12,h:0.14,fontSize:9,bold:true,color,align:'center',margin:0,fit:'shrink'});}
function arrow(s,x,y,w=0.45,color=C.cyan){s.addShape(S.chevron,{x,y,w,h:0.42,fill:{color,transparency:8},line:{color,transparency:100}});}
function bottom(s,text,color=C.navy){s.addShape(S.roundRect,{x:1.0,y:5.88,w:11.3,h:0.62,rectRadius:0.05,fill:{color:'EEF3FA'},line:{color:'D9E2EF'}});s.addText(text,{x:1.25,y:6.08,w:10.8,h:0.22,fontSize:13,bold:true,color,align:'center',margin:0,fit:'shrink'});}
function slide(){return pptx.addSlide('MASTER');}

// 04 AI时代背景：8页
{
 const s=slide(); title(s,'AI已经发展70多年，为什么最近几年突然爆发？','AI时代背景');
 const a=[['1950s','人工智能概念诞生'],['1990—2010','机器学习进入搜索、推荐和风控'],['2012','深度学习推动图像与语音识别'],['2017','Transformer成为大模型重要基础'],['2022','生成式AI进入大众市场'],['2024至今','推理、工具调用与Agent']];
 a.forEach((d,i)=>card(s,0.7+(i%3)*4.18,1.55+Math.floor(i/3)*2.1,3.78,1.62,d[0],d[1],[C.blue,C.purple,C.orange,C.cyan,C.green,C.red][i],i+1));
 bottom(s,'最近几年爆发的不是“AI这个概念”，而是普通人可以直接使用的生成式AI。');
}
{
 const s=slide(); title(s,'AI、机器学习、NLP、AIGC，不是同一层概念','概念地图');
 card(s,0.78,1.58,3.7,3.95,'AI要达到什么目标','人工智能 AI\n\n让机器具备感知、理解、推理、决策和行动等能力。',C.blue,'1');
 card(s,4.82,1.58,3.7,3.95,'AI通过什么实现','规则系统\n机器学习\n深度学习\n大模型',C.purple,'2');
 card(s,8.87,1.58,3.7,3.95,'AI正在解决什么','自然语言处理\n计算机视觉\n语音识别\n推荐、预测与内容生成',C.orange,'3');
 bottom(s,'AI是目标，机器学习是方法，NLP是问题领域，AIGC是能力与应用形态。');
}
{
 const s=slide(); title(s,'AI是一个大范围，不等于聊天机器人','什么是AI');
 const a=[['推荐','判断用户可能喜欢什么'],['预测','地图预测到达时间'],['识别','风控判断交易是否异常'],['理解','客服识别用户的真实问题'],['生成','大模型生成文字、图片和代码']];
 a.forEach((d,i)=>card(s,0.7+(i%3)*4.18,1.55+Math.floor(i/3)*2.2,3.78,1.7,d[0],d[1],[C.blue,C.purple,C.orange,C.cyan,C.green][i],i+1));
 bottom(s,'判断、预测和识别是AI；生成新的内容是AIGC；AIGC不等于全部AI。');
}
{
 const s=slide(); title(s,'AI如何从“人写规则”走向“机器学习规律”','机器学习与深度学习');
 card(s,0.78,1.55,5.65,3.85,'传统软件','人编写规则\n↓\n机器执行规则\n↓\n得到结果\n\n例：出现“非常差”就判断为负面评论。',C.orange,'1');
 card(s,6.88,1.55,5.65,3.85,'机器学习','提供大量案例\n↓\n机器学习数据规律\n↓\n预测新的结果\n\n例：从大量已标注评论中学习负面表达。',C.blue,'2');
 bottom(s,'人工智能 AI  ⊃  机器学习 ML  ⊃  深度学习 DL  ⊃  大模型与生成模型');
}
{
 const s=slide(); title(s,'NLP、视觉和语音：AI正在处理哪一类信息','AI能力领域');
 const a=[['NLP 自然语言处理','分类、摘要、翻译、问答'],['CV 计算机视觉','图片识别、检测、审核'],['ASR 语音识别','将人的语音转成文字'],['TTS 语音合成','将文字信息转成语音'],['推荐与预测','推荐、评分、风险预测']];
 a.forEach((d,i)=>card(s,0.68+(i%3)*4.2,1.52+Math.floor(i/3)*2.2,3.82,1.72,d[0],d[1],[C.blue,C.purple,C.orange,C.cyan,C.green][i],i+1));
 bottom(s,'销售拜访记录助手会同时使用ASR转写录音、NLP理解对话、AIGC生成记录。');
}
{
 const s=slide(); title(s,'AIGC：AI从“识别内容”走向“生成内容”','生成式AI');
 const rows=[['用户评论','判断正面或负面','识别 / 分类'],['用户评论','生成评论总结','AIGC'],['商品图片','判断是否违规','识别 / 审核'],['商品描述','生成宣传图片','AIGC']];
 const xs=[0.9,4.1,9.4], ws=[3.0,5.1,2.7]; ['输入','AI输出','类型'].forEach((h,i)=>{s.addShape(S.rect,{x:xs[i],y:1.55,w:ws[i],h:0.62,fill:{color:C.navy},line:{color:C.navy}});s.addText(h,{x:xs[i]+0.1,y:1.75,w:ws[i]-0.2,h:0.2,fontSize:12,bold:true,color:C.white,align:'center',margin:0});});
 rows.forEach((r,ri)=>r.forEach((v,i)=>{s.addShape(S.rect,{x:xs[i],y:2.17+ri*0.75,w:ws[i],h:0.75,fill:{color:ri%2?'FFFFFF':'F0F4FA'},line:{color:C.line}});s.addText(v,{x:xs[i]+0.12,y:2.42+ri*0.75,w:ws[i]-0.24,h:0.22,fontSize:11.5,bold:i===2,color:i===2&&v==='AIGC'?C.green:C.ink,align:'center',margin:0,fit:'shrink'});}));
 bottom(s,'一个AI产品可以同时使用识别、预测和生成等多种能力。');
}
{
 const s=slide(); title(s,'模型不等于产品','大模型、LLM与ChatGPT');
 const a=[['基础模型','提供语言理解、生成和推理能力'],['产品系统','加入知识库、工具、流程、权限和评测'],['产品功能','问答、总结、分析、生成和执行'],['用户价值','帮助用户更快、更好地完成任务']];
 a.forEach((d,i)=>{card(s,0.65+i*3.16,1.65,2.82,3.8,d[0],d[1],[C.blue,C.purple,C.orange,C.green][i],i+1);if(i<3)arrow(s,3.49+i*3.16,3.25,0.42)});
 bottom(s,'大模型像发动机，但只有发动机还不能成为一辆可以上路的汽车。');
}
{
 const s=slide(); title(s,'AI正在从“生成内容”走向“执行任务”','最近几年的加速');
 const a=[['2017','Transformer','大规模训练与语言处理'],['2018—21','预训练模型','一个基础模型适应多个任务'],['2022','生成式AI普及','普通用户用自然语言使用AI'],['2023','大模型与多模态','同时处理文字、图片等信息'],['2024','长上下文与工具调用','读取更多资料并连接外部工具'],['2025至今','推理与Agent','围绕目标连续完成多个步骤']];
 a.forEach((d,i)=>card(s,0.68+(i%3)*4.2,1.5+Math.floor(i/3)*2.18,3.82,1.72,d[0]+'｜'+d[1],d[2],[C.blue,C.purple,C.orange,C.cyan,C.green,C.red][i],i+1));
 bottom(s,'识别信息  →  生成内容  →  分析问题  →  调用工具  →  执行任务');
}
{
 const s=slide(); title(s,'AI不只是给原来的产品增加一个聊天框','AI产品的变化');
 card(s,0.78,1.52,5.55,3.75,'传统产品','用户点击功能\n↓\n系统执行固定规则\n↓\n返回相对确定的结果',C.orange,'1');
 card(s,6.98,1.52,5.55,3.75,'AI产品','用户表达目标\n↓\nAI理解任务并获取知识\n↓\n生成或执行，由用户确认结果',C.blue,'2');
 bottom(s,'AI产品设计的核心，是重新设计用户、模型与业务流程之间的分工。');
}

// 05 案例精讲：7页
{
 const s=slide(); title(s,'Case：销售拜访记录助手','课堂案例'); pill(s,'先独立思考2分钟',10.2,0.7,C.orange,2.25);
 card(s,0.78,1.52,5.5,3.92,'拜访后需要记录','客户核心诉求\n当前使用的解决方案\n主要异议与合作意向\n下一步跟进计划\n负责人与跟进时间',C.blue,'1');
 card(s,6.82,1.52,5.7,3.92,'现实中经常发生','拜访结束后没有时间记录\n依赖销售个人回忆\nCRM内容过于简单\n关键问题和待办事项被遗漏',C.red,'2');
 bottom(s,'如果让你负责这个场景，你会怎么设计AI产品？');
}
{
 const s=slide(); title(s,'不要从“AI能做什么”出发','第一步｜确认问题');
 card(s,0.78,1.52,5.55,3.95,'先问五个问题','谁遇到了问题？\n问题发生在什么环节？\n用户现在怎么解决？\n问题发生频率有多高？\n这个问题造成了什么损失？',C.orange,'?');
 card(s,6.92,1.52,5.6,3.95,'本案例的初步判断','用户：经常拜访客户的销售和BD\n场景：拜访或电话后录入CRM\n现状：手写、录音、聊天记录与回忆\n问题：整理耗时，关键信息容易遗漏',C.blue,'1');
 bottom(s,'AI产品机会不是“可以使用AI的地方”，而是“值得被解决的用户问题”。');
}
{
 const s=slide(); title(s,'AI应该承担哪一个具体步骤？','第二步｜拆解流程');
 const a=['拜访前准备','进行客户沟通','记录沟通内容','整理关键信息','填写CRM','创建后续任务','持续跟进客户'];
 a.forEach((v,i)=>{const x=0.42+i*1.83;s.addShape(S.roundRect,{x,y:1.7,w:1.55,h:1.0,rectRadius:0.06,fill:{color:i>=2&&i<=5?'EAF2FF':'FFFFFF'},line:{color:i>=2&&i<=5?C.blue:C.line,width:i>=2&&i<=5?1.5:1}});s.addText(v,{x:x+0.12,y:2.0,w:1.31,h:0.36,fontSize:10.5,bold:i>=2&&i<=5,color:i>=2&&i<=5?C.blue:C.gray,align:'center',margin:0,fit:'shrink'});if(i<6)arrow(s,x+1.58,1.99,0.24,C.line)});
 card(s,0.78,3.2,5.65,2.35,'适合AI优先介入','录音转文字、提取诉求与异议、总结沟通结论、生成CRM记录草稿和待办事项。',C.blue,'AI');
 card(s,6.88,3.2,5.65,2.35,'仍然需要人判断','确认客户真实意愿、确认承诺和报价、决定销售策略、确认后正式写入CRM。',C.orange,'人');
}
{
 const s=slide(); title(s,'AI负责整理，人负责确认和决策','第三步｜设计人机分工');
 const a=[['完成拜访','销售'],['录音转文字','AI'],['提取诉求与异议','AI'],['生成结构化记录','AI'],['检查与修改','销售'],['写入CRM并创建任务','系统']];
 a.forEach((d,i)=>{const x=0.43+i*2.12;card(s,x,1.55,1.86,2.25,d[0],d[1],d[1]==='AI'?C.blue:(d[1]==='销售'?C.orange:C.green),i+1);if(i<5)arrow(s,x+1.88,2.45,0.25,C.line)});
 s.addShape(S.roundRect,{x:1.2,y:4.3,w:10.9,h:1.25,rectRadius:0.06,fill:{color:'EEF4FF'},line:{color:'C8D9FA'}});
 s.addText('示例输出：客户核心诉求｜主要异议｜合作意向｜下一步行动｜负责人｜截止时间',{x:1.55,y:4.74,w:10.2,h:0.34,fontSize:13.5,bold:true,color:C.navy,align:'center',margin:0,fit:'shrink'});
 bottom(s,'AI不是直接替销售做决定，而是先降低信息整理和录入成本。');
}
{
 const s=slide(); title(s,'从对话录音到可确认的结构化记录','AI输出示例');
 card(s,0.78,1.52,3.72,3.95,'客户核心诉求','希望减少门店活动配置的操作时间。\n\n主要异议：担心新系统的培训成本过高。',C.blue,'1');
 card(s,4.8,1.52,3.72,3.95,'合作意向','愿意参加小范围试用，但暂不愿意全量采购。\n\n信息状态：需要销售人工确认。',C.orange,'2');
 card(s,8.82,1.52,3.72,3.95,'下一步行动','1. 周五前发送试用方案\n负责人：销售A\n\n2. 下周安排产品演示\n负责人：产品B',C.green,'3');
 bottom(s,'产品不仅要展示结果，还应允许销售查看原文、修改内容并确认后写入。');
}
{
 const s=slide(); title(s,'能生成结果，不代表结果可以直接使用','第四步｜风险与兜底');
 const a=[['说话人混淆','把销售的建议当成客户承诺'],['术语识别错误','产品名、金额和日期被错记'],['信息遗漏','遗漏负责人、截止时间或异议'],['无中生有','生成对话中没有出现的信息'],['敏感信息泄露','客户信息进入不合规系统']];
 a.forEach((d,i)=>card(s,0.68+(i%3)*4.2,1.5+Math.floor(i/3)*2.18,3.82,1.72,d[0],d[1],[C.red,C.orange,C.purple,C.blue,C.green][i],i+1));
 bottom(s,'低风险结果允许编辑；关键事实展示原文；承诺、报价和敏感信息必须人工确认。');
}
{
 const s=slide(); title(s,'不只看模型准不准，还要看业务有没有变好','第五步｜验证价值');
 card(s,0.78,1.55,3.72,3.95,'效率指标','每次拜访后的平均记录时间\nCRM录入完成时间\n销售使用助手的频率',C.blue,'1');
 card(s,4.8,1.55,3.72,3.95,'质量指标','CRM关键字段完整率\n行动项遗漏率\n销售修改AI结果的比例\n关键信息提取准确率',C.orange,'2');
 card(s,8.82,1.55,3.72,3.95,'业务指标','客户跟进任务按时完成率\n有效客户信息沉淀率\n长期观察客户转化率',C.green,'3');
 bottom(s,'“生成一份摘要”是产品输出；“更快完成记录并减少遗漏”才是产品价值。');
}
{
 const s=slide(); title(s,'用五个问题完成一次AI场景分析','案例复盘');
 const a=[['用户是谁？','经常拜访客户的销售和BD'],['问题是什么？','整理耗时，关键信息容易遗漏'],['AI具体做什么？','转写、提取、总结、生成草稿与行动项'],['人需要做什么？','确认事实、修改结果、判断意向与决策'],['怎么验证价值？','记录时间、完整率、修改率、遗漏率与跟进完成率']];
 a.forEach((d,i)=>card(s,0.68+(i%3)*4.2,1.5+Math.floor(i/3)*2.18,3.82,1.72,d[0],d[1],[C.blue,C.purple,C.orange,C.cyan,C.green][i],i+1));
 bottom(s,'完整AI产品方案 = 用户问题 + AI任务 + 人机分工 + 错误兜底 + 验证指标');
}

// 06 课后作业：3页 + 结尾
{
 const s=slide(); title(s,'回到自己的工作场景，找到一个真实问题','课后作业');
 card(s,0.78,1.55,5.55,3.95,'不要从这里开始','“我们能不能做一个AI助手？”\n\n“最近哪个AI技术最热？”\n\n“我能不能也做一个Agent？”',C.red,'×');
 card(s,6.95,1.55,5.55,3.95,'建议从这里开始','“谁在什么情况下，反复遇到了什么问题？”\n\n“现在如何解决，最痛的环节在哪里？”',C.green,'✓');
 bottom(s,'选题不追求大而创新，优先选择你亲自接触过、可以继续访谈和验证的工作场景。');
}
{
 const s=slide(); title(s,'使用统一模板完成你的第一次AI场景分析','作业提交模板');
 const a=[['1. Case名称','用一句话命名'],['2. 目标用户','谁会使用？'],['3. 使用场景','什么时候遇到问题？'],['4. 当前问题','现在怎么做？痛点在哪？'],['5. AI承担什么','输入、处理与输出'],['6. 人机分工','哪些给AI？哪些人确认？'],['7. 风险与兜底','AI可能错在哪里？'],['8. 验证指标','至少一个效率指标和质量指标']];
 a.forEach((d,i)=>card(s,0.55+(i%4)*3.17,1.48+Math.floor(i/4)*2.22,2.85,1.72,d[0],d[1],[C.blue,C.purple,C.orange,C.green][i%4],i+1));
 bottom(s,'不要求完整PRD和页面设计；优先把问题、流程、分工和验证方式说清楚。');
}
{
 const s=slide(); title(s,'提交前，先问自己五个问题','作业自检');
 const a=[['用户','我能否说清楚谁遇到了问题？'],['真实','这个问题是否真实发生？'],['具体','AI承担的任务是否足够具体？'],['兜底','是否设计了人工确认和错误处理？'],['验证','我是否知道怎么判断它有没有价值？']];
 a.forEach((d,i)=>card(s,0.68+(i%3)*4.2,1.52+Math.floor(i/3)*2.18,3.82,1.72,d[0],d[1],[C.blue,C.purple,C.orange,C.green,C.cyan][i],i+1));
 bottom(s,'第一份作业不追求“大而创新”，只追求“真实、具体、可以验证”。');
}
{
 const s=pptx.addSlide();s.background={color:C.navy};
 s.addShape(S.ellipse,{x:9.6,y:-1.2,w:4.8,h:4.8,fill:{color:C.blue,transparency:42},line:{color:C.blue,transparency:100}});
 s.addText('今天带走四个结论',{x:0.82,y:0.68,w:7.8,h:0.65,fontSize:30,bold:true,color:C.white,margin:0});
 const a=['AI是目标，机器学习是方法，NLP是领域，AIGC是能力形态','AI产品不只是模型，还包含知识、工具、流程与权限','分析AI Case要考虑问题、分工、风险、兜底和验证','课后回到自己的工作场景，找到一个真实问题'];
 a.forEach((v,i)=>{s.addShape(S.ellipse,{x:0.88,y:1.72+i*0.93,w:0.44,h:0.44,fill:{color:[C.cyan,C.purple,C.orange,C.green][i]},line:{color:C.white,transparency:100}});s.addText(String(i+1),{x:0.88,y:1.84+i*0.93,w:0.44,h:0.14,fontSize:9,bold:true,color:C.white,align:'center',margin:0});s.addText(v,{x:1.58,y:1.72+i*0.93,w:10.2,h:0.46,fontSize:15.5,bold:true,color:C.white,margin:0,fit:'shrink'});});
 s.addShape(S.line,{x:0.85,y:5.75,w:11.4,h:0,line:{color:'4A6486',width:1}});
 s.addText('课后作业：在自己的工作场景中，完成一份AI场景分析卡',{x:0.88,y:6.08,w:9.7,h:0.3,fontSize:13,bold:true,color:C.cyan,margin:0});
 s.addText('下一课：AI产品经理岗位认知与转型路径',{x:0.88,y:6.58,w:7.0,h:0.22,fontSize:10,color:'91A4BE',margin:0});
}

pptx.writeFile({ fileName: generated }).catch(e=>{console.error(e);process.exit(1)});
