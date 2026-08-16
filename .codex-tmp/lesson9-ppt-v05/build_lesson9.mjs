import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第09课_正式PPT.pptx";
const PREVIEW_DIR = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-ppt-v05/rendered";
const MONTAGE = "/Users/keivn/Project/AI-Course/.codex-tmp/lesson9-ppt-v05/第09课_总览.webp";

const C = {
  bg: "#07151F", bg2: "#0B202C", ink: "#F5F2E9", muted: "#A7B5BC",
  teal: "#36D6C2", teal2: "#168E85", orange: "#FF7657", yellow: "#F0C75E",
  line: "#26404A", white: "#FFFFFF", dark: "#031016", green: "#6EE7B7",
};

const slides = [
  {type:"cover", title:"AI产品评测与数据反馈闭环", sub:"第09课｜从“Demo能跑”到“有依据地上线”", pages:"1"},
  {type:"question", title:"AI产品上线前的评测问题", lead:"因为一次漂亮回答，无法代表真实用户、真实边界和真实版本。", items:["换一种问法，还能答对吗？","用户不给完整信息，会发生什么？","知识库、Prompt或模型更新后，结果会不会退化？"], pages:"14–16、25"},
  {type:"statement", title:"AI输出的不确定性", lead:"AI输出具有概率性；产品要验证的是一组场景下的稳定表现。", accent:"从单点演示，转向成组评测", pages:"25"},
  {type:"process", title:"AI产品的版本组合", lead:"任何一层变化，都可能改变最终输出。", items:["模型版本","System Prompt","知识库","工作流","工具/API"], footer:"评测结论必须绑定版本，而不是只写“模型效果不错”。", pages:"25–27"},
  {type:"compare", title:"AI产品测试的三个层次", leftTitle:"功能验收", left:["流程能不能走通","接口与权限是否正常"], rightTitle:"离线评测 × 线上监控", right:["固定样本上效果如何","真实流量中是否持续稳定"], accent:"上线前测得准，上线后看得住", pages:"27–29"},
  {type:"map", title:"本课内容与学习路径", items:["风险动机","Eval体系","Bad Case与回归","数据反馈闭环"], pages:"9、24、72"},
  {type:"statement", title:"大模型幻觉及其业务风险", lead:"当答案进入客服、营销、医疗、金融或企业流程，错误会变成业务不可控。", accent:"可接受的产品，不等于永不出错；而是错误可发现、可拦截、可追踪。", risk:true, pages:"17、20"},
  {type:"bullets", title:"AI错误的业务影响", lead:"错误的严重度，不由“错误率”单独决定。", items:["高频问题答错：持续侵蚀信任","低频高风险问题答错：直接触发合规或资金损失","语气看似确定：让用户更难识别风险"], pages:"18"},
  {type:"three", title:"大模型幻觉的三种类型", items:[{h:"事实幻觉",t:"编造不存在的政策、价格、数据或出处"},{h:"指令幻觉",t:"没有遵守限制条件，擅自承诺或执行"},{h:"上下文幻觉",t:"忽略当前会话、用户身份或前置条件"}], pages:"19"},
  {type:"statement", title:"AI产品评测的核心价值", lead:"如果错误不可预测、不可观测、不可复现，团队就无法决定是否上线。", accent:"评测的价值：把“感觉不稳定”变成可讨论的证据。", risk:true, pages:"20"},
  {type:"process", title:"AI产品的五道风险防线", items:["数据与知识","Prompt与规则","模型与工具","输出校验","人工兜底"], footer:"本课聚焦：如何用评测判断这些防线是否真的有效。", pages:"21"},
  {type:"bullets", title:"客服场景的主要测试风险", items:["虚构退款、促销或物流政策","未核验身份就泄露订单信息","越权承诺赔付或到账时间","情绪激化、攻击性或歧视性表达","没有答案时仍强行作答"], pages:"22"},
  {type:"three", title:"AI产品评测的三个常见误区", items:[{h:"只测正常问题",t:"边界、异常和高风险样本被遗漏"},{h:"只看平均分",t:"严重错误被大量普通样本稀释"},{h:"改完只测原题",t:"修复一个问题，却可能伤到其他能力"}], pages:"23"},
  {type:"statement", title:"从风险识别到评测体系", lead:"接下来把风险变成样本、指标、门槛和回归机制。", accent:"从“主观体验”进入“Eval体系”", pages:"23–24"},

  {type:"section", no:"01", title:"AI产品评测体系的构建", sub:"固定题、统一判卷、重复运行", pages:"24"},
  {type:"compare", title:"AI测试与传统测试的差异", leftTitle:"传统软件", left:["输入—输出相对确定","Pass / Fail 边界清晰","相同版本容易复现"], rightTitle:"AI产品", right:["同一输入可能多种合理答案","质量标准包含主观判断","模型与上下文让结果波动"], pages:"25"},
  {type:"four", title:"AI产品测试的四个难点", items:["输出不确定","标准不唯一","覆盖成本高","版本变化快"], footer:"解决方案不是找一个万能指标，而是建立组合评测系统。", pages:"26"},
  {type:"six", title:"AI产品测试的六个维度", items:["功能","AI任务效果","可用性","非功能","鲁棒性","安全与伦理"], pages:"27"},
  {type:"statement", title:"本课的重点评测范围", lead:"先确保流程可用，再判断回答是否正确、相关、完整、安全、符合风格。", accent:"其他测试类型最后统一速览", pages:"28"},
  {type:"process", title:"AI产品评测的三种执行方式", items:["用户场景模拟","人工抽测","自动化脚本"], footer:"小规模靠人工理解，大规模靠自动化复现；两者必须互相校准。", pages:"29"},
  {type:"three", title:"AI产品评测的判分难点", items:[{h:"非确定",t:"同一题多次结果不完全相同"},{h:"主观标准",t:"正确之外，还有语气、风格和完整性"},{h:"规模覆盖",t:"真实用户表达远多于测试人员能穷举的范围"}], pages:"30"},
  {type:"statement", title:"量化评测的作用", lead:"量化的作用，是让版本可比较、问题可定位、上线门槛可执行。", accent:"没有统一尺子，就没有可靠回归。", pages:"32"},
  {type:"process", title:"评测对象与版本控制", items:["记录版本组合","锁定评测集","统一评分规则","重复运行","比较差异"], footer:"每次报告都要能回答：测的是哪个版本？用的哪套题？谁来判？", pages:"46–50"},
  {type:"bullets", title:"评测集的四类样本来源", items:["真实用户问题：最接近线上分布","历史 Bad Case：最值得回归","业务规则与红线：最不能犯错","AI扩写变体：提高表达覆盖"], pages:"47、54"},
  {type:"five", title:"评测集的五类样本", items:["典型","边界","异常","无答案","高风险"], footer:"比例不必平均；应按业务频率与风险加权。", pages:"47"},
  {type:"table", title:"评测用例的字段设计", headers:["字段","作用"], rows:[["Case ID / 场景","确保可追踪"],["用户输入 / 上下文","还原真实条件"],["参考答案 / 评分点","明确判题依据"],["风险等级 / 标签","支持加权与筛选"],["版本 / 实际输出","支持回归比较"]], pages:"47–50"},
  {type:"case", title:"评测用例示例：促销叠加规则", q:"用户：会员券能和满减活动同时使用吗？", a:"评分点：说明适用条件；不得承诺所有门店通用；信息不足时引导核验活动页。", tags:["业务规则","高频","中风险"], pages:"47–49"},
  {type:"statement", title:"语义相似度的作用与边界", lead:"它比较答案在语义空间中的距离，而不是逐字比对。", accent:"适合判断表达接近度，不适合单独判断事实真伪。", pages:"31、33"},
  {type:"visual", title:"余弦相似度的基本原理", lead:"方向越接近，相似度越高；但“方向接近”不代表关键事实一定正确。", metric:"0.86", metricLabel:"语义接近", pages:"33"},
  {type:"case", title:"语义相似度评测示例", q:"参考：订单付款后24小时内可以申请取消。", a:"AI：支付完成后的1天内，用户仍可发起取消申请。", tags:["核心条件一致","表达不同","高相似"], pages:"34–38"},
  {type:"compare", title:"语义相似度的局限", leftTitle:"参考答案", left:["退款到账：3—5个工作日"], rightTitle:"AI答案", right:["退款到账：5—7个工作日","其余表达几乎完全相同"], accent:"一个关键数字错误，足以让答案不可用", risk:true, pages:"40–41"},
  {type:"four", title:"AI产品的四种评测方法", items:["精确匹配 / 规则","语义相似度","LLM-as-Judge","人工评审"], footer:"客观字段用规则，开放答案用语义与裁判，关键决策由人工校准。", pages:"42、49"},
  {type:"method", title:"方法一：精确匹配与规则", lead:"适合明确、可枚举、不可妥协的判断。", items:["关键词、格式、JSON Schema","金额、日期、枚举值","禁答词、敏感信息、必含提示"], accent:"优点：稳定、便宜、可解释｜局限：覆盖不了开放表达", pages:"42、49"},
  {type:"method", title:"方法二：语义相似度", lead:"适合参考答案明确、表达方式多样的任务。", items:["FAQ问答","摘要主旨","意图或主题匹配"], accent:"优点：规模化｜局限：可能放过关键事实错误", pages:"31–42"},
  {type:"method", title:"方法三：LLM-as-Judge", lead:"让模型按照评分维度、评分点和红线给出判断。", items:["开放式回答","多维质量评分","批量解释失败原因"], accent:"优点：灵活｜局限：裁判也会偏、会漂移、会被提示影响", pages:"49、55–57"},
  {type:"method", title:"方法四：人工评审", lead:"用于高风险终审、疑难样本和自动评分校准。", items:["定义金标准","复核边界案例","抽查自动评测偏差"], accent:"不要把人工评审当成“无限人肉测试”。", pages:"58"},
  {type:"five", title:"AI输出质量的五个评测维度", items:["正确","相关","完整","安全","风格"], footer:"不同业务可以增删维度，但必须写清定义、评分点和红线。", pages:"48"},
  {type:"table", title:"评测维度与判分方法的匹配", headers:["维度","优先方法","提醒"], rows:[["正确性","规则＋人工/LLM","核对关键事实与数字"],["相关性","语义＋LLM","避免答非所问"],["完整性","评分点覆盖","检查必要步骤"],["安全性","规则＋人工","红线单独统计"],["风格","LLM＋抽检","统一语气与格式"]], pages:"49、59"},
  {type:"statement", title:"自动化评测的适用场景", lead:"当题库从20条增长到2000条，人工逐题检查不再可持续。", accent:"自动化负责规模，人工负责标准和校准。", pages:"44–45"},
  {type:"visual", title:"Eval的定义", lead:"固定一套题，用统一规则判卷，并在每次版本变化后重复运行。", metric:"同题", metricLabel:"同尺 · 同流程", pages:"46"},
  {type:"process", title:"评测集的设计方法", items:["按业务场景分层","加入边界与无答案","保留历史 Bad Case","标记频率与风险"], footer:"题库代表什么，最终分数就代表什么。", pages:"47"},
  {type:"process", title:"评分点与红线的定义", items:["写清理想答案","拆出必含信息","定义可接受变体","列出一票否决项"], footer:"先写判卷规则，再看模型答案，避免事后迁就。", pages:"48"},
  {type:"process", title:"评分方法的选择", items:["规则判硬条件","相似度看语义","LLM判开放质量","人工做终审校准"], footer:"组合方法比单一总分更可信。", pages:"49"},
  {type:"four", title:"自动评测中心的核心模块", items:["用例库","批量运行","评分与归因","版本对比"], footer:"产品经理不一定写平台，但必须定义平台要回答的问题。", pages:"50"},
  {type:"demo", title:"批量评测的执行与结果", items:["选择版本与评测集","批量发送请求并留存输出","按维度自动评分","筛选低分与红线用例","生成版本对比报告"], pages:"51–52"},
  {type:"process", title:"自动测试 Agent 的工作流程", items:["Test Case","被测产品","Judge","结果库","报告"], footer:"关键不是“用了Agent”，而是过程可重复、结果可审计。", pages:"53"},
  {type:"process", title:"AI扩写评测用例的方法", items:["给定原始样本","指定变化维度","生成表达变体","人工抽检去重","加入候选题库"], footer:"扩写用于覆盖，不用于凭空创造业务规则。", pages:"54"},
  {type:"case", title:"Judge Prompt 的设计", q:"请按正确性、完整性、安全性分别评分，并引用答案证据。", a:"红线：若虚构政策、越权承诺或泄露隐私，则安全性直接为0，并标记原因。", tags:["维度定义","评分锚点","输出JSON"], pages:"55–57"},
  {type:"three", title:"LLM裁判的三个局限", items:[{h:"偏好漂移",t:"模型或Prompt变化会改变评分尺度"},{h:"位置与措辞偏差",t:"更长、更像参考答案的回答可能被高估"},{h:"自评偏差",t:"同源模型可能偏爱自己的表达方式"}], pages:"55–57"},
  {type:"statement", title:"人工评审的校准作用", lead:"定期抽取自动高分、低分和争议样本，与人工金标准比较。", accent:"当裁判偏了，先修评分体系，再讨论被测模型。", pages:"58"},
  {type:"table", title:"评测报告的结构", headers:["层级","必须回答"], rows:[["总体","通过率、红线数、核心维度"],["场景","哪个场景退化或提升"],["问题","低分样本为什么失败"],["版本","相对基线变化多少"],["决策","上线、灰度、退回或转人工"]], pages:"60–62"},
  {type:"four", title:"AI产品评测的成熟度", items:["L1 人工抽查","L2 固定题库","L3 自动评测","L4 线上反馈驱动"], footer:"成熟度提升的标志：结果越来越可复现，反馈越来越快进入下一轮。", pages:"60–62"},
  {type:"four", title:"AI产品经理在评测中的职责", items:["测什么","怎样算好","何时能上线","失败如何进入优化"], footer:"评测不是测试同学的收尾工作，而是产品定义的一部分。", pages:"63"},

  {type:"two", title:"可用性与用户体验测试", leftTitle:"可用性", left:["任务能否完成","提示与反馈是否清楚","失败后能否恢复"], rightTitle:"用户体验", right:["等待感与节奏","结果可理解、可编辑","信任与控制感"], pages:"65–66"},
  {type:"three", title:"非功能、鲁棒性与安全伦理测试", items:[{h:"非功能",t:"延迟、成本、并发、稳定性"},{h:"鲁棒性",t:"错别字、噪声、提示攻击、异常输入"},{h:"安全与伦理",t:"隐私、偏见、越权、合规与可追责"}], pages:"67–71"},
  {type:"section", no:"02", title:"Bad Case分析与版本回归", sub:"归因、修复、验证与发布", pages:"72、85–95"},
  {type:"process", title:"低分会话的 Bad Case 入库流程", items:["触发低分/红线","去敏与去重","人工复核","归因与定级","加入回归集"], footer:"不是每个点踩都是模型问题；先确认用户意图与业务规则。", pages:"85–86"},
  {type:"six", title:"Bad Case 的六类问题归因", items:["数据","Prompt","知识","流程","工具","产品边界"], footer:"归因错了，越优化越乱。", pages:"85–95"},
  {type:"process", title:"修复方案与产品实验", items:["问题归因","方案设计","Prompt / 流程 A/B","Eval验证","发布决策"], footer:"优化目标应对应具体失败类型，而不是笼统追求“更聪明”。", pages:"91–95"},
  {type:"process", title:"版本回归测试", items:["原失败样本","同类高风险样本","代表性正常样本","历史关键 Bad Case"], footer:"新版本只有同时通过修复验证与防退化验证，才有资格进入灰度。", pages:"60–63、91–95"},
  {type:"section", no:"03", title:"数据反馈与持续迭代", sub:"让真实使用进入下一轮优化", pages:"72–90"},
  {type:"statement", title:"AI产品数据反馈的价值", lead:"真正的壁垒，不只是拥有数据，而是能把反馈稳定地转成更好的产品。", accent:"采集—理解—修复—验证—再上线", pages:"73"},
  {type:"four", title:"AI产品数据的四类来源", items:["业务原始数据","用户输入与输出","显性反馈","隐性行为"], footer:"采集必须满足隐私、授权、最小必要和可追踪。", pages:"74–75"},
  {type:"compare", title:"显性反馈与隐性反馈", leftTitle:"显性反馈", left:["点赞 / 点踩","原因选择","文字评价"], rightTitle:"隐性反馈", right:["复制、收藏、分享","重生成、改写、放弃","转人工、任务完成"], pages:"77–79"},
  {type:"process", title:"作业类产品的行为反馈", items:["直接采用","部分采用","大幅修改","弃用"], footer:"修改前后差异，比一个“满意/不满意”更接近真实质量。", pages:"80–81"},
  {type:"bullets", title:"常见用户行为信号的解读", items:["收藏：结果具有复用价值","重生成：当前答案未满足需求","放大/细看：内容值得进一步判断","复制/导出：可能进入真实工作流","迅速退出：可能无价值，也可能已完成任务"], pages:"82–83"},
  {type:"process", title:"客服产品的反馈机制", items:["回答后轻量反馈","失败原因快捷选择","低置信度主动转人工","人工处理结果回流","关闭会话前确认解决状态"], footer:"反馈机制要低摩擦，否则只会收集到极端意见。", pages:"84"},
  {type:"three", title:"Bad Case 标注规范", items:[{h:"定义",t:"标签到底描述什么，不描述什么"},{h:"正例",t:"哪些样本应该被标记"},{h:"反例",t:"相似但不应标记的边界样本"}], pages:"87–90"},
  {type:"process", title:"数据反馈闭环的基本流程", items:["采集","清洗与标注","问题归因","产品优化","回归评测"], footer:"没有回归验证的“数据优化”，只是未经证明的修改。", pages:"76"},
  {type:"exercise", title:"课堂练习：完成一轮最小评测闭环", items:["设计8条用例：覆盖典型、边界、异常、无答案和高风险","选择6条运行并记录输出","找出2条 Bad Case，完成归因","修改一个版本点，再做1次版本对照"], footer:"产出：用例表＋评分结果＋Bad Case归因＋上线建议", pages:"102–112"},
  {type:"close", title:"评测结果与产品决策", items:["上线","灰度","退回优化","转人工兜底"], footer:"固定评测集 × 组合评分 × Bad Case × 版本回归 × 数据反馈闭环", pages:"60–63、72–101"},
];

function addShape(slide, geometry, pos, fill="none", lineFill="none", lineWidth=0, radius=0) {
  const config = { geometry, position: pos, fill, line: {style:"solid", fill:lineFill, width:lineWidth} };
  if (radius && ["rect","textbox","roundRect"].includes(geometry)) config.borderRadius = radius;
  return slide.shapes.add(config);
}

function addText(slide, text, pos, size=24, color=C.ink, bold=false, align="left", name) {
  const sh = addShape(slide, "textbox", pos, "none");
  if (name) sh.name = name;
  sh.text = text;
  sh.text.style = {fontSize:size, color, bold, alignment:align, fontFamily:"PingFang SC"};
  return sh;
}

function addChrome(slide, index, section="EVAL") {
  addText(slide, `AI PRODUCT · ${section}`, {left:68,top:28,width:300,height:22}, 13, C.teal, true);
  addText(slide, String(index+1).padStart(2,"0"), {left:1170,top:28,width:42,height:22}, 13, C.muted, true, "right");
  addShape(slide, "rect", {left:68,top:680,width:1144,height:2}, C.line);
  addShape(slide, "rect", {left:68,top:680,width:Math.max(18,1144*((index+1)/slides.length)),height:2}, C.teal);
}

function titleBlock(slide, title, lead) {
  addText(slide, title, {left:68,top:78,width:1110,height:68}, 38, C.ink, true, "left", "slide-title");
  if (lead) addText(slide, lead, {left:70,top:152,width:1080,height:58}, 21, C.muted, false, "left", "slide-lead");
}

function noteFor(d) {
  const pdf = "/Users/keivn/Project/AI-Course/准备材料/参考课程/知乎AI课程/正课/27.【测试与数据】大模型产品开发中的数据准备与反馈闭环.pdf";
  const docx = "/Users/keivn/Project/AI-Course/准备材料/参考课程/知乎AI课程/正课/录音/27.【测试与数据】大模型产品开发中的数据准备与反馈闭环_原文.docx";
  const guide = "/Users/keivn/Project/AI-Course/课程生产/第一阶段/第09课_AI产品测试、指标与Bad Case_重做工作区/第09课_课程设计与老师稿.md";
  return `[Sources]\n- ${pdf}${d.pages ? `（参考页：${d.pages}）` : ""}\n- ${docx}\n- ${guide}`;
}

function bulletList(slide, items, top=236, color=C.ink, numbered=false) {
  const n = items.length;
  const gap = Math.min(78, 360/Math.max(1,n));
  items.forEach((item,i)=>{
    const y = top+i*gap;
    addText(slide, numbered?String(i+1).padStart(2,"0"):"—", {left:76,top:y,width:54,height:38}, numbered?17:23, numbered?C.teal:C.orange, true);
    addText(slide, String(item), {left:140,top:y-2,width:1010,height:50}, 24, color, i===0 && n<=3);
  });
}

function renderSlide(p, d, i) {
  const slide = p.slides.add();
  slide.background.fill = d.risk ? "#1D1112" : C.bg;
  slide.speakerNotes.textFrame.setText(noteFor(d));
  slide.speakerNotes.setVisible(true);

  if (d.type === "cover") {
    addShape(slide,"rect",{left:0,top:0,width:24,height:720},C.teal);
    addText(slide,"AI PRODUCT MANAGER",{left:80,top:70,width:350,height:30},16,C.teal,true);
    addText(slide,d.title,{left:78,top:210,width:1040,height:150},58,C.ink,true,"left","deck-title");
    addText(slide,d.sub,{left:82,top:392,width:760,height:46},24,C.muted,false);
    addShape(slide,"rect",{left:82,top:490,width:260,height:6},C.orange);
    addText(slide,"EVAL · BAD CASE · REGRESSION · DATA LOOP",{left:82,top:520,width:700,height:26},15,C.muted,true);
    addText(slide,"09",{left:1000,top:486,width:160,height:110},82,C.line,true,"right");
    return slide;
  }
  if (d.type === "section") {
    addShape(slide,"rect",{left:0,top:0,width:1280,height:720},C.bg2);
    addText(slide,d.no,{left:72,top:100,width:240,height:170},122,C.teal,true);
    addShape(slide,"rect",{left:74,top:300,width:100,height:5},C.orange);
    addText(slide,d.title,{left:72,top:348,width:1080,height:80},48,C.ink,true);
    addText(slide,d.sub,{left:74,top:455,width:940,height:45},24,C.muted,false);
    const sectionLabel = d.no==="01" ? "EVAL SYSTEM" : d.no==="02" ? "BAD CASE" : "DATA LOOP";
    addChrome(slide,i,sectionLabel);
    return slide;
  }

  const sec = i<14 ? "WHY EVAL" : i<55 ? "EVAL SYSTEM" : i<60 ? "BAD CASE" : "DATA LOOP";
  addChrome(slide,i,sec);
  titleBlock(slide,d.title,d.type === "statement" ? undefined : d.lead);

  if (["statement","visual"].includes(d.type)) {
    addShape(slide,"rect",{left:70,top:250,width:8,height:245},d.risk?C.orange:C.teal);
    if (d.type==="visual" && d.metric) {
      addText(slide,d.metric,{left:120,top:245,width:380,height:160},90,C.teal,true);
      addText(slide,d.metricLabel,{left:130,top:420,width:390,height:42},24,C.muted,true);
      addShape(slide,"ellipse",{left:690,top:266,width:160,height:160},"none",C.teal,5);
      addShape(slide,"ellipse",{left:825,top:318,width:160,height:160},"none",C.orange,5);
      addText(slide,"答案A",{left:710,top:330,width:120,height:30},19,C.ink,true,"center");
      addText(slide,"答案B",{left:845,top:382,width:120,height:30},19,C.ink,true,"center");
    } else {
      addText(slide,d.accent||d.lead,{left:120,top:285,width:990,height:140},32,d.risk?C.orange:C.teal,true);
      if (d.lead && d.accent) addText(slide,d.lead,{left:120,top:445,width:980,height:80},22,C.muted,false);
    }
  } else if (["bullets","method","demo","exercise"].includes(d.type)) {
    bulletList(slide,d.items,226,C.ink,d.type==="demo"||d.type==="exercise");
    if(d.accent) addText(slide,d.accent,{left:140,top:600,width:990,height:38},18,C.teal,true);
    if(d.footer) addText(slide,d.footer,{left:140,top:605,width:990,height:36},18,C.muted,false);
  } else if (["compare","two"].includes(d.type)) {
    const x1=74,x2=666,w=540;
    addShape(slide,"rect",{left:x1,top:236,width:w,height:5},C.teal);
    addShape(slide,"rect",{left:x2,top:236,width:w,height:5},d.risk?C.orange:C.yellow);
    addText(slide,d.leftTitle,{left:x1,top:264,width:w,height:45},27,C.teal,true);
    addText(slide,d.rightTitle,{left:x2,top:264,width:w,height:45},27,d.risk?C.orange:C.yellow,true);
    d.left.forEach((t,j)=>addText(slide,"—  "+t,{left:x1,top:330+j*66,width:w-20,height:52},22,C.ink,j===0));
    d.right.forEach((t,j)=>addText(slide,"—  "+t,{left:x2,top:330+j*66,width:w-20,height:52},22,C.ink,j===0));
    if(d.accent) addText(slide,d.accent,{left:74,top:598,width:1130,height:38},20,d.risk?C.orange:C.teal,true,"center");
  } else if (["three","four","five","six"].includes(d.type)) {
    const raw = d.items.map(x=>typeof x==="string"?{h:x,t:""}:x);
    const cols = d.type==="three"?3:(d.type==="four"?2:(d.type==="five"?5:3));
    const rows = Math.ceil(raw.length/cols);
    const usableW=1136,gap=24,cellW=(usableW-gap*(cols-1))/cols;
    const startY=240,cellH=rows===1?270:150;
    raw.forEach((x,j)=>{
      const col=j%cols,row=Math.floor(j/cols),left=72+col*(cellW+gap),top=startY+row*(cellH+28);
      addText(slide,String(j+1).padStart(2,"0"),{left,top,width:52,height:30},15,C.teal,true);
      addShape(slide,"rect",{left,top:top+42,width:cellW,height:3},j%2?C.orange:C.teal);
      addText(slide,x.h,{left,top:top+65,width:cellW-10,height:48},25,C.ink,true);
      if(x.t) addText(slide,x.t,{left,top:top+120,width:cellW-16,height:76},18,C.muted,false);
    });
    if(d.footer) addText(slide,d.footer,{left:72,top:620,width:1136,height:36},18,C.muted,false,"center");
  } else if (["process","map"].includes(d.type)) {
    const items=d.items, n=items.length, gap=18, totalW=1136, w=(totalW-gap*(n-1))/n, y=300;
    items.forEach((t,j)=>{
      const x=72+j*(w+gap);
      if(j<n-1) addShape(slide,"rightArrow",{left:x+w-4,top:y+45,width:gap+14,height:30},C.line);
      const box=addShape(slide,"roundRect",{left:x,top:y,width:w,height:120},j===n-1?C.teal:C.bg2,C.line,1,16);
      addText(slide,String(j+1).padStart(2,"0"),{left:x+14,top:y+14,width:40,height:24},14,j===n-1?C.dark:C.teal,true);
      addText(slide,t,{left:x+12,top:y+53,width:w-24,height:48},21,j===n-1?C.dark:C.ink,true,"center");
    });
    if(d.footer) addText(slide,d.footer,{left:110,top:490,width:1060,height:55},20,C.muted,false,"center");
  } else if (d.type==="table") {
    const x=72,y=230,w=1136,rowH=66,col1=340;
    addShape(slide,"rect",{left:x,top:y,width:w,height:rowH},C.teal);
    addText(slide,d.headers[0],{left:x+18,top:y+17,width:col1-30,height:32},20,C.dark,true);
    addText(slide,d.headers[1],{left:x+col1+18,top:y+17,width:w-col1-30,height:32},20,C.dark,true);
    d.rows.forEach((r,j)=>{
      const yy=y+rowH*(j+1);
      addShape(slide,"rect",{left:x,top:yy,width:w,height:rowH},j%2?C.bg2:C.bg,C.line,1);
      addText(slide,r[0],{left:x+18,top:yy+17,width:col1-30,height:34},19,C.ink,true);
      addText(slide,r[1]+(r[2]?`｜${r[2]}`:""),{left:x+col1+18,top:yy+17,width:w-col1-30,height:34},18,C.muted,false);
    });
  } else if (d.type==="case") {
    addText(slide,"INPUT",{left:74,top:230,width:120,height:28},14,C.teal,true);
    addText(slide,d.q,{left:74,top:270,width:1060,height:80},27,C.ink,true);
    addShape(slide,"rect",{left:74,top:382,width:1130,height:2},C.line);
    addText(slide,"SCORING POINTS",{left:74,top:410,width:180,height:28},14,C.orange,true);
    addText(slide,d.a,{left:74,top:452,width:1080,height:95},22,C.muted,false);
    d.tags.forEach((t,j)=>{
      const x=74+j*220;
      addShape(slide,"roundRect",{left:x,top:580,width:196,height:42},C.bg2,C.line,1,18);
      addText(slide,t,{left:x+8,top:590,width:180,height:24},15,C.teal,true,"center");
    });
  } else if (d.type==="close") {
    const xs=[100,380,660,940];
    d.items.forEach((t,j)=>{
      addText(slide,String(j+1).padStart(2,"0"),{left:xs[j],top:280,width:120,height:50},28,C.teal,true,"center");
      addText(slide,t,{left:xs[j]-20,top:350,width:160,height:55},25,C.ink,true,"center");
    });
    addShape(slide,"rect",{left:110,top:470,width:1060,height:4},C.orange);
    addText(slide,d.footer,{left:120,top:510,width:1040,height:70},22,C.muted,true,"center");
  }
  return slide;
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(PREVIEW_DIR,{recursive:true});
  const p=Presentation.create({slideSize:{width:1280,height:720}});
  slides.forEach((d,i)=>renderSlide(p,d,i));
  for (const [i,slide] of p.slides.items.entries()) {
    const png=await p.export({slide,format:"png",scale:1});
    await writeBlob(`${PREVIEW_DIR}/slide-${String(i+1).padStart(2,"0")}.png`,png);
    const layout=await slide.export({format:"layout"});
    await fs.writeFile(`${PREVIEW_DIR}/slide-${String(i+1).padStart(2,"0")}.layout.json`,await layout.text());
  }
  await writeBlob(MONTAGE,await p.export({format:"webp",montage:true,scale:0.5}));
  const pptx=await PresentationFile.exportPptx(p);
  await pptx.save(OUT);
  console.log(JSON.stringify({slides:p.slides.items.length,out:OUT,montage:MONTAGE}));
}

main().catch(err=>{console.error(err);process.exitCode=1;});
