const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第05课：Agent与受控工具调用';
pptx.title = 'Agent与受控工具调用';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };
const S = pptx.ShapeType;
const ROOT = '课程生产/第一阶段/第05课_Agent与受控工具调用_重做工作区';
const OUT = `${ROOT}/10_正式PPT_第05课_Agent与受控工具调用.pptx`;
const C = {
  bg:'F4F0E8', paper:'FBF9F4', ink:'20211F', text:'292A27', muted:'77766F', line:'D8D0C2',
  white:'FFFFFF', red:'C35331', teal:'387B77', orange:'D18A38', purple:'74546A',
  green:'527553', blue:'405D78', softRed:'F2E2DB', softTeal:'DDEBE8',
  softOrange:'F2E7D5', softPurple:'E8E0E7', softGreen:'E2E9DE', softBlue:'DDE5EB'
};
const A=[C.red,C.teal,C.orange,C.purple,C.green,C.blue];
const F=[C.softRed,C.softTeal,C.softOrange,C.softPurple,C.softGreen,C.softBlue];
function tx(s,t,x,y,w,h,o={}) {
  s.addText(t,{x,y,w,h,margin:0,fontFace:'PingFang SC',fontSize:16,color:C.text,
    fit:'shrink',valign:'top',breakLine:false,...o});
}
function sh(s,type,x,y,w,h,fill,line=fill,o={}) {
  s.addShape(type,{x,y,w,h,fill:{color:fill},line:{color:line,width:1},...o});
}
function pill(s,t,x,y,w,c) {
  sh(s,S.roundRect,x,y,w,.36,c,c,{rectRadius:.05});
  tx(s,t,x+.04,y+.08,w-.08,.18,{fontSize:10,bold:true,color:C.white,align:'center'});
}
function mod(n) {
  if(n<=8)return'01 Agent边界';
  if(n<=13)return'02 最小循环';
  if(n<=22)return'03 工具调用';
  if(n<=30)return'04 动态路径';
  if(n===31)return'课间休息';
  if(n<=36)return'05 纸面设计';
  if(n<=45)return'06 风险与接管';
  if(n<=51)return'07 Coze测试';
  if(n<=53)return'08 MCP与多Agent';
  return'09 总结与作业';
}
function base(n,title) {
  const s=pptx.addSlide();
  s.background={color:C.bg};
  pill(s,String(n).padStart(2,'0'),.62,.3,.62,A[(n-1)%6]);
  tx(s,mod(n).replace(/^\d+\s*/,''),1.42,.38,5.8,.18,
    {fontSize:10,bold:true,color:C.muted,charSpacing:1});
  tx(s,title,.62,.78,12,.54,{fontSize:25,bold:true,color:C.ink});
  sh(s,S.line,.62,1.42,12.05,0,C.line,C.line);
  tx(s,'AI 产品经理系统课程 · 第一阶段',.62,7.06,4.5,.16,{fontSize:8,color:C.muted});
  tx(s,mod(n),8,7.06,3.9,.16,{fontSize:8,color:C.muted,align:'right'});
  tx(s,String(n).padStart(2,'0'),12.1,7.04,.55,.18,{fontSize:8.5,color:C.muted,align:'right'});
  return s;
}
function contentSlide(n,title,body,tag='') {
  const s=base(n,title);
  const parts=body.split('|');
  const cols=parts.length<=2?parts.length:Math.min(parts.length,4);
  const rows=Math.ceil(parts.length/cols);
  const gap=.24,x0=.78,total=11.78,w=(total-gap*(cols-1))/cols;
  const h=rows===1?3.85:(rows===2?1.9:1.35);
  parts.forEach((p,i)=>{
    const [head,...rest]=p.split('：');
    const x=x0+(i%cols)*(w+gap), y=1.78+Math.floor(i/cols)*(h+.24), c=A[i%6];
    sh(s,S.roundRect,x,y,w,h,C.paper,C.line,{rectRadius:.05,
      shadow:{type:'outer',color:'000000',blur:1,angle:45,distance:1,opacity:.05}});
    sh(s,S.rect,x,y,.07,h,c,c);
    tx(s,head,x+.22,y+.2,w-.42,.42,{fontSize:parts.length>4?15:18,bold:true,color:C.ink,
      align:parts.length>2?'center':'left'});
    tx(s,rest.join('：'),x+.22,y+.72,w-.44,h-.9,{fontSize:parts.length>4?11.5:14.5,
      color:C.muted,align:parts.length>2?'center':'left',valign:'mid'});
  });
  if(tag) {
    sh(s,S.roundRect,1.02,6.1,11.28,.54,F[(n-1)%6],F[(n-1)%6],{rectRadius:.05});
    tx(s,tag,1.25,6.23,10.82,.25,{fontSize:14,bold:true,color:A[(n-1)%6],align:'center'});
  }
  return s;
}
function flowSlide(n,title,steps,tag='') {
  const s=base(n,title);
  const gap=.16,x0=.78,total=11.78,w=(total-gap*(steps.length-1))/steps.length,y=2.35,h=2.35;
  steps.forEach((p,i)=>{
    const [head,...rest]=p.split('：');
    sh(s,S.roundRect,x0+i*(w+gap),y,w,h,F[i%6],A[i%6],{rectRadius:.05});
    tx(s,String(i+1).padStart(2,'0'),x0+i*(w+gap)+.12,y+.13,.4,.24,{fontSize:11,bold:true,color:A[i%6]});
    tx(s,head,x0+i*(w+gap)+.14,y+.62,w-.28,.5,{fontSize:16,bold:true,color:C.ink,align:'center'});
    tx(s,rest.join('：'),x0+i*(w+gap)+.14,y+1.25,w-.28,.68,{fontSize:12.5,color:C.muted,align:'center'});
    if(i<steps.length-1)tx(s,'→',x0+i*(w+gap)+w-.03,y+1.0,.22,.35,{fontSize:20,bold:true,color:C.muted,align:'center'});
  });
  if(tag) {
    sh(s,S.roundRect,1.02,5.62,11.28,.58,C.paper,C.line,{rectRadius:.05});
    tx(s,tag,1.25,5.77,10.82,.28,{fontSize:14.5,bold:true,color:C.teal,align:'center'});
  }
  return s;
}

// cover
{
  const s=pptx.addSlide(); s.background={color:C.bg};
  sh(s,S.rect,0,0,.22,7.5,C.teal);
  tx(s,'05',.88,.72,1.1,.66,{fontSize:35,bold:true,color:C.teal});
  tx(s,'Agent与受控工具调用',.88,1.58,9.4,.72,{fontSize:36,bold:true,color:C.ink});
  tx(s,'当任务路径无法完全提前写死，如何让模型在目标、工具与权限约束下判断下一步',.92,2.62,10.1,.58,{fontSize:17,color:C.muted});
  ['目标','工具','观察','受控'].forEach((x,i)=>pill(s,x,9.55,1.15+i*1.05,2.1,A[i+1]));
  pill(s,'第一阶段 · 第05课',9.65,6.43,2.55,C.teal);
}

const D=[
['今天不是学习“更会聊天”','不是：记住更多Agent术语|不是：从空白搭一个复杂平台Demo|而是：设计一个会用工具、会停止、可验证的执行闭环','从回答问题，走向受控完成任务。'],
['今天要完成一份受控Agent方案','会判断：什么时候需要Agent|会设计：目标、工具、参数与动态路径|会控制：确认、异常、停止与人工接管|会验证：用5条测试检查真实行为','课堂产出：《Agent受控执行设计与测试表》V0.1'],
['一个不能提前写死的门店任务','用户：最近门店经营有问题，帮我找原因并给建议|先查什么：门店、日期、关注指标|可能结果：新客下降 / 补贴异常 / 数据不足','结果不同，下一步也不同。'],
['查完数据后，下一步并不固定','新客下降：查询新客活动规则|补贴异常：查询成本限制|数据不足：追问是否扩大时间范围','同一目标，工具返回会改变执行路径。'],
['Workflow与Agent不是高低关系','Workflow：路径主要提前规定；稳定、可预测|Agent：围绕目标，根据状态判断下一步|组合：Workflow控制主干，Agent处理局部动态判断','选择标准是任务特征，不是谁更“高级”。'],
['Chatbot、Workflow、Agent分别描述什么','Chatbot：对话交互形态|Workflow：预设执行路径|Agent：目标驱动的动态执行','有聊天框不等于有Agent。'],
['三个任务，你会怎么选','固定报销审批：Workflow|开放式产品问答：Chatbot/大模型|根据多次查询结果制定方案：评估Agent','先说任务路径，再说技术名词。'],
['Agent的最小运行循环','目标：最终要完成什么|状态：当前知道什么、缺什么|行动：追问、调用工具或输出|观察：工具返回和外部变化|结束：继续、确认、停止或转人工','循环的每一步都应可观察、可限制。'],
['第一步不是调用工具，是看当前状态','已知：用户关注新客下降|缺少：门店ID、开始日期、结束日期|下一步：追问必要信息','信息不完整时，不调用工具、不自行猜测。'],
['门店经营助手怎样走完一次循环','目标：解释新客下降并给建议|补参：确认门店与日期|查询：读取经营指标|观察：获得诊断标签|行动：查询规则并给建议|结束：建议完成或确认后建草稿','每次行动都基于当前状态。'],
['工具结果怎样改变下一步','new_customer_decline：查询新客活动规则|subsidy_efficiency_risk：转查补贴限制|insufficient_data：追问扩大日期','Agent不是固定把所有工具调用一遍。'],
['产品经理应该观察什么','目标：系统想完成什么|信息：当前是否足够|选择：为什么调用这个工具|参数：实际传了什么|结果：工具真实返回什么|状态：下一步和停止是否正确','不要只验收最后一句话。'],
['从“会说”到“能做”缺什么','模型擅长：理解、分类、生成、总结|真实任务需要：查实时数据、读规则、创建草稿|中间能力：外部工具与应用系统','工具把语言能力连接到真实任务。'],
['Function Calling解决什么','工具选择：模型表达想调用哪个工具|结构参数：模型表达需要传入什么|不负责：真正执行、权限校验和风险控制','Function Calling是结构化调用表达。'],
['模型提出调用，应用真正执行','模型：选择工具并生成参数|应用：校验参数、权限、风险与确认|工具：执行API、数据库或Workflow|模型：读取结果，继续或停止','模型提出调用 ≠ 操作已经发生。'],
['一个含糊的工具为什么容易被选错','名称：query_data|描述：查询数据|缺失：查什么、何时用、传什么、失败怎么办','工具描述模糊，模型选择和系统校验都无从下手。'],
['把query_data改成可用工具','用途：查询指定门店、指定日期范围的经营指标|必填：门店ID、开始/结束日期、指标|不适用：不查活动规则、不创建活动|权限：不访问未授权门店','工具说明同时服务模型选择和产品治理。'],
['工具设计十字段','识别：名称、用途、使用条件|输入：参数、类型、必填/选填|输出：成功返回、异常返回|控制：权限、确认','这是一份模型、应用和产品共同使用的契约。'],
['参数不仅要有，还要能校验','门店：store_id是否存在且有权限|日期：范围是否有效、是否超限|指标：是否在允许枚举中|确认：写入前是否真实获得用户确认','参数正确性由应用校验，不靠模型自觉。'],
['“查询失败”为什么无法驱动下一步','模糊：查询失败|结构化：STORE_NOT_FOUND / retryable=false|下一步：请用户核对门店ID','错误也是Agent继续判断所需的状态。'],
['课堂微练习：补全两个工具','工具A：search_rules｜搜索规则|工具B：create_activity｜创建活动|补充：用途、必填、返回、异常、权限、确认','8分钟填写，随后互换检查。'],
['本课只使用三个工具','query_store_metrics：查询经营指标与诊断标签|search_activity_rules：查询活动或成本规则|create_campaign_draft：确认后创建草稿','工具少而清晰，优于工具多而重叠。'],
['工具一：查询经营数据','用途：查询指定门店和日期的经营指标|输入：store_id、start_date、end_date、metrics|输出：数值、变化、完整性、诊断标签与依据|风险：读取类，先校验门店权限','异常阈值由确定性规则返回，不由模型发明。'],
['工具二：查询活动规则','用途：查询活动或成本规则|输入：store_id、rule_type、query_date|输出：适用范围、门槛、限制、版本与来源|风险：读取知识，可自动调用并留痕','规则冲突时不由模型自行裁决。'],
['工具三：创建活动草稿','用途：使用已确认参数创建草稿|输入：活动、人群、优惠、门槛、时间、预算|控制：user_confirmed必须为true|边界：只有创建草稿能力，没有发布能力','写入前必须展示完整参数。'],
['路径一：新客下降','查数据：返回new_customer_decline|查规则：获得适用拉新规则|生成：给出有依据的活动建议|确认：用户确认完整参数|执行：创建活动草稿','建议可以自动生成，写入必须确认。'],
['路径二：补贴异常','查数据：新客正常、补贴异常|改路：不查询新客活动规则|查规则：查询补贴与成本限制|输出：给补贴效率优化建议|停止：不创建新活动','这条路径证明结果会改变下一步。'],
['路径三：数据不足','查数据：只返回3天有效数据|追问：是否扩大到最近30天|同意：重新查询|拒绝：说明限制并停止','数据不足时，不用想象补齐数据。'],
['为什么不能把三个工具固定调用一遍','无效：数据正常仍查询活动|错误：补贴异常却生成拉新方案|风险：未确认就触发写入|成本：增加调用、时延与排错难度','Agent的价值是有依据地选择，而不是多调用。'],
['休息5分钟','学员：保存工具练习，回来完成Agent V0.1|老师：检查Coze、三个工具和T1—T5备用记录','回来后先做纸面设计，再看平台Demo。'],
['先设计，再打开平台','平台按钮：会变化|产品骨架：目标、工具、边界、异常、停止|顺序：纸面V0.1 → 同伴检查 → Coze验证','避免只会模仿平台操作。'],
['案例背景与可用能力','角色：连锁餐饮门店店长|问题：经营异常，需要分析和建议|工具：查询指标、查询规则、创建草稿|限制：不能直接发布活动','不要增加案例中不存在的工具。'],
['8分钟完成Agent V0.1','填写：服务对象、目标、完成标准|设计：必要输入与三个工具|标出：动态路径、确认、异常、停止|产出：一张可评审的方案草稿','现在开始，8分钟。'],
['交换检查四个问题','目标：Agent知道何时算完成吗|工具：职责清晰且不重复吗|风险：写入前有确认吗|异常：失败后知道如何处理吗','4分钟，两人互查。'],
['一份可控Agent方案应回答十件事','用户、问题、目标、完成标准|输入、工具、动态路径、输出|确认、异常、停止、人工接管','不仅设计成功路径，也设计如何安全失败。'],
['Agent可以建议，不代表可以执行','建议：生成运营方案|草稿：写入系统但仍可修改|发布：影响真实用户和预算|资金/删除：可能难以撤销','能力判断之后，还要做风险判断。'],
['从读取到发布，风险逐步上升','公开知识：可自动并记录|业务数据：校验身份与权限|生成建议：展示依据|创建草稿：明确确认|对外发布：强确认或审批|不可逆动作：默认禁止自动执行','自动化程度应随风险变化。'],
['两条轴判断自动化程度','横轴：出错后的影响，从低到高|纵轴：是否容易撤销，从容易到困难|高影响＋难撤销：默认不让Agent独立完成','风险不是只看动作名称。'],
['信息不足：追问，不猜','错误：自行猜测门店和日期|正确：指出缺失字段并只追问必要信息','补全前不调用工具。'],
['工具失败：说明，不编','无结果：查询成功但没有数据|调用失败：接口未完成任务|正确：展示状态、重试条件或停止建议','不能用看似合理的内容填补失败。'],
['权限超界：停止，不绕过','用户要求：直接发布活动|当前能力：只有创建草稿|正确：说明边界，确认后可建草稿，发布转人工','没有工具，就是产品边界。'],
['重复失败：停止并转人工','控制：设置最大重试次数|停止：连续失败或结果冲突|接管：保留目标、参数、调用和失败位置','人工不应让用户从头再说。'],
['一句“确认吗”为什么不够','对象：哪家门店|方案：人群、优惠与门槛|时间：开始和结束日期|成本：预算上限|后果：只建草稿，还是正式发布','确认必须让用户知道将要发生什么。'],
['转人工时不要让用户从头再说','用户目标：原始需求|已确认：门店、日期、方案|执行记录：工具、参数、结果|失败位置：为什么无法继续|人工建议：下一步应处理什么','上下文完整，人工接管才有效率。'],
['Coze演示看什么','1 信息是否完整|2 选择了哪个工具|3 参数传了什么|4 工具返回什么|5 下一步如何变化|6 何时停止|7 何时确认','不只看最终回答是否流畅。'],
['三个模拟工具与固定数据','数据工具：返回指标、诊断标签和异常|规则工具：返回适用规则、限制和来源|草稿工具：确认后返回draft_id|数据范围：全部使用教学模拟数据','不连接真实公司系统。'],
['五条测试，不只测正常路径','T1：信息完整|T2：缺少参数|T3：结果改变路径|T4：工具失败|T5：越权请求','每条记录工具、参数、返回、下一步和是否通过。'],
['T1＋T2：正常完成与缺少参数','T1：完整门店、日期和新客问题 → 查询并给建议|T2：缺少门店和日期 → 先追问，不调用工具','一个测完成，一个测不乱动。'],
['T3：结果改变，路径也要改变','返回：新客正常、补贴异常|预期：不查新客活动规则|改路：查询补贴与成本限制|结束：给优化建议，不创建拉新草稿','这是本课验证Agent动态选择的核心测试。'],
['T4＋T5：工具失败与越权请求','T4：STORE_NOT_FOUND → 提示核对并停止|T5：要求直接发布 → 说明无发布权限|允许：信息完整且确认后，只创建草稿','受控失败比“强行完成”更专业。'],
['Function Calling、MCP与Agent','Function Calling：结构化表达工具与参数|MCP：标准化连接工具、资源和上下文|Agent：围绕目标判断下一步','MCP不是Agent，也不替Agent做决策。'],
['什么时候才需要Multi-Agent','可以考虑：专业子任务、不同工具或权限、可拆分校验|不建议：单Agent已能完成、角色边界不清、成本时延敏感','默认先用单Agent＋少量工具跑通闭环。'],
['四句话总结与Agent V1.0','Agent：围绕目标判断下一步|工具：职责、参数、权限、异常要清楚|受控：确认、停止、人工接管不可少|作业：自己的场景＋2—4工具＋5条测试','先证明必要、可控、可测试，再增加复杂度。']
];

D.forEach((d,i)=>{
  const n=i+2;
  const flowPages=new Set([12,16,27,28,29]);
  if(flowPages.has(n)) flowSlide(n,d[0],d[1].split('|'),d[2]);
  else contentSlide(n,d[0],d[1],d[2]);
});
if(pptx._slides.length!==54)throw new Error(`Expected 54 slides, got ${pptx._slides.length}`);
pptx.writeFile({fileName:OUT});
