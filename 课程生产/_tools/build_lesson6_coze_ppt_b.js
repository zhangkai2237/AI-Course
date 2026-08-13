const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'AI 产品经理系统课程';
pptx.subject = '第06课：Coze综合实战';
pptx.title = 'Coze综合实战：从Prompt到Agent搭建小红书图文内容工厂';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'PingFang SC', bodyFontFace: 'PingFang SC', lang: 'zh-CN' };
const S = pptx.ShapeType;
const ROOT = '课程生产/第一阶段/第06课_Coze综合实战_小红书图文内容工厂_重做工作区';
const OUT = `${ROOT}/10_正式PPT_第06课_Coze综合实战_从Prompt到Agent.pptx`;
const C = {
  bg:'F4F0E8', paper:'FBF9F4', ink:'20211F', text:'292A27', muted:'77766F', line:'D8D0C2',
  white:'FFFFFF', dark:'172528', red:'B95843', teal:'2E7772', orange:'D18A38', purple:'705275',
  green:'527553', blue:'405D78', softRed:'F2E2DB', softTeal:'DDEBE8', softOrange:'F2E7D5',
  softPurple:'E8E0E7', softGreen:'E2E9DE', softBlue:'DDE5EB', dark2:'22363A'
};
const F = {red:C.softRed,teal:C.softTeal,orange:C.softOrange,purple:C.softPurple,green:C.softGreen,blue:C.softBlue};
function tx(s,t,x,y,w,h,o={}) {
  s.addText(t,{x,y,w,h,margin:0,fontFace:'PingFang SC',fontSize:16,color:C.text,
    fit:'shrink',valign:'top',breakLine:false,...o});
}
function sh(s,type,x,y,w,h,fill,line=fill,o={}) {
  s.addShape(type,{x,y,w,h,fill:{color:fill},line:{color:line,width:1},...o});
}
function pill(s,t,x,y,w,c,textColor=C.white) {
  sh(s,S.roundRect,x,y,w,.36,c,c,{rectRadius:.05});
  tx(s,t,x+.04,y+.08,w-.08,.18,{fontSize:10,bold:true,color:textColor,align:'center'});
}
function moduleInfo(n){
  if(n<=7)return {name:'01 最终效果与环境',c:C.teal,f:C.softTeal};
  if(n<=15)return {name:'02 P0 Prompt',c:C.red,f:C.softRed};
  if(n<=24)return {name:'03 P1 Workflow',c:C.teal,f:C.softTeal};
  if(n===25)return {name:'课间休息',c:C.orange,f:C.softOrange};
  if(n<=33)return {name:'04 P2 RAG',c:C.blue,f:C.softBlue};
  if(n<=42)return {name:'05 P3 Agent',c:C.purple,f:C.softPurple};
  if(n<=45)return {name:'06 四版本统一测试',c:C.orange,f:C.softOrange};
  return {name:'07 总结与作业',c:C.green,f:C.softGreen};
}
function base(n,title) {
  const m=moduleInfo(n),s=pptx.addSlide();
  s.background={color:C.bg};
  pill(s,String(n).padStart(2,'0'),.62,.3,.62,m.c);
  tx(s,m.name.replace(/^\d+\s*/,''),1.42,.38,5.8,.18,{fontSize:10,bold:true,color:C.muted,charSpacing:1});
  tx(s,title,.62,.78,12,.54,{fontSize:25,bold:true,color:C.ink});
  sh(s,S.line,.62,1.42,12.05,0,C.line,C.line);
  tx(s,'AI 产品经理系统课程 · 第一阶段',.62,7.06,4.5,.16,{fontSize:8,color:C.muted});
  tx(s,m.name,8.1,7.06,3.8,.16,{fontSize:8,color:C.muted,align:'right'});
  tx(s,String(n).padStart(2,'0'),12.1,7.04,.55,.18,{fontSize:8.5,color:C.muted,align:'right'});
  return s;
}
function cards(n,title,items,tag='') {
  const s=base(n,title),m=moduleInfo(n);
  const cols=items.length<=2?items.length:(items.length===3?3:(items.length===4?4:3));
  const rows=Math.ceil(items.length/cols),gap=.22,x0=.78,total=11.78,w=(total-gap*(cols-1))/cols;
  const h=rows===1?3.86:(rows===2?1.84:1.3);
  items.forEach((item,i)=>{
    const [head,...rest]=item.split('：'),x=x0+(i%cols)*(w+gap),y=1.75+Math.floor(i/cols)*(h+.22);
    sh(s,S.roundRect,x,y,w,h,C.paper,C.line,{rectRadius:.05,shadow:{type:'outer',color:'000000',blur:1,angle:45,distance:1,opacity:.05}});
    sh(s,S.rect,x,y,.07,h,m.c,m.c);
    tx(s,head,x+.2,y+.18,w-.4,.38,{fontSize:items.length>4?14.5:17,bold:true,color:C.ink,align:cols>2?'center':'left'});
    tx(s,rest.join('：'),x+.2,y+.65,w-.4,h-.83,{fontSize:items.length>4?11.5:14,color:C.muted,align:cols>2?'center':'left',valign:'mid'});
  });
  if(tag){
    sh(s,S.roundRect,1.02,6.1,11.28,.55,m.f,m.f,{rectRadius:.05});
    tx(s,tag,1.24,6.24,10.84,.24,{fontSize:14,bold:true,color:m.c,align:'center'});
  }
  return s;
}
function flow(n,title,steps,tag='') {
  const s=base(n,title),m=moduleInfo(n);
  const gap=.15,x0=.75,total=11.85,w=(total-gap*(steps.length-1))/steps.length,y=2.28,h=2.45;
  steps.forEach((p,i)=>{
    const [head,...rest]=p.split('：'),x=x0+i*(w+gap);
    sh(s,S.roundRect,x,y,w,h,i%2?C.paper:m.f,m.c,{rectRadius:.05});
    tx(s,String(i+1).padStart(2,'0'),x+.12,y+.14,.38,.22,{fontSize:10,bold:true,color:m.c});
    tx(s,head,x+.13,y+.65,w-.26,.46,{fontSize:steps.length>6?12.5:15,bold:true,color:C.ink,align:'center'});
    tx(s,rest.join('：'),x+.13,y+1.27,w-.26,.72,{fontSize:steps.length>6?10.5:12,color:C.muted,align:'center',valign:'mid'});
    if(i<steps.length-1)tx(s,'→',x+w-.04,y+1.02,.23,.34,{fontSize:19,bold:true,color:m.c,align:'center'});
  });
  if(tag){
    sh(s,S.roundRect,1.02,5.66,11.28,.58,C.paper,C.line,{rectRadius:.05});
    tx(s,tag,1.25,5.81,10.82,.27,{fontSize:14,bold:true,color:m.c,align:'center'});
  }
  return s;
}
function demo(n,title,project,actions,observe,returnPage){
  const m=moduleInfo(n),s=pptx.addSlide(); s.background={color:C.dark};
  pill(s,String(n).padStart(2,'0'),.62,.3,.62,m.c);
  tx(s,'COZE LIVE DEMO',1.42,.38,4.5,.18,{fontSize:10,bold:true,color:'9AB0AF',charSpacing:1.2});
  tx(s,title,.62,.82,12,.62,{fontSize:28,bold:true,color:C.white});
  sh(s,S.roundRect,.72,1.72,4.05,4.6,C.dark2,'365359',{rectRadius:.05});
  tx(s,'此处切换至',1.08,2.28,3.35,.4,{fontSize:18,color:'AFC1C0',align:'center'});
  tx(s,'Coze 投屏',1.02,2.85,3.48,.72,{fontSize:34,bold:true,color:C.white,align:'center'});
  pill(s,project,1.05,4.0,3.42,m.c);
  tx(s,`操作完成后切回 PPT 第 ${returnPage} 页`,1.0,5.23,3.5,.3,{fontSize:12,color:'AFC1C0',align:'center'});
  sh(s,S.roundRect,5.1,1.72,3.55,4.6,'1D3033','365359',{rectRadius:.05});
  tx(s,'现场动作',5.4,2.02,2.95,.3,{fontSize:16,bold:true,color:m.c});
  actions.forEach((a,i)=>{pill(s,String(i+1),5.42,2.62+i*.66,.34,m.c);tx(s,a,5.9,2.63+i*.66,2.38,.36,{fontSize:12.5,color:C.white});});
  sh(s,S.roundRect,8.92,1.72,3.7,4.6,'1D3033','365359',{rectRadius:.05});
  tx(s,'学员观察',9.22,2.02,3.05,.3,{fontSize:16,bold:true,color:C.orange});
  observe.forEach((a,i)=>{tx(s,'□',9.22,2.6+i*.67,.28,.3,{fontSize:15,bold:true,color:C.orange});tx(s,a,9.62,2.61+i*.67,2.55,.38,{fontSize:12.5,color:C.white});});
  tx(s,'AI 产品经理系统课程 · 第06课',.62,7.05,4.5,.16,{fontSize:8,color:'809796'});
  return s;
}
function quoteSlide(n,title,quote,items,tag=''){
  const s=base(n,title),m=moduleInfo(n);
  sh(s,S.roundRect,.88,1.78,11.56,1.28,m.f,m.f,{rectRadius:.05});
  tx(s,`“${quote}”`,1.2,2.11,10.9,.54,{fontSize:22,bold:true,color:m.c,align:'center'});
  const w=(11.38-.22*(items.length-1))/items.length;
  items.forEach((x,i)=>{const [h,...r]=x.split('：'),xx=.98+i*(w+.22);sh(s,S.roundRect,xx,3.5,w,2.12,C.paper,C.line,{rectRadius:.05});tx(s,h,xx+.15,3.83,w-.3,.38,{fontSize:15.5,bold:true,color:C.ink,align:'center'});tx(s,r.join('：'),xx+.16,4.45,w-.32,.72,{fontSize:12,color:C.muted,align:'center',valign:'mid'});});
  if(tag)tx(s,tag,1.1,6.15,11.1,.28,{fontSize:14,bold:true,color:m.c,align:'center'});
  return s;
}

// 01 cover
{
  const s=pptx.addSlide(); s.background={color:C.bg};
  sh(s,S.rect,0,0,.22,7.5,C.teal);
  tx(s,'06',.88,.7,1.1,.66,{fontSize:35,bold:true,color:C.teal});
  tx(s,'Coze综合实战',.88,1.54,7.7,.68,{fontSize:37,bold:true,color:C.ink});
  tx(s,'从Prompt到Agent搭建小红书图文内容工厂',.92,2.4,9.35,.56,{fontSize:21,bold:true,color:C.teal});
  tx(s,'同一个项目连续升级四次｜只做到Coze内搭建、运行与测试',.92,3.2,9.3,.36,{fontSize:15.5,color:C.muted});
  [['P0','Prompt',C.red],['P1','Workflow',C.teal],['P2','RAG',C.blue],['P3','Agent',C.purple]].forEach((a,i)=>{pill(s,a[0],9.65,1.08+i*1.02,.75,a[2]);tx(s,a[1],10.58,1.17+i*1.02,1.52,.26,{fontSize:14,bold:true,color:a[2]});});
  pill(s,'第一阶段 · 第06课',9.65,6.43,2.55,C.teal);
}

flow(2,'今天只做一个项目，连续升级四次',['P0：一个LLM节点','P1：固定内容SOP','P2：接入指定知识','P3：按目标选工具'],'后一个版本不是替代前一个版本，而是在真实边界上增加一层能力。');
cards(3,'3小时后，你会留下四个可运行版本',['P0 文案生成器：结构化Prompt＋两条测试','P1 图文Workflow：固定流程＋逐节点排错','P2 RAG Workflow：知识检索＋无命中边界','P3 内容策划Agent：追问、调用、确认与停止'],'课堂产物不是“看过四个Demo”，而是四份可以继续迭代的Coze项目。');
cards(4,'今天做到哪里，也明确不做到哪里',['今天完成：Coze内搭建、运行、测试和能力对照','今天完成：知识文件、工具调用、边界和失败测试','今天不做：独立网页、API、登录、支付','今天不做：真实账号、自动发布、MCP和Multi-Agent'],'先把能力组合跑通，不把课堂实操包装成已上线产品。');
quoteSlide(5,'先看最终版：它会追问、调用、确认和停止','先给我看，不要直接保存或发布。',['检查信息：缺什么就先问','调用P2：生成完整图文','遵守权限：不保存、不发布'],'演示时看调用过程，不只看最终文案。');
demo(6,'运行P3老师完成版','P3_小红书内容策划Agent',['输入统一最终请求','展开工具调用记录','查看P2内部结果','确认没有保存发布'],['有没有重复追问','调用了哪个工具','是否带知识与审核','是否遵守权限'],7);
cards(7,'开始前完成7项环境检查',['登录：进入课堂Coze工作区','模板：能打开P0—P3','模型：能选择指定模型','运行：能完成一次测试','记录：能看节点输入输出','材料：能访问知识文件','备用：知道图片Prompt路线'],'一个问题排查超过3分钟，切换老师中间版本；不让单点故障拖住全班。');

flow(8,'第一轮只增加一个结构化Prompt',['Start：接收变量','LLM：理解与生成','End：返回结果'],'Start和End只是最小运行载体，本轮核心仍是一次Prompt调用。');
cards(9,'先把会变化的信息做成变量',['topic：写什么主题','audience：写给谁看','goal：希望产生什么行为','source_material：允许使用哪些事实','tone：用什么表达风格'],'变量让同一套Prompt可以换主题、受众和事实，而不是每次重写整段。');
cards(10,'一个可测试Prompt至少说清五件事',['角色：模型以什么身份工作','任务：这一次要完成什么','输入：有哪些变量和事实','规则：结构、事实与禁止事项','输出：必须返回哪些字段'],'Prompt不是越长越好，而是输入、输出和边界能被测试。');
demo(11,'完成P0并绑定变量','P0_小红书文案生成器',['检查5个Start变量','复制LLM Prompt','逐项绑定变量','运行统一正常输入'],['运行输入是否非空','是否正好5个标题','是否使用事实素材','是否出现编造'],12);
cards(12,'正常输入先检查“是否按契约输出”',['字段：标题、正文、封面、标签是否齐全','数量：标题和标签数量是否符合要求','对象：内容是否真正面向初级产品经理','事实：观点是否来自source_material','风险：有没有公司、经历和效果编造'],'先验收契约和事实，再讨论个人风格偏好。');
quoteSlide(13,'删掉事实素材，看看Prompt会不会编','source_material留空',['最好：明确要求补材料','可接受：只给通用结构并标记无依据','不通过：补出经历、模块或效果'],'规则文字不能凭空创造可靠事实来源。');
cards(14,'一次只改一处，才知道为什么变好',['Bad Case：标题否定Prompt价值','单点修改：增加一条标题边界','相同输入：重新运行','前后记录：改善与新问题'],'同时改五处，只能得到新结果，得不到可解释的迭代结论。');
cards(15,'Prompt解决当前节点，但不能创造事实和流程',['能做：理解当前输入','能做：约束结构和表达','不能：获得未提供的可靠个人知识','不能：确定性控制多步流程','不能：动态选择和调用工具'],'下一轮不继续把Prompt写长，而是把稳定步骤拆开。');

cards(16,'P0的问题不是文案不够长，而是步骤挤在一起',['缺失：某一部分偶尔不输出','难改：角度和正文只能一起重跑','难查：不知道哪一步出错','失控：缺素材仍继续生成'],'Workflow解决的是步骤协作、变量传递和错误定位。');
flow(17,'把内容生产拆成一条稳定SOP',['输入检查：缺失就停止','内容角度：先做策划','标题正文：生成主内容','图片脚本：拆页面','图片：生成或备用','审核：事实与一致性','End：统一返回'],'路径可以提前确定，所以先用Workflow，不需要Agent。');
cards(18,'每个节点只承担一个主要职责',['输入检查：决定是否继续','内容角度：痛点、观点和大纲','标题正文：文案与封面','图片脚本：页数、版式和提示词','内容审核：事实、承诺和一致性'],'如果一个LLM节点仍做完所有任务，只是外面多画线，并没有真正拆流程。');
flow(19,'Workflow最容易错在变量没有传下去',['Start：原始事实','策划：content_plan','正文：copy_package','图片：image_script','审核：review_result'],'最终结果不对时，先看每个节点真实收到什么，不要先改Prompt。');
demo(20,'输入检查＋内容生产节点','P1_小红书图文Workflow',['配置缺信息条件','单测内容角度节点','绑定content_plan','单测标题正文节点'],['缺素材是否停止','上游输出是否传递','正文是否遵循大纲','事实是否仍受限制'],21);
demo(21,'图片脚本＋审核＋End','P1_小红书图文Workflow',['配置图片脚本','选择图片/Prompt路线','配置内容审核','绑定End并完整运行'],['页数是否一致','图片是否匹配正文','审核是否发现风险','失败是否有备用输出'],22);
cards(22,'P1必须跑过正常、缺素材和审核三条测试',['W1 正常：所有节点到End','W2 缺素材：后续节点不运行','W4 虚构数字：审核必须revise'],'跑通不是完成；边界和失败行为同样属于产品行为。');
quoteSlide(23,'页数写5，结果只出3页：问题在哪一层','page_count＝5',['输入：Start确实收到5','定位：图片脚本仍写死3','修复：让节点遵循变量'],'错误可以定位到具体节点，就是Workflow可维护性的价值。');
cards(24,'Workflow让流程可重复、错误可定位',['新增：确定性输入检查','新增：稳定SOP与局部重跑','新增：逐节点观察和审核','仍缺：指定个人与课程知识','仍缺：开放状态下动态选择'],'休息后不改主流程，只给它增加可检索知识。');

quoteSlide(25,'休息7分钟｜回来给Workflow接入知识','先保存P0和P1',['学员：保存项目与测试记录','老师：检查知识库和P2模板','备用：打开知识片段截图'],'第90分钟准时开始P2。');

cards(26,'P1流程稳定，但表达仍然像“通用编辑”',['开头：不一定像你的切入方式','节奏：句子长短和停顿不同','结构：解释概念的顺序不同','结尾：可能变成通用营销话术'],'这一轮RAG只学习怎样表达，不学习历史文档里的事实。');
cards(27,'语气样本不要求是小红书',['课程讲义：学习解释概念的方法','长文复盘：学习观点展开节奏','产品分析：学习拆解问题方式','社群答疑：学习口语程度与收口'],'语气来自样本，输出体裁仍由当前Prompt决定。');
cards(28,'事实、语气和体裁必须来自三个不同位置',['事实 source_material：本次唯一事实来源','语气 style_profile：历史文字抽象结果','体裁 target_format：当前任务要求'],'历史原文不能替本次输入补公司、项目、数据和结论。');
flow(29,'历史原文不能直接进入写作节点',['检索原文：找到3—5篇样本','提取语气：删除所有历史事实','生成新文：只接语气卡＋本次事实','双重审核：防事实迁移和抄句'],'变量隔离比一句“不要使用历史事实”更可靠。');
demo(30,'导入历史文字并完成语气版P2','P2_个人语气_RAG图文Workflow',['上传脱敏语气样本','抽查至少3个片段','生成抽象语气卡','运行P1/P2同题对照'],['语气卡是否只讲表达','原文是否进入写作节点','两版事实是否一致','非小红书样本能否迁移'],31);
cards(31,'同一份事实，P1和P2只应该差在表达',['事实集合：两版必须完全一致','开头节奏：P2是否更像老师','论述结构：差异能否由语气卡解释','输出体裁：P2仍然是小红书'],'P2不是知道得更多，而是在事实不变时表达得更一致。');
cards(32,'五条测试防止“学语气”变成“抄历史”',['历史事实：旧公司项目不能迁移','跨体裁：讲义语气也能写小红书','无命中：使用默认语气','防抄句：相似长句必须重写','缺素材：不能用历史事实补足'],'语气相似不能抵消事实错误。');
cards(33,'语气来自样本，事实来自本次输入',['语气：检索历史文字并抽象','事实：只使用source_material','体裁：由target_format和Prompt规定','审核：同时查事实迁移与风格匹配'],'下一轮让Agent决定何时调用这个个人语气Workflow。');

cards(34,'P2每次走同一路径，但用户目标并不相同',['信息完整：不需要追问','只要选题：不需要生成全文','需要近期资料：才需要搜索','只想查看：不应该保存','要求发布：当前没有权限','信息不足：应该先补齐'],'动态路径有价值时，才引入Agent。');
cards(35,'P3只有三个工具，而且职责不能重叠',['生成工具：run_xhs_rag_workflow','搜索工具：search_topic_material','保存工具：save_content_draft'],'工具少而清晰，优于工具多但使用条件重叠。');
flow(36,'Agent的价值是选择下一步，不是调用所有工具',['检查：目标和信息','追问：缺什么补什么','搜索：明确需要才用','生成：调用P2','保存：确认后执行','停止：完成或越权'],'每一次调用都应有当前状态依据。');
cards(37,'Agent指令必须写清使用条件和停止条件',['信息：生成前必须具备什么','不搜索：素材完整时不要重复查','要搜索：明确需要近期资料','要生成：调用P2而非随便写','要保存：展示并确认后调用','要停止：失败、取消或越权'],'没有发布工具，就不能通过语言假装已经发布。');
demo(38,'接入P2与三个工具，配置Agent','P3_小红书内容策划Agent',['接入P2 Workflow','配置搜索与模拟保存','复制Agent系统指令','运行A1—A5'],['为什么选择这个工具','实际参数传了什么','工具是否真实成功','何时确认与停止'],39);
cards(39,'A1＋A2：完整就生成，缺信息就先问',['A1完整输入：不搜索，直接调用P2','A2只有主题：追问受众、目标和事实','验收证据：看真实工具调用记录'],'一篇看起来不错的文案，不代表缺信息测试通过。');
flow(40,'A3：只有明确需要近期资料时才搜索',['用户：要求近期讨论','搜索：返回材料与来源','检查：区分可靠与未确认','生成：把材料传给P2'],'Agent不会天然联网；只有配置了工具并正确调用才具备搜索能力。');
cards(41,'A4＋A5：保存前确认，没有发布就停止',['A4保存：复述对象→用户确认→调用保存→返回真实draft_id','A5发布：说明没有发布工具→不冒充执行→给出允许替代动作'],'已生成、已保存、已审核和已发布是四种不同状态。');
cards(42,'Agent增加动态选择，也增加新的风险',['新增能力：按状态选择下一步','新增能力：按需调用流程与工具','新增风险：选错工具或参数','新增风险：无效调用和越权','新增成本：更复杂的测试与治理'],'Agent不是默认终点；动态价值必须覆盖新增成本。');

cards(43,'同一组任务同时测试P0—P3',['T1：完整事实生成','T2：个人语气对照','T3：历史事实隔离','T4：缺必要输入','T5：要求近期材料','T6：要求直接发布'],'同题比较，才能分清流程、语气和动态工具分别解决什么。');
flow(44,'四层能力不是替代关系，而是组合关系',['Agent：选择下一步','语气RAG Workflow：风格＋固定流程','LLM节点：各自使用Prompt'],'事实来自本次输入；语气来自风格卡；Agent按需调用它们。');
cards(45,'简单版本能解决，就不要为了Agent而Agent',['P0：单次、输入完整、只需生成','P1：路径固定、需要稳定和排错','P2：固定流程＋指定可更新知识','P3：下一步确实需要动态判断'],'选择标准是任务特征、风险和成本，不是名词新旧。');
cards(46,'保存四个版本，并迁移到自己的场景',['Prompt：决定节点如何理解和输出','Workflow：把固定步骤变成可重复流程','语气RAG：检索样本并抽象表达方式','Agent：根据状态选择下一步和工具','提交：项目＋正常测试＋边界测试＋对照表'],'最重要的产物，是一套知道数据从哪里来、怎样隔离、怎样验证的方法。');

if(pptx._slides.length!==46) throw new Error(`Expected 46 slides, got ${pptx._slides.length}`);
pptx.writeFile({fileName:OUT});
