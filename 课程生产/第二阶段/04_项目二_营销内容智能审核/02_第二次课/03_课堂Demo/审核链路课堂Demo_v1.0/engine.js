/* Pure, deterministic classroom checks. This is not a semantic model. */
(function(root,factory){const x=factory();if(typeof module==='object')module.exports=x;root.ReviewEngine=x;})(typeof globalThis!=='undefined'?globalThis:this,function(){
'use strict';
const SAMPLE={
 initial:'养猫家庭的日常清扫，交给净小伴P1标准版。最高12000Pa吸力，65℃热水直接拖地。本次活动领券到手1799元，人人同价！活动期内下单就送宠物清洁包。',
 partial:'净小伴P1标准版，最高12000Pa吸力，基站支持65℃热水洗拖布。本次活动新会员领券到手1799元。活动期内购买P1标准版并支付成功，按活动系统确认的支付顺序，前200名每单赠一份宠物清洁包，赠完即止。',
 complete:'净小伴P1标准版，最高12000Pa吸力，基站支持65℃热水洗拖布。活动期内，领取并使用400元活动券后1999元；新会员再领取并使用200元专享券后1799元。活动期内购买P1标准版并支付成功，按活动系统确认的支付顺序，前200名每单赠一份宠物清洁包，赠完即止。清扫前请收好电线和袜子。',
 boundary:'净小伴P1标准版，65℃深度清洁，宠物家庭也能轻松打理。本次活动价格请以适用活动方案为准。',
 feature:'净小伴P1标准版，最高12000Pa吸力，基站支持65℃热水洗拖布。最后再说一下它的自动清扫功能。',
 updated:'净小伴P1标准版，最高12000Pa吸力，基站支持65℃热水洗拖布。活动期内，领取并使用400元活动券后1999元；新会员再领取并使用200元专享券后1799元。宠物清洁包赠品名额已用完，本次推广不再承诺赠送。已符合赠品资格的订单，按原约定履行。'
};
const RULES=[
{id:'R01',version:1,name:'温度与功能对象',enabled:true,source:'F-PROD-01',instruction:'出现温度表达时，核对温度数值与功能对象。65℃对应基站洗拖布，不是热水拖地。对象不清交人工；未提温度不强制补写。'},
{id:'R02',version:1,name:'价格与资格条件',enabled:true,source:'F-PROMO-01',instruction:'结合全文核对优惠路径、人群资格与领用条件。1799元需要新会员资格、领用400元活动券及200元专享券；1999元需要领用400元活动券。条件可在同篇文案其他位置，不因句式不同直接认定缺失。'},
{id:'R03',version:1,name:'赠品承诺与当前权益',enabled:true,source:'F-GIFT-01',instruction:'核对活动、购买商品、成功支付、限量顺序、每单数量与赠完即止。已确认名额用完后，当前推广不应继续承诺赠送；此前已获资格订单按原约定履行。'}];
const SOURCES=[
{id:'F-PROD-01',version:1,title:'P1标准版商品事实',sku:'P1标准版',activity:'*',from:'2026-10-01',to:'2026-12-31',enabled:true,owner:'产品负责人',text:'净小伴P1标准版，标称最高12000Pa吸力；65℃热水用于基站清洗拖布，不是热水直接拖地。清扫前请收好电线、袜子等小件。'},
{id:'F-PROMO-01',version:1,title:'P1十月活动价格方案',sku:'P1标准版',activity:'P1十月活动',from:'2026-10-01',to:'2026-10-31',enabled:true,owner:'活动负责人',text:'本方案仅适用于P1十月活动中的P1标准版。原价2399元，领取并使用400元活动券后1999元；新会员再领取并使用200元新会员专享券后1799元。会员身份及两类券的领取、使用条件需清晰说明，资格与用券条件结合全文判断。'},
{id:'F-GIFT-01',version:1,title:'P1活动赠品方案',sku:'P1标准版',activity:'P1十月活动',from:'2026-10-01',to:'2026-10-31',enabled:true,owner:'活动负责人',text:'活动期内购买本次活动P1标准版并支付成功，按活动系统确认的支付顺序，前200名每单赠一份宠物清洁包，赠完即止。名额是否用完以活动负责人确认事件为准，不能由原始限量数字推断。'}];
function clone(x){return JSON.parse(JSON.stringify(x));}
function defaultConfig(){return {rules:clone(RULES),sources:clone(SOURCES),giftEnded:false,eventVersion:0};}
function contextKey(task,cfg){return JSON.stringify({sku:task.sku,activity:task.activity,asOf:task.asOf,rules:cfg.rules,sources:cfg.sources,giftEnded:cfg.giftEnded,eventVersion:cfg.eventVersion});}
function snapshot(task,cfg){const t=clone(task),c=clone(cfg);return {task:t,config:c,key:contextKey(t,c),createdAt:new Date().toISOString()};}
function matchSources(snap){const t=snap.task;return snap.config.sources.filter(s=>s.enabled&&s.sku===t.sku&&(s.activity==='*'||s.activity===t.activity)&&s.from<=t.asOf&&s.to>=t.asOf);}
function mk(id,status,reason,quote,source,advice){return {ruleId:id,status,reason,quote:quote||'',sourceId:source?.id||'',sourceVersion:source?.version||null,evidence:source?.text||'',advice:advice||'',origin:'本地规则演示'};}
function excerpt(text,regex){const m=text.match(regex);if(!m)return '';const at=m.index,start=Math.max(0,text.lastIndexOf('。',at-1)+1);let end=text.indexOf('。',at);if(end<0)end=text.length;return text.slice(start,Math.min(end+1,text.length));}
function simulate(snap,fault='none'){
 const t=snap.task,text=t.text,cfg=snap.config,sources=matchSources(snap),out=[];
 for(const r of cfg.rules){
  if(!r.enabled){out.push(mk(r.id,'incomplete','该审核点已停用，首版必查范围尚未完成。'));continue;}
  const source=sources.find(s=>s.id===r.source);
  if(!source||(fault==='missing'&&r.id==='R02')){out.push(mk(r.id,'incomplete','未取得与商品、活动、适用时间匹配的有效资料。'));continue;}
  if(fault==='timeout'&&r.id==='R02'){out.push(mk(r.id,'incomplete','模拟调用超时。本项未完成，可重试或人工补核。',null,source));continue;}
  if(r.id==='R01'){
   if(!/(℃|°C|摄氏|温度)/i.test(text)){out.push(mk(r.id,'na','未出现温度主张，本项不适用。','',source));continue;}
   const q=excerpt(text,/(℃|°C|摄氏|温度)/i);
   if(/(?:热水|65\s*[℃°]).{0,8}(?:直接拖地|热水拖地|拖洗地面)/.test(text))out.push(mk(r.id,'issue','文案将温度用于拖地，给定资料对应基站清洗拖布。',q,source,'明确为基站65℃热水洗拖布，并复核完整文案。'));
   else if(/基站.{0,12}65\s*℃.{0,12}(洗拖布|清洗拖布)/.test(text)&&!(text.match(/(?:℃|°C)/g)||[]).some((_,i)=>i>0))out.push(mk(r.id,'supported','命中课堂已覆盖的“基站65℃洗拖布”表达；仍需人工确认语义。',q,source));
   else out.push(mk(r.id,'uncertain','本地规则无法可靠确认温度对象或多处主张，请结合全文人工判断。',q,source,'确认功能对象，必要时补充表达。'));
  }
  if(r.id==='R02'){
   if(!/(\d{3,5}\s*元|到手|领券|同价)/.test(text)){out.push(mk(r.id,'na','未出现本课价格主张，本项不适用。','',source));continue;}
   const q=excerpt(text,/(1799|1999|到手|领券|同价)/)||excerpt(text,/\d{3,5}\s*元/);
   const newMember=/新会员/.test(text),four=/400\s*元.{0,5}活动券|活动券.{0,5}400\s*元/.test(text),two=/200\s*元.{0,8}(?:专享券|新会员券)|(?:专享券|新会员券).{0,5}200\s*元/.test(text),use=/领取并使用|领取.{0,8}使用|领用/.test(text);
   if(/人人同价|所有人.{0,4}1799|无门槛.{0,5}1799/.test(text))out.push(mk(r.id,'issue','1799元需要限定的新会员与两类券路径，“人人同价”等扩大了适用人群。',q,source,'删除扩大范围的承诺，明确新会员资格及两类券的领取和使用条件。'));
   else if(/1799/.test(text)&&!(newMember&&four&&two&&use)){
    const miss=[!newMember&&'新会员资格',!four&&'400元活动券',!two&&'200元专享券',!use&&'领取与使用条件'].filter(Boolean);
    out.push(mk(r.id,'issue','课堂规则未识别到：'+miss.join('、')+'。自由改写可能产生误报，需人工核对全文。',q,source,'补齐缺失条件，或记录文案中已经表达这些条件的位置。'));
   }else if(/1999/.test(text)&&!(four&&use))out.push(mk(r.id,'issue','1999元未识别到400元活动券的领取与使用条件。',q,source,'明确领用400元活动券后1999元。'));
   else if(/1799|1999/.test(text))out.push(mk(r.id,'supported','已识别本课价格与条件组合，金额路径经程序核算。语义与适用性仍需人工确认。',q,source));
   else out.push(mk(r.id,'uncertain','该价格表达不在本地规则可靠覆盖范围，请人工核对，不计为无问题。',q,source));
  }
  if(r.id==='R03'){
   if(!/赠|礼包|清洁包/.test(text)){out.push(mk(r.id,'na','未出现赠品主张，本项不适用。','',source));continue;}
   const q=excerpt(text,/(赠|礼包|清洁包)/);
   if(cfg.giftEnded&&/名额已用完|不再承诺|停止赠送|赠品已结束/.test(text))out.push(mk(r.id,'supported','已识别名额用完说明，需内容负责人核验实际使用版本。',q,source));
   else if(cfg.giftEnded)out.push(mk(r.id,'issue','E-GIFT-01已确认名额用完，当前文案仍含赠送承诺。',q,{...source,text:source.text+'\nE-GIFT-01：活动负责人确认名额已用完，自教学事件登记后适用；已取得资格的订单权益保留。'},'确认当前传播用途，修改赠品承诺并核验实际使用位置。'));
   else if(/前\s*200\s*名/.test(text)&&/支付成功/.test(text)&&/支付顺序/.test(text)&&/每单.{0,3}一份/.test(text)&&/赠完即止/.test(text)&&/活动期/.test(text)&&/P1/.test(text))out.push(mk(r.id,'supported','已识别本课赠品条件组合，仍需人工确认适用范围。',q,source));
   else out.push(mk(r.id,'issue','赠品的购买、支付顺序、限量、数量或有效范围条件未完整识别。',q,source,'按有效活动方案说明条件；不得由“前200名”直接推断当前剩余名额。'));
  }
 }
 if(fault==='quote'){const item=out.find(x=>x.status==='issue'||x.status==='supported');if(item)item.quote='原文里不存在的1799元承诺';}
 const calculation={base:2399,activityCoupon:400,memberCoupon:200,afterActivity:2399-400,afterMember:2399-400-200};
 return {checks:out,raw:{checks:clone(out)},preparation:{calculation},engine:'simulation',model:null};
}
function validate(snap,result){
 const allowed=matchSources(snap),checks=[],problems=[];
 for(const rule of snap.config.rules){
  const found=(result.checks||[]).filter(x=>x.ruleId===rule.id);let item=found.length===1?clone(found[0]):mk(rule.id,'incomplete',found.length?'返回重复审核点，不能直接采用。':'模型未返回该审核点，本项未完成。');
  const errors=[];if(!['issue','supported','uncertain','na','incomplete'].includes(item.status))errors.push('结论枚举无效');
  if(!rule.enabled){item=mk(rule.id,'incomplete','停用审核点不能算作已完成。');}
  const src=allowed.find(x=>x.id===item.sourceId&&x.version===item.sourceVersion);
  if(!['incomplete'].includes(item.status)&&!src)errors.push('引用资料不在本次有效输入中');
  if(['issue','supported','uncertain'].includes(item.status)){
   if(typeof item.quote!=='string'||!item.quote||!snap.task.text.includes(item.quote))errors.push('原句不属于本次正文');
   if(!item.evidence||!src||!src.text.includes(item.evidence)){
    // Event evidence is explicitly part of the frozen input bundle.
    const event=snap.config.giftEnded?'\nE-GIFT-01：活动负责人确认名额已用完，自教学事件登记后适用；已取得资格的订单权益保留。':'';
    if(!src||!(src.text+event).includes(item.evidence||'')||!item.evidence)errors.push('依据引文未通过来源校验');
   }
  }
  if(errors.length){problems.push({ruleId:rule.id,errors});item={...item,status:'incomplete',reason:'结果校验未通过：'+errors.join('；'),invalid:true};}
  item.position=item.quote&&!item.invalid?snap.task.text.indexOf(item.quote):-1;item.validated=!errors.length;checks.push(item);
 }
 const unknown=(result.checks||[]).filter(x=>!snap.config.rules.some(r=>r.id===x.ruleId));if(unknown.length)problems.push({ruleId:'unknown',errors:['存在未请求的审核点，已排除']});
 return {...result,checks,validation:{problems,passed:checks.filter(x=>x.validated).length,total:checks.length,note:'程序校验来源与结构，不能证明语义判断正确。'}};
}
function runnable(run,task,cfg){return !!run&&run.snapshot.task.version===task.version&&run.snapshot.task.text===task.text&&run.snapshot.key===contextKey(task,cfg);}
function canApprove(run,task,cfg,decisions={}){
 if(!runnable(run,task,cfg))return '正文、上下文、资料或审核点已变化，请针对提交版本重新检查。';
 for(const c of run.result.checks){
  const d=decisions[c.ruleId];
  if(c.status==='issue'&&d?.action!=='dismiss')return '仍有未解决或未判断的问题，请退回修改或记录不成立理由。';
  if(['uncertain','incomplete'].includes(c.status)&&d?.action!=='manual')return '存在未完成或待确认项，请重试或记录人工补核。';
 }
 return '';
}
return {SAMPLE,RULES,SOURCES,clone,defaultConfig,contextKey,snapshot,matchSources,simulate,validate,runnable,canApprove};
});
