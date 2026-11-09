const requirements=[
 {id:'R-001',src:'r1',type:'资格',text:'具备有效期内的信息安全管理体系认证，并提供证书复印件',fact:'现有证书已过有效期',state:'gap',label:'存在差异',action:'资质补办 / 决定是否参标'},
 {id:'R-002',src:'r2',type:'资格',text:'近三年不少于三项政务行业同类项目案例',fact:'2 项政务案例 + 1 项零售案例',state:'partial',label:'部分满足',action:'确认案例口径 / 寻找联合体'},
 {id:'R-003',src:'r3',type:'技术',text:'支持政务内网私有化部署，并发用户数不低于 5,000',fact:'支持私有部署；实测并发 10,000',state:'satisfied',label:'已满足',action:'引用测试报告与部署方案'},
 {id:'R-004',src:'r4',type:'交付',text:'合同签订后 60 个自然日内完成部署、联调和初验',fact:'标准交付周期 90 天',state:'gap',label:'存在差异',action:'压缩范围 / 增加资源 / 谈判'},
 {id:'R-005',src:'r5',type:'技术',text:'兼容国产操作系统及数据库，并提供适配说明',fact:'已确认麒麟 V10、达梦 8、人大金仓',state:'satisfied',label:'已满足',action:'按已确认产品逐项响应'},
 {id:'R-006',src:'r6',type:'服务',text:'提供 7×24 小时技术支持可得 3 分',fact:'标准服务为 5×8',state:'pending',label:'待确认',action:'管理层确认服务升级成本'}
];
const risks=[
 {level:'高',title:'资质承诺与现状冲突',desc:'草稿若写“具备有效认证”将构成无依据承诺。应删除并进入参标决策。'},
 {level:'高',title:'交付周期无法由现有计划支撑',desc:'招标要求 60 天，标准交付为 90 天。需要新的资源与范围方案。'},
 {level:'中',title:'同类案例数量口径不一致',desc:'零售案例不能直接计入政务行业案例，需人工确认评分口径。'},
 {level:'中',title:'服务章节存在跨章节不一致',desc:'章节 5.1 写 7×24，报价与人员计划仍按 5×8 测算。'},
 {level:'中',title:'国产兼容范围可能被过度泛化',desc:'只能承诺已验证的麒麟 V10、达梦 8、人大金仓，不能写“全面兼容”。'},
 {level:'低',title:'并发能力缺少材料引用编号',desc:'10,000 并发结论应关联性能测试报告的版本与页码。'}
];
let parsed=false,drafted=false,checked=false,resolved=0;
const views=[...document.querySelectorAll('.view')],navs=[...document.querySelectorAll('.nav-item')];
const titles={overview:'项目总览',parse:'要求解析',matrix:'响应矩阵',editor:'标书编辑',risk:'风险检查'};
function showView(id){views.forEach(v=>v.classList.toggle('active',v.id===id));navs.forEach(n=>n.classList.toggle('active',n.dataset.view===id));document.querySelector('#viewTitle').textContent=titles[id];const labels={overview:parsed?'查看要求':'开始解析',parse:parsed?'查看矩阵':'开始解析',matrix:'生成章节',editor:drafted?'检查风险':'生成草稿',risk:checked?'查看结果':'开始检查'};document.querySelector('#nextStep').textContent=labels[id]}
navs.forEach(n=>n.addEventListener('click',()=>showView(n.dataset.view)));
function toast(msg){const el=document.querySelector('#toast');el.textContent=msg;el.classList.add('show');setTimeout(()=>el.classList.remove('show'),2100)}
function renderReqs(){const wrap=document.querySelector('#requirementCards');wrap.innerHTML=requirements.map(r=>`<article class="req-card" data-src="${r.src}"><div class="req-top"><strong>${r.id} · ${r.type}</strong><span class="status ${r.state}">${r.label}</span></div><p>${r.text}</p><small>来源：招标文件第 42 页 · 点击定位原文</small></article>`).join('');wrap.querySelectorAll('.req-card').forEach(c=>c.addEventListener('click',()=>{wrap.querySelectorAll('.req-card').forEach(x=>x.classList.remove('active'));c.classList.add('active');document.querySelectorAll('.source-line').forEach(x=>x.classList.toggle('highlight',x.dataset.req===c.dataset.src))}))}
function renderMatrix(){document.querySelector('#matrixBody').innerHTML=requirements.map(r=>`<tr><td><b>${r.id}</b></td><td>${r.text}</td><td>${r.type}</td><td>${r.fact}</td><td><span class="status ${r.state}">${r.label}</span></td><td>${r.action}<br><button class="action-btn">查看依据</button></td></tr>`).join('')}
function setFlow(n){document.querySelectorAll('.flow-step').forEach((s,i)=>{s.classList.toggle('done',i<n);s.classList.toggle('active',i===n);s.querySelector('small').textContent=i<n?'已完成':i===n?'进行中':'等待'});document.querySelector('#progressText').textContent=`${Math.min(n,5)} / 5`}
function parse(){if(parsed){showView('parse');return}const buttons=[document.querySelector('#nextStep'),document.querySelector('#parseNow')].filter(Boolean);buttons.forEach(b=>{b.disabled=true;b.textContent='正在拆解…'});setFlow(1);setTimeout(()=>{parsed=true;document.querySelector('#parseEmpty').classList.add('hidden');document.querySelector('#requirementCards').classList.remove('hidden');renderReqs();renderMatrix();document.querySelector('#metricReq').textContent='6';document.querySelector('#metricReady').textContent='2';document.querySelector('#metricPending').textContent='2';document.querySelector('#metricRisk').textContent='2';document.querySelector('#projectStatus').textContent='要求已解析';setFlow(2);buttons.forEach(b=>{b.disabled=false;b.textContent='查看矩阵'});toast('已拆解为 6 条最小要求');showView('parse')},900)}
document.querySelector('#parseFromCard').addEventListener('click',parse);document.querySelector('#parseNow').addEventListener('click',parse);
function generate(){if(!parsed){parse();return}const btn=document.querySelector('#generateDraft');btn.disabled=true;btn.textContent='生成中…';setTimeout(()=>{drafted=true;document.querySelector('#draftText').className='draft-placeholder generated';document.querySelector('#draftText').innerHTML=`<p>启明数科平台支持在客户政务内网环境中进行私有化部署，应用、模型服务与数据存储均可按安全域独立部署。<sup>[材料 01]</sup></p><p>当前已完成对 <strong>麒麟 V10、达梦数据库 8、人大金仓数据库</strong> 的适配验证，并可提供相应兼容性清单与部署说明。<sup>[材料 02]</sup></p><p><mark>待人工确认：</mark>其他国产软硬件环境需在项目启动后依据实际版本进行适配评估，本方案不作未经验证的兼容承诺。</p>`;btn.disabled=false;btn.textContent='重新生成';document.querySelector('#projectStatus').textContent='章节生成中';setFlow(4);toast('已生成受控草稿，并保留待确认项')},850)}
document.querySelector('#generateDraft').addEventListener('click',generate);
function runCheck(){if(!drafted){showView('editor');toast('请先生成一段受控草稿');return}const btn=document.querySelector('#runCheck');btn.disabled=true;btn.textContent='检查中…';setTimeout(()=>{checked=true;document.querySelector('#riskEmpty').classList.add('hidden');document.querySelector('#riskResults').classList.remove('hidden');document.querySelector('#riskList').innerHTML=risks.map((r,i)=>`<article class="risk-item ${r.level==='高'?'high':''}" data-i="${i}"><div class="risk-level">${r.level}</div><div><h4>${r.title}</h4><p>${r.desc}</p></div><button class="action-btn resolve">标记已处理</button></article>`).join('');document.querySelectorAll('.resolve').forEach(b=>b.addEventListener('click',()=>{const item=b.closest('.risk-item');if(!item.classList.contains('resolved')){item.classList.add('resolved');b.textContent='已处理';resolved++;document.querySelector('#resolvedCount').textContent=resolved;toast('处理状态已更新')}}));btn.disabled=false;btn.textContent='重新检查';document.querySelector('#projectStatus').textContent='等待人工审核';setFlow(5);toast('检查完成：发现 6 个问题')},950)}
document.querySelector('#runCheck').addEventListener('click',runCheck);
document.querySelector('#nextStep').addEventListener('click',()=>{const active=document.querySelector('.view.active').id;if(active==='overview'||active==='parse')parse();else if(active==='matrix')showView('editor');else if(active==='editor')generate();else runCheck()});
document.querySelector('#resetDemo').addEventListener('click',()=>location.reload());
renderMatrix();showView('overview');
