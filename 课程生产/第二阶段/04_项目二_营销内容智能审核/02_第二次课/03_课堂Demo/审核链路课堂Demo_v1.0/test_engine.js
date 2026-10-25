const assert=require('node:assert/strict'),E=require('./engine.js');let count=0;
function run(text,config=E.defaultConfig(),fault='none'){const snapshot=E.snapshot({docId:'D01',version:1,title:'test',sku:'P1标准版',activity:'P1十月活动',asOf:'2026-10-21',text},config);return {snapshot,result:E.validate(snapshot,E.simulate(snapshot,fault))};}
function test(name,fn){fn();console.log('PASS',name);count++;}
test('problem draft yields three actionable findings',()=>assert.deepEqual(run(E.SAMPLE.initial).result.checks.map(x=>x.status),['issue','issue','issue']));
test('complete draft has three supported points',()=>assert.deepEqual(run(E.SAMPLE.complete).result.checks.map(x=>x.status),['supported','supported','supported']));
test('member-only edit still misses coupon conditions',()=>assert.equal(run(E.SAMPLE.partial).result.checks.find(x=>x.ruleId==='R02').status,'issue'));
test('boundary temperature is uncertain and feature has no gift requirement',()=>{assert.equal(run(E.SAMPLE.boundary).result.checks[0].status,'uncertain');assert.equal(run(E.SAMPLE.feature).result.checks[2].status,'na');});
test('missing activity materials cannot become pass',()=>assert.equal(run(E.SAMPLE.complete,E.defaultConfig(),'missing').result.checks[1].status,'incomplete'));
test('fabricated quotation fails validation',()=>assert.equal(run(E.SAMPLE.initial,E.defaultConfig(),'quote').result.checks[0].status,'incomplete'));
test('stale version and updated sources invalidate reuse',()=>{const r=run(E.SAMPLE.complete),t={...r.snapshot.task,version:2};assert.equal(E.runnable(r,t,E.defaultConfig()),false);const cfg=E.defaultConfig();cfg.sources[1].version++;assert.equal(E.runnable(r,r.snapshot.task,cfg),false);});
test('disabled required rule is incomplete',()=>{const c=E.defaultConfig();c.rules[1].enabled=false;assert.equal(run(E.SAMPLE.complete,c).result.checks[1].status,'incomplete');});
test('gift change flags active old promise and retains qualified rights text',()=>{const c=E.defaultConfig();c.giftEnded=true;c.eventVersion=1;assert.equal(run(E.SAMPLE.complete,c).result.checks[2].status,'issue');assert.equal(run(E.SAMPLE.updated,c).result.checks[2].status,'supported');});
test('formal approval blocked by unresolved findings and failures',()=>{let r=run(E.SAMPLE.initial);assert.ok(E.canApprove(r,r.snapshot.task,E.defaultConfig()));r=run(E.SAMPLE.complete);assert.equal(E.canApprove(r,r.snapshot.task,E.defaultConfig()),'');r=run(E.SAMPLE.complete,E.defaultConfig(),'timeout');assert.ok(E.canApprove(r,r.snapshot.task,E.defaultConfig()));});
test('out-of-scope evidence rejected',()=>{const r=run(E.SAMPLE.complete),raw=E.simulate(r.snapshot);raw.checks[1].sourceId='OTHER';assert.equal(E.validate(r.snapshot,raw).checks[1].status,'incomplete');});
test('different activity yields unavailable activity materials',()=>{const r=run(E.SAMPLE.complete);r.snapshot.task.activity='another';const x=E.validate(r.snapshot,E.simulate(r.snapshot));assert.equal(x.checks[1].status,'incomplete');});
console.log(count+' meaningful checks passed.');
