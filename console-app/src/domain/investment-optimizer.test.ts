import test from 'node:test';
import assert from 'node:assert/strict';
import { dates, buildBaseline, defaultScenario, project, type Snapshot } from './simulation';
import { compareOptimizationWindows, constantBudgetScenario, historicalBase, interventionBases, investmentGrowth, optimizationSettings, optimizeWindow, validateOptimization } from './investment-optimizer';
const calendar = dates('2026-09-22','2026-10-21');
const snapshot: Snapshot = { start: calendar[0], end: calendar.at(-1)!, loadedAt: '', warnings: [], targets: [{ month:'2026-09-01',tax_pct:8 }], revenue:calendar.map(date=>({date,revenue:2000,paid_orders:100})), bureau:calendar.map(date=>({date,custo_bureau:200})), ads:calendar.map(date=>({date,campaign_id:'a',cost_micros:1000e6})) };
const base=buildBaseline(snapshot), settings={...optimizationSettings(defaultScenario(base)),controlBudget:1000};
test('stationary point, engine projections, and chart points agree on net profit including extra costs',()=>{
 const p={...settings,extraFeePct:2,fixedDay:50},r=optimizeWindow(base,snapshot.targets,p);
 const expected=base.adsDay*Math.pow(.8*base.ordersDay*(20*.9-2)/base.adsDay,5);
 assert.ok(Math.abs(r.steady.budget-expected)<12);
 const check=project(base,snapshot.targets,r.steady.scenario).scenario;assert.equal(r.steady.profit,check.profit);assert.equal(r.steady.projection.fixed,4900);assert.ok(r.steady.projection.fees>0);
 assert.equal(r.curve.find(c=>c.budget===r.steady.budget)!.steady,r.steady.profit);
 for(const c of r.curve)assert.ok(c.steady<=r.steady.profit+.5);
});
test('one source snapshot yields 7/14/30-day bases ending on the same date with their own weighted units',()=>{
 const s=structuredClone(snapshot);s.revenue.at(-1)!.revenue=4000;s.bureau.at(-1)!.custo_bureau=400;
 const r=compareOptimizationWindows(s,settings);assert.deepEqual(r.map(x=>x.base.days),[7,14,30]);assert.ok(r.every(x=>x.end===snapshot.end));assert.ok(r[0].base.ticket>r[2].base.ticket);assert.equal(r[0].start,'2026-10-15');
});
test('keeping avoids learning; the transition curve does not silently give an exemption at the control budget',()=>{
 const p={...settings,controlBudget:3900,learningPct:100,learningDays:98},r=optimizeWindow(base,snapshot.targets,p);
 assert.equal(r.learning.status,'keep');assert.equal(r.learning.budget,3900);assert.equal(r.learning.profit,r.control.profit);
 assert.ok(r.changedPeak.profit<r.control.profit);assert.ok(r.curve.find(c=>c.budget===3900)!.learning<r.control.profit);
});
test('learning occurs once, is clamped to the horizon, and zero-loss/zero-duration modes reconcile',()=>{
 const p={...settings,days:3,learningDays:7},s=constantBudgetScenario(base,p,1200,true);assert.equal(s.cycles.length,1);assert.equal(s.cycles[0].days,3);assert.equal(s.cycles[0].learningPct,6);
 const r=optimizeWindow(base,snapshot.targets,{...settings,learningDays:0});assert.equal(r.steady.profit,r.learning.profit);
 const zero=optimizeWindow(base,snapshot.targets,{...settings,learningPct:0});assert.equal(zero.steady.profit,zero.learning.profit);
});
test('linear and superlinear responses report a boundary instead of inventing an interior optimum',()=>{
 for(const elasticity of [1,1.2]){const r=optimizeWindow(base,snapshot.targets,{...settings,elasticity,maxBudget:5000});assert.equal(r.steady.status,'upper-bound');assert.equal(r.steady.budget,5000);}
});
test('unprofitable unit economics or zero responsive share select the lower bound, including zero',()=>{
 const targets=[{month:'2026-09-01',tax_pct:100}],loss=optimizeWindow(base,targets,settings);assert.equal(loss.steady.budget,0);assert.equal(loss.steady.profit,0);
 const insensitive=optimizeWindow(base,snapshot.targets,{...settings,responsivePct:0});assert.equal(insensitive.steady.budget,0);
});
test('constraints are respected and control outside the search range remains a reference only',()=>{
 const r=optimizeWindow(base,snapshot.targets,{...settings,minBudget:500,maxBudget:800});assert.ok(r.steady.budget>=500&&r.steady.budget<=800);assert.equal(r.steady.budget,800);assert.equal(r.learning.status==='keep',false);assert.equal(r.control.budget,1000);
 assert.ok(validateOptimization({...settings,maxBudget:0}).length);assert.ok(validateOptimization({...settings,learningDays:1.5}).length);assert.ok(validateOptimization({...settings,minBudget:NaN}).length);
});
test('missing older dates block only the affected comparison windows',()=>{
 const partial={...snapshot,start:'2026-10-15',revenue:snapshot.revenue.slice(-7),bureau:snapshot.bureau.slice(-7),ads:snapshot.ads.slice(-7)},r=compareOptimizationWindows(partial,settings);
 assert.ok(r[0].steady);assert.equal(r[1].steady,undefined);assert.ok(r[2].base.issues.length>0);
});
test('future tax changes and sub-100 media share remain consistent with the financial engine',()=>{
 const targets=[{month:'2026-09-01',tax_pct:8},{month:'2026-11-01',tax_pct:12}],p={...settings,responsivePct:60};
 const r=optimizeWindow(base,targets,p);for(const o of [r.steady,r.learning])assert.equal(o.profit,project(base,targets,o.scenario).scenario.profit);
 const b=r.changedPeak.budget, scenario=constantBudgetScenario(base,p,b,true,true);assert.equal(r.curve.find(c=>c.budget===b)!.learning,project(base,targets,scenario).scenario.profit);
});
test('intervention separates the entire transition day, includes every later closed day, and shares one control across five bases',()=>{
 const s=structuredClone(snapshot);s.revenue.find(r=>r.date==='2026-10-16')!.revenue=100;s.revenue.at(-1)!.revenue=4000;
 const i={changeDate:'2026-10-16',recentDays:2},b=interventionBases(s,i),r=compareOptimizationWindows(s,settings,i);
 assert.equal(b.before.days,7);assert.equal(b.transition.revenue,100);assert.equal(b.post.days,5);assert.equal(b.post.revenue,12000);assert.equal(b.recent.days,2);assert.equal(b.recent.revenue,6000);
 assert.deepEqual(r.map(x=>x.window),['post-change','recent',7,14,30]);assert.equal(r[0].start,'2026-10-17');assert.equal(r[1].start,'2026-10-20');assert.ok(r.every(x=>x.control?.budget===1000));
 // The actual transition loss remains in the ordinary 7-day financial history.
 assert.equal(r[2].base.revenue,14100);
});
test('recent window never reaches before the intervention and same-day interventions cannot invent complete post-change data',()=>{
 const b=interventionBases(snapshot,{changeDate:'2026-10-20',recentDays:7});assert.equal(b.recentStart,'2026-10-21');assert.equal(b.recent.days,1);
 const r=compareOptimizationWindows(snapshot,settings,{changeDate:'2026-10-21',recentDays:2});assert.equal(r[0].steady,undefined);assert.equal(r[1].steady,undefined);assert.ok(r[2].steady);
 assert.throws(()=>interventionBases(snapshot,{changeDate:'2026-10-16',recentDays:0}));assert.throws(()=>interventionBases(snapshot,{changeDate:'2026-02-31',recentDays:2}));
});
test('missing transition data cannot silently contaminate or block an otherwise complete post-change base',()=>{
 const s={...snapshot,revenue:snapshot.revenue.filter(r=>r.date!=='2026-10-16')};
 const r=compareOptimizationWindows(s,settings,{changeDate:'2026-10-16',recentDays:2});assert.ok(r[0].steady);assert.ok(r[1].steady);assert.equal(r[2].steady,undefined);assert.ok(historicalBase(s,'2026-09-01').issues.length);
});
test('growth bridge reconciles all revenue, tax, fees and media to incremental profit without double-counting learning',()=>{
 const p={...settings,extraFeePct:2,fixedDay:35},targets=[{month:'2026-09-01',tax_pct:8},{month:'2026-11-01',tax_pct:12}];
 const r=investmentGrowth(base,targets,p,1400,true),d=r.delta;
 assert.ok(Math.abs(d.revenue-d.ads-d.bureau-d.tax-d.fees-d.fixed-d.profit)<.00001);
 assert.equal(r.scenario.profit,project(base,targets,constantBudgetScenario(base,p,1400,true)).scenario.profit);
 assert.equal(d.ads,39200);assert.equal(d.fixed,0);assert.ok(d.fees>0);assert.ok(r.learningCost>0);assert.equal(r.incrementalReturn,d.profit/39200);
 const keep=investmentGrowth(base,targets,p,p.controlBudget,true);assert.equal(keep.delta.profit,0);assert.equal(keep.learningCost,0);assert.equal(keep.incrementalReturn,null);
 assert.throws(()=>investmentGrowth(base,targets,p,NaN,true));
});
test('growth changes when actual recent efficiency or the hypothetical media response changes',()=>{
 const improved={...base,ordersDay:base.ordersDay*1.3},normal=investmentGrowth(base,snapshot.targets,settings,1500,false),better=investmentGrowth(improved,snapshot.targets,settings,1500,false);
 assert.ok(better.delta.profit>normal.delta.profit);
 const proportional=investmentGrowth(base,snapshot.targets,{...settings,elasticity:1},1500,false);assert.ok(proportional.delta.profit>normal.delta.profit);
});
