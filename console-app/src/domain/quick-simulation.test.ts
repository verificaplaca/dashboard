import test from 'node:test';
import assert from 'node:assert/strict';
import { buildBaseline, type Snapshot } from './simulation';
import { manualCurve, manualProjection, quickCurve, quickProjection, type QuickSettings } from './quick-simulation';
const s:Snapshot={start:'2026-10-28',end:'2026-10-28',loadedAt:'2026-10-29T12:00:00Z',warnings:[],targets:[{month:'2026-10-01',tax_pct:8}],revenue:[{date:'2026-10-28',revenue:1600,paid_orders:100}],bureau:[{date:'2026-10-28',custo_bureau:200}],ads:[{date:'2026-10-28',campaign_id:'a',cost_micros:1000e6}]};
const base=buildBaseline(s),settings:QuickSettings={start:'2026-10-29',days:30,reference:1000,learning:false,feeRate:0,fixedDay:0};
test('profitable business can lose incremental profit with diminishing efficiency while constant CAC keeps growing',()=>{
 assert.equal(base.profit,272);
 const same=quickProjection(base,s.targets,settings,5000,1),moderate=quickProjection(base,s.targets,settings,5000,.8);
 assert.ok(same.delta>0);assert.ok(moderate.delta<0);assert.ok(moderate.proposed.profit<0);assert.ok(moderate.marginalCac!>moderate.unitContribution);
});
test('keeping the reference avoids new learning and equal budgets reconcile exactly',()=>{
 const steady=quickProjection(base,s.targets,settings,1000,.8),learning=quickProjection(base,s.targets,{...settings,learning:true},1000,.8);
 assert.deepEqual(steady,learning);assert.equal(learning.delta,0);
 const changed=quickProjection(base,s.targets,{...settings,learning:true},1500,.8),without=quickProjection(base,s.targets,settings,1500,.8);
 assert.ok(changed.proposed.profit<without.proposed.profit);assert.equal(changed.current.profit,without.current.profit);assert.equal(changed.end,'2026-11-27');
});
test('DRE fees scale with revenue and fixed costs apply to both alternatives once, including rates over 100%',()=>{
 const naked=quickProjection(base,s.targets,settings,1500,.8),costed=quickProjection(base,s.targets,{...settings,feeRate:.05,fixedDay:100},1500,.8);
 assert.equal(costed.proposed.fees,Math.round(naked.proposed.revenue*.05*100)/100);assert.equal(costed.proposed.fixed,3000);assert.equal(costed.current.fixed,3000);
 assert.ok(Math.abs(costed.proposed.profit-(naked.proposed.profit-costed.proposed.fees-3000))<.011);
 assert.ok(quickProjection(base,s.targets,{...settings,feeRate:2},1500,.8).proposed.profit<0);
});
test('the marked point beats the grid or keeping and the upper bound is identified without claiming an unlimited optimum',()=>{
 const c=quickCurve(base,s.targets,settings);assert.ok(c.best.budget>=0&&c.best.budget<=10000);assert.ok(c.series[1].points.every(p=>p.y<=c.best.profit+.005));assert.ok(c.best.profit>=quickProjection(base,s.targets,settings,settings.reference,.8).proposed.profit);
 const efficient={...base,ads:100,adsDay:100};const bounded=quickCurve(efficient,s.targets,settings);assert.equal(bounded.bounded,true);assert.equal(bounded.best.budget,10000);
});
test('future tax changes and learning duration affect results, and invalid choices never return projections',()=>{
 const higher=quickProjection(base,[...s.targets,{month:'2026-11-01',tax_pct:15}],settings,1500,.8);assert.ok(higher.proposed.profit<quickProjection(base,s.targets,settings,1500,.8).proposed.profit);
 for(const budget of [NaN,-1,10001])assert.throws(()=>quickProjection(base,s.targets,settings,budget,.8));
 for(const bad of [{...settings,days:0},{...settings,reference:0},{...settings,feeRate:-1},{...settings,fixedDay:Infinity}])assert.throws(()=>quickProjection(base,s.targets,bad,1000,.8));
});
test('manual calculator starts at historical CAC and doubles sales and variable amounts with spend',()=>{
 const unchanged=manualProjection(base,s.targets,{...settings,learning:true},1000);
 assert.equal(unchanged.delta,0);
 const doubled=manualProjection(base,s.targets,{...settings,learning:true},2000);
 assert.equal(doubled.proposed.orders,6000);
 assert.equal(doubled.proposed.cpa,10);
 for(const key of ['orders','revenue','bureau','tax','profit'] as const)assert.equal(doubled.proposed[key],unchanged.proposed[key]*2);
 assert.deepEqual(doubled.current,unchanged.current);
});
test('choosing a worse CAC reduces sales and profit without changing the historical control',()=>{
 const same=manualProjection(base,s.targets,settings,2000);
 const worse=manualProjection(base,s.targets,{...settings,orderRate:1/12},2000);
 assert.ok(Math.abs(worse.proposed.orders-2000*30/12)<1e-6);
 assert.ok(Math.abs(worse.proposed.cpa!-12)<1e-6);
 assert.ok(worse.proposed.profit<same.proposed.profit);
 assert.deepEqual(worse.current,same.current);
});
test('chosen CAC remains constant when moving budget again or changing ticket and bureau',()=>{
 const chosen={...settings,learning:true,orderRate:1/8,ticket:20,bureauUnit:3};
 const low=manualProjection(base,s.targets,chosen,1600),high=manualProjection(base,s.targets,chosen,3200);
 assert.equal(low.proposed.cpa,8);assert.equal(high.proposed.cpa,8);
 assert.equal(high.proposed.orders,low.proposed.orders*2);
 assert.equal(high.proposed.revenue,high.proposed.orders*20);
 assert.equal(high.proposed.bureau,high.proposed.orders*3);
 assert.deepEqual(high.current,manualProjection(base,s.targets,settings,3200).current);
});
test('manual fees follow revenue and fixed costs stay fixed, including a zero-budget scenario',()=>{
 const chosen={...settings,feeRate:.05,fixedDay:100};
 const one=manualProjection(base,s.targets,chosen,1000),two=manualProjection(base,s.targets,chosen,2000);
 assert.equal(two.proposed.fees,one.proposed.fees*2);assert.equal(two.proposed.fixed,one.proposed.fixed);
 const zero=manualProjection(base,s.targets,chosen,0);
 assert.equal(zero.proposed.orders,0);assert.equal(zero.proposed.revenue,0);assert.equal(zero.proposed.bureau,0);assert.equal(zero.proposed.profit,-3000);assert.equal(zero.proposed.cpa,null);
});
test('manual chart contains one chosen-CAC curve with no extra efficiency scenarios',()=>{
 const curve=manualCurve(base,s.targets,{...settings,learning:true});
 assert.equal(curve.series.length,1);
 for(const budget of [1000,2000,5000])assert.equal(curve.series[0].points.find(p=>p.x===budget)!.y,manualProjection(base,s.targets,settings,budget).proposed.profit);
});
test('manual invalid efficiency never silently falls back to historical CAC',()=>{
 for(const orderRate of [NaN,-1,Infinity])assert.throws(()=>manualProjection(base,s.targets,{...settings,orderRate},2000));
});
