import test from 'node:test';
import assert from 'node:assert/strict';
import { buildBaseline, type Snapshot } from './simulation';
import { quickCurve, quickProjection, type QuickSettings } from './quick-simulation';
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
