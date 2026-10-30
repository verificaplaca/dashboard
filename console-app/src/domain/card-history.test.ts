import test from 'node:test';
import assert from 'node:assert/strict';
import { applyCardEdit, cardRates, cardSparks, financialCardValues, resolveCardEdits, type CardHistory, type CardEdits } from './card-history';
import { buildBaseline, type Snapshot } from './simulation';
import { quickCurve, quickProjection, responsesFor } from './quick-simulation';
const s:Snapshot={start:'2026-10-27',end:'2026-10-28',loadedAt:'2026-10-29T12:00:00Z',warnings:[],targets:[{month:'2026-10-01',tax_pct:8}],revenue:['2026-10-27','2026-10-28'].map(date=>({date,revenue:1600,paid_orders:100})),bureau:['2026-10-27','2026-10-28'].map(date=>({date,custo_bureau:200})),ads:['2026-10-27','2026-10-28'].map(date=>({date,campaign_id:'a',cost_micros:1000e6}))};
const h:CardHistory={upsell:s.revenue.map(r=>({date:r.date,orders:20,revenue:100})),checkouts:s.revenue.map(r=>({date:r.date,count:1000})),global:{value:800,start:'2026-01-01',end:s.end,daily:[400,400]},warnings:[]};
const b=buildBaseline(s),rates=cardRates(s,b,h,s.start,s.end),settings={start:'2026-10-29',days:30,reference:1000,learning:false,feeRate:.03,fixedDay:10};
const project=(edits:CardEdits={},budget=1500)=>quickProjection(b,s.targets,{...settings,...resolveCardEdits(b,rates,edits)},budget,.8);
const near=(a:number,b:number,tolerance=.31)=>assert.ok(Math.abs(a-b)<=tolerance,`${a} vs ${b}`);
test('cloned financial definitions reconcile with the original panel while net includes known DRE expenses',()=>{
 const p=project().proposed,v=financialCardValues(p,rates,{},h.global);
 assert.equal(v.ct,p.ads+p.bureau);near(v.p,p.revenue-v.ct,.011);assert.equal(v.ro,p.revenue/v.ct);assert.equal(v.cpv,v.ct/p.orders);assert.equal(v.ca,p.ads/p.orders);assert.equal(v.m,v.p/p.revenue*100);assert.equal(v.ml,p.profit/p.revenue*100);assert.equal(v.up,20);assert.equal(v.eb,800);
 assert.equal(v.npo,p.profit);near(v.npo,v.p-p.tax-p.fees-p.fixed,.011);
});
test('price edits keep volume, costs and the historical control, but flow through revenue, tax, fees and profit',()=>{
 const old=project(),edits=applyCardEdit('tk',20,'',b,rates,{},old.proposed,30).edits,p=project(edits);
 assert.equal(p.proposed.orders,old.proposed.orders);assert.equal(p.proposed.bureau,old.proposed.bureau);assert.deepEqual(p.current,old.current);near(p.proposed.revenue-old.proposed.revenue,p.proposed.orders*4);assert.ok(p.proposed.tax>old.proposed.tax);assert.ok(p.proposed.fees>old.proposed.fees);assert.ok(p.delta>old.delta);
});
test('revenue inversion has an explicit choice: price changes or sales efficiency changes, never both silently',()=>{
 const old=project(),target=old.proposed.revenue*1.2;
 const byPrice=project(applyCardEdit('r',target,'ticket',b,rates,{},old.proposed,30).edits);
 const byVolumeEdits=applyCardEdit('r',target,'volume',b,rates,{},old.proposed,30).edits,byVolume=project(byVolumeEdits);
 assert.equal(byPrice.proposed.orders,old.proposed.orders);assert.ok(byVolume.proposed.orders>old.proposed.orders);near(byPrice.proposed.revenue,target);near(byVolume.proposed.revenue,target);near(byVolume.proposed.bureau/byVolume.proposed.orders,b.bureauUnit,.001);assert.deepEqual(byVolume.current,old.current);
});
test('manual CAC defines one final efficiency without stacking elasticity or an invented learning penalty',()=>{
 const old=project(),e=applyCardEdit('ca',8,'',b,rates,{},old.proposed,30).edits;
 const st={...settings,learning:true,...resolveCardEdits(b,rates,e)},p=quickProjection(b,s.targets,st,1500,.8),other=quickProjection(b,s.targets,st,1500,.6);
 near(p.proposed.cpa!,8,.0001);assert.equal(p.proposed.orders,1500*30/8);assert.deepEqual(p.proposed,other.proposed);assert.equal(responsesFor(st).length,1);assert.equal(quickCurve(b,s.targets,st).series.length,1);assert.deepEqual(p.current,old.current);
 const metrics=financialCardValues(p.proposed,rates,e,null);near(metrics.cv!,rates.conversionPct!,.0001);assert.ok(metrics.cv!<=100);
 const double=project(e,3000);assert.equal(double.proposed.orders,p.proposed.orders*2);
});
test('orders and checkout conversion edit the same acquisition driver, including an explicit zero-sales loss',()=>{
 const old=project(),e=applyCardEdit('tr',5000,'',b,rates,{},old.proposed,30).edits,p=project(e);near(p.proposed.orders,5000,.0001);
 const c=project(applyCardEdit('cv',20,'',b,rates,{},old.proposed,30).edits),v=financialCardValues(c.proposed,rates,{},null);near(v.cv!,20,.0001);
 const zero=project(applyCardEdit('tr',0,'',b,rates,{},old.proposed,30).edits);assert.equal(zero.proposed.orders,0);assert.equal(zero.proposed.revenue,0);assert.equal(financialCardValues(zero.proposed,rates,{},null).ml,null);assert.ok(zero.proposed.profit<0);assert.equal(financialCardValues(zero.proposed,rates,{},null).ca,null);
});
test('cost edits retain Ads or sales as disclosed and preserve costs when changing upsell or price again',()=>{
 const old=project();assert.throws(()=>applyCardEdit('ct',old.proposed.ads-1,'',b,rates,{},old.proposed,30),/abaixo/);
 const e=applyCardEdit('gb',old.proposed.bureau*1.2,'',b,rates,{},old.proposed,30).edits,p=project(e);near(p.proposed.bureau,old.proposed.bureau*1.2);assert.equal(p.proposed.ads,old.proposed.ads);assert.equal(p.proposed.orders,old.proposed.orders);
 const mix=applyCardEdit('up',40,'',b,rates,{},old.proposed,30,{unitRevenue:5,unitBureau:2}).edits,m=project(mix);
 near(m.proposed.revenue-old.proposed.revenue,m.proposed.orders);near(m.proposed.bureau-old.proposed.bureau,m.proposed.orders*.4);
 const repriced=project(applyCardEdit('tk',20,'',b,rates,mix,m.proposed,30).edits);assert.equal(repriced.proposed.bureau,m.proposed.bureau);
 const recosted=project(applyCardEdit('gb',m.proposed.bureau+100,'',b,rates,mix,m.proposed,30).edits);assert.equal(recosted.proposed.revenue,m.proposed.revenue);
 assert.throws(()=>applyCardEdit('up',40,'',b,rates,{},old.proposed,30),/Custo desconhecido/);
 assert.throws(()=>applyCardEdit('up',40,'',b,rates,{},old.proposed,30,{unitRevenue:5,unitBureau:20}),/incompatível/);
});
test('missing secondary sources remain unknown and sparklines retain losses and gaps',()=>{
 const partial=cardRates(s,b,{...h,checkouts:h.checkouts!.slice(1),upsell:null},s.start,s.end);assert.equal(partial.checkoutPerReal,null);assert.equal(partial.upsellPct,null);assert.equal(financialCardValues({...project().proposed,orders:project().proposed.ads*rates.checkoutPerReal!*2},rates,{},null).cv,null);assert.throws(()=>applyCardEdit('cv',20,'',b,partial,{},project().proposed,30),/checkouts/);
 const loss={...s,revenue:[{...s.revenue[0],revenue:0,paid_orders:0},s.revenue[1]]};const sparks=cardSparks(loss,null,s.start,s.end);assert.equal(sparks.p[0],-1200);assert.equal(sparks.cv[0],null);assert.equal(financialCardValues(project().proposed,partial,{},null).eb,null);
 const gap=cardSparks({...s,bureau:s.bureau.slice(1)},h,s.start,s.end);assert.equal(gap.p[0],null);
});
test('invalid edits fail rather than emitting projections or changing the observed snapshot',()=>{
 const before=JSON.stringify(s),p=project().proposed;
 for(const n of [NaN,Infinity,-1])assert.throws(()=>applyCardEdit('tk',n,'',b,rates,{},p,30));
 assert.throws(()=>applyCardEdit('ca',0,'',b,rates,{},p,30));assert.throws(()=>applyCardEdit('ga',300001,'',b,rates,{},p,30));assert.throws(()=>applyCardEdit('cv',101,'',b,rates,{},p,30));
 assert.equal(applyCardEdit('ga',150000,'',b,rates,{},p,30).budget,5000);assert.equal(JSON.stringify(s),before);
});
