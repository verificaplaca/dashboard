import test from 'node:test';
import assert from 'node:assert/strict';
import { buildDre, emptyExpenses, expensePayload, monthPeriod, type FinanceSnapshot } from './finance';
const s:FinanceSnapshot={start:'2026-10-01',end:'2026-10-01',loadedAt:'2026-10-02T12:00:00Z',warnings:[],targets:[{month:'2026-09-01',tax_pct:8}],revenue:[{date:'2026-10-01',revenue:1000,paid_orders:60}],bureau:[{date:'2026-10-01',custo_bureau:100}],ads:[{date:'2026-10-01',campaign_id:'a',cost_micros:200e6}],refunds:[{date:'2026-10-01',refund_count:5,refund_value:500}],refundWarning:''};
const zero={fees:0,people:0,systems:0,other:0,financial:0};
test('DRE reconciles each subtotal and never subtracts already-excluded canceled orders again',()=>{
 const d=buildDre(s,zero);assert.deepEqual([d.sales,d.tax,d.netRevenue,d.consultations,d.grossProfit,d.media,d.panelProfit,d.profit,d.refunds],[1000,80,920,100,820,200,620,620,500]);assert.equal(d.complete,true);assert.equal(d.margin,62);
 const changed=buildDre({...s,refunds:[{date:s.start,refund_count:10,refund_value:900}]},zero);assert.equal(changed.profit,d.profit);assert.equal(changed.refunds,900);
});
test('blank costs remain pending; explicit zero completes the categories without inventing expenses',()=>{
 const incomplete=buildDre(s);assert.equal(incomplete.complete,false);assert.equal(incomplete.missing.length,5);assert.equal(incomplete.expenses.fees,null);assert.equal(incomplete.profit,620);
 const full=buildDre(s,{fees:50,people:100,systems:0,other:0,financial:20});assert.equal(full.complements,170);assert.equal(full.profit,450);assert.equal(full.complete,true);
 const unknown=buildDre({...s,refunds:null,refundWarning:'Indisponível'},zero);assert.equal(unknown.refunds,null);assert.equal(unknown.profit,620);
});
test('periods exclude today, handle leap years and reject months without closed days',()=>{
 assert.deepEqual(monthPeriod('2026-10','2026-10-29'),{start:'2026-10-01',end:'2026-10-28',calendarDays:31,partial:true});
 assert.equal(monthPeriod('2024-02','2026-10-29').end,'2024-02-29');assert.equal(monthPeriod('2026-09','2026-10-29').calendarDays,30);
 for(const month of ['2026-13','2026-1','','2026-11'])assert.throws(()=>monthPeriod(month,'2026-10-29'));
 assert.throws(()=>monthPeriod('2026-10','2026-10-01'));
});
test('zero-sales days preserve losses; missing or duplicate financial sources cannot appear complete',()=>{
 const loss=buildDre({...s,revenue:[{date:s.start,revenue:0,paid_orders:0}]},zero);assert.equal(loss.profit,-300);assert.equal(loss.margin,null);assert.equal(loss.complete,true);
 for(const broken of [{...s,bureau:[]},{...s,revenue:[...s.revenue,s.revenue[0]]},{...s,ads:[]},{...s,revenue:[{...s.revenue[0],revenue:NaN}]}]){const d=buildDre(broken,zero);assert.equal(d.complete,false);assert.ok(d.issues.length);}
 assert.equal(buildDre({...s,warnings:['Imposto em fallback']},zero).complete,false);
});
test('expense validation distinguishes null, zero and cents and rejects malformed amounts',()=>{
 assert.deepEqual(expensePayload({}),emptyExpenses());assert.equal(expensePayload({fees:'0',people:'100.25'}).fees,0);assert.equal(expensePayload({fees:'0',people:'100.25'}).people,100.25);
 for(const value of ['-1','NaN','Infinity','1.001','1000000001'])assert.throws(()=>expensePayload({fees:value}));
});
