import test from 'node:test';
import assert from 'node:assert/strict';
import { loadFinanceData } from './finance-data';
Object.defineProperty(globalThis,'location',{value:{href:'http://localhost/'},configurable:true});
const fixture=(table:string)=>({revenue_daily:[{date:'2026-09-01',revenue:'100',paid_orders:10}],bureau_daily:[{date:'2026-09-01',custo_bureau:'20'}],google_ads_campaign_daily:[{date:'2026-09-01',campaign_id:'1',cost_micros:'30000000'}],monthly_targets:[{month:'2026-09-01',tax_pct:8}],refunds_daily:[{date:'2026-09-01',refund_count:'1',refund_value:'15'}]}[table]??[]);
test('DRE reads the original sources and informative refunds using only GET',async t=>{
 const seen=new Set<string>();t.mock.method(globalThis,'fetch',async(input,init)=>{const table=new URL(String(input)).pathname.split('/').at(-1)!;seen.add(table);assert.equal(init?.method??'GET','GET');return new Response(JSON.stringify(fixture(table)));});
 const data=await loadFinanceData('2026-09-01','2026-09-01');assert.equal(seen.size,5);assert.deepEqual(data.refunds,[{date:'2026-09-01',refund_count:1,refund_value:15}]);assert.equal(data.refundWarning,'');assert.equal(data.revenue[0].revenue,100);
});
test('unavailable refunds remain unknown and do not become zero or stop the core DRE',async t=>{
 t.mock.method(globalThis,'fetch',async input=>{const table=new URL(String(input)).pathname.split('/').at(-1)!;return table==='refunds_daily'?new Response('',{status:503}):new Response(JSON.stringify(fixture(table)));});
 const data=await loadFinanceData('2026-09-01','2026-09-01');assert.equal(data.refunds,null);assert.match(data.refundWarning,/indisponíveis/);assert.equal(data.ads[0].cost_micros,30000000);
});
test('DRE refuses broken required sources rather than understating costs',async t=>{
 t.mock.method(globalThis,'fetch',async input=>{const table=new URL(String(input)).pathname.split('/').at(-1)!;return table==='bureau_daily'?new Response('',{status:503}):new Response(JSON.stringify(fixture(table)));});
 await assert.rejects(loadFinanceData('2026-09-01','2026-09-01'),/bureau_daily/);
});
test('malformed refund amounts are rejected and an incomplete current day is excluded',async t=>{
 t.mock.method(globalThis,'fetch',async input=>{const table=new URL(String(input)).pathname.split('/').at(-1)!;return new Response(JSON.stringify(table==='refunds_daily'?[{date:'2026-09-01',refund_count:1,refund_value:null}]:fixture(table)));});
 await assert.rejects(loadFinanceData('2026-09-01','2026-09-01'),/estornos/);
 await assert.rejects(loadFinanceData('2099-01-01','2099-01-01'),/dias encerrados/);
});
