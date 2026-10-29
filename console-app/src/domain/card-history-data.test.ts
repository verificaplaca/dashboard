import test from 'node:test';
import assert from 'node:assert/strict';
import { loadCardHistory } from './card-history';
Object.defineProperty(globalThis,'location',{value:{href:'http://localhost/'},configurable:true});
const fixture=(table:string)=>({revenue_daily:[{date:'2026-10-28',revenue:1600,paid_orders:100}],bureau_daily:[{date:'2026-10-28',custo_bureau:200}],google_ads_campaign_daily:[{date:'2026-10-28',campaign_id:'a',cost_micros:1000e6}],monthly_targets:[{month:'2026-01-01',tax_pct:8}],upsell_daily:[{date:'2026-10-28',upsell_orders:20,upsell_revenue:100}],checkouts_daily:[{date:'2026-10-28',checkouts_started:1000}]}[table]??[]);
test('card context uses only original aggregate GET sources, with checkouts on the original site project',async t=>{
 const seen:string[]=[];t.mock.method(globalThis,'fetch',async(input,init)=>{const url=new URL(String(input)),table=url.pathname.split('/').at(-1)!;seen.push(table);assert.equal(init?.method??'GET','GET');assert.equal(url.hostname,table==='checkouts_daily'?'ozquoloetuzynnyzkado.supabase.co':'ftmgmfdqdqxboiktxcoj.supabase.co');return new Response(JSON.stringify(fixture(table)));});
 const data=await loadCardHistory('2026-10-28','2026-10-28');assert.equal(seen.length,6);assert.equal(data.global?.value,400);assert.equal(data.global?.start,'2026-01-01');assert.equal(data.upsell?.[0].revenue,100);assert.equal(data.checkouts?.[0].count,1000);
});
test('broken checkout or upsell context stays unknown, without invalidating independent historical totals',async t=>{
 t.mock.method(globalThis,'fetch',async input=>{const table=new URL(String(input)).pathname.split('/').at(-1)!;return table==='checkouts_daily'?new Response('',{status:503}):new Response(JSON.stringify(table==='upsell_daily'?[{date:'2026-10-28',upsell_orders:20,upsell_revenue:null}]:fixture(table)));});
 const data=await loadCardHistory('2026-10-28','2026-10-28');assert.equal(data.checkouts,null);assert.equal(data.upsell,null);assert.equal(data.global?.value,400);assert.equal(data.warnings.length,2);
});
test('required global cost failures never create a fake EBITDA and cancelled context loads are aborted',async t=>{
 t.mock.method(globalThis,'fetch',async input=>{const table=new URL(String(input)).pathname.split('/').at(-1)!;return table==='bureau_daily'?new Response('',{status:503}):new Response(JSON.stringify(fixture(table)));});
 const data=await loadCardHistory('2026-10-28','2026-10-28');assert.equal(data.global,null);assert.equal(data.upsell?.[0].orders,20);
 const c=new AbortController();c.abort();await assert.rejects(loadCardHistory('2026-10-28','2026-10-28',c.signal),{name:'AbortError'});
});
