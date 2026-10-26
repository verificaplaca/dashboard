import test from 'node:test';
import assert from 'node:assert/strict';
import { loadSimulationData } from './simulation-data';
Object.defineProperty(globalThis, 'location', { value: { href: 'http://localhost/' }, configurable: true });
const valid = (table: string) => table === 'revenue_daily' ? [{ date: '2026-09-01', revenue: '100', paid_orders: 10 }] : table === 'bureau_daily' ? [{ date: '2026-09-01', custo_bureau: '20' }] : table === 'monthly_targets' ? [{ month: '2026-09-01', tax_pct: null }] : [{ date: '2026-09-01', campaign_id: '1', cost_micros: '30000000' }];
test('loader pages Ads and uses only GET to original dashboard sources', async t => {
  let calls = 0;
  t.mock.method(globalThis, 'fetch', async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input)), table = url.pathname.split('/').at(-1)!;
    assert.equal(url.hostname, 'ftmgmfdqdqxboiktxcoj.supabase.co'); assert.equal(init?.method || 'GET', 'GET');
    if (table === 'google_ads_campaign_daily') { calls++; return new Response(JSON.stringify(calls === 1 ? Array.from({ length: 1000 }, (_, i) => ({ ...valid(table)[0], campaign_id: String(i) })) : [])); }
    return new Response(JSON.stringify(valid(table)));
  });
  const data = await loadSimulationData('2026-09-01', '2026-09-01'); assert.equal(calls, 2); assert.equal(data.ads.length, 1000); assert.equal(data.revenue[0].revenue, 100); assert.ok(data.warnings.some(w => w.includes('8%')));
});
test('required source failure rejects instead of inventing zero costs or demo data', async t => {
  t.mock.method(globalThis, 'fetch', async input => String(input).includes('/bureau_daily?') ? new Response('', { status: 503 }) : new Response(JSON.stringify(valid(new URL(String(input)).pathname.split('/').at(-1)!))));
  await assert.rejects(loadSimulationData('2026-09-01', '2026-09-01'), /bureau_daily/);
});
test('optional targets failure keeps explicit 8 percent fallback', async t => {
  t.mock.method(globalThis, 'fetch', async input => String(input).includes('/monthly_targets?') ? new Response('', { status: 503 }) : new Response(JSON.stringify(valid(new URL(String(input)).pathname.split('/').at(-1)!))));
  const data = await loadSimulationData('2026-09-01', '2026-09-01'); assert.deepEqual(data.targets, []); assert.ok(data.warnings.some(w => w.includes('Metas indisponíveis')));
});
test('missing numeric costs are rejected and current-day baseline cannot be loaded', async t => {
  t.mock.method(globalThis, 'fetch', async input => new Response(JSON.stringify(String(input).includes('/bureau_daily?') ? [{ date: '2026-09-01', custo_bureau: null }] : valid(new URL(String(input)).pathname.split('/').at(-1)!))));
  await assert.rejects(loadSimulationData('2026-09-01', '2026-09-01'), /custo_bureau/);
  await assert.rejects(loadSimulationData('2099-01-01', '2099-01-01'), /dias encerrados/);
});
