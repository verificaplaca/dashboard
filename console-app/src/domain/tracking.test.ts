import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { channelStatus, dailyDispatches, dispatchLatency, emptyTrackingFilters, eventIdCheck, filterTracking, matchIdsDetail, matchQuality, orderIsFull, orderHasFailure, readTrackingOrders, trackingCac, trackingSummary, trackingWindows, type TrackingDispatch } from './tracking';

const now = Date.parse('2026-10-31T15:00:00Z');
const row = (overrides: Partial<TrackingDispatch> = {}): TrackingDispatch => ({ checkout_id: 'checkout-1', order_nsu: 'order_1', event_id: 'evt_order_1', value: 79.9, currency: 'BRL', ga4_status: 'success', ads_status: 'success', meta_status: 'success', attempts: 1, created_at: '2026-10-30T12:00:00Z', paid_at: '2026-10-30T11:59:40Z', ...overrides });
const plain = (value: unknown) => JSON.parse(JSON.stringify(value));
function legacy() {
  const html = fs.readFileSync(new URL('../../../tracking-gateway.html', import.meta.url), 'utf8');
  const script = html.match(/<script>([\s\S]*?)<\/script>/)![1].replace(/\/\/ ─── BOOTSTRAP[\s\S]*/, '');
  const elements = new Map<string, any>();
  const element = (id: string) => { if (!elements.has(id)) elements.set(id, { innerHTML: '', textContent: '', style: {}, classList: { add() {}, remove() {}, toggle() {} }, addEventListener() {}, value: '' }); return elements.get(id); };
  class Clock extends Date { static now() { return now; } }
  const context = vm.createContext({ console, Date: Clock, setTimeout, clearTimeout, requestAnimationFrame() {}, URLSearchParams, navigator: {}, window: { innerWidth: 1200, addEventListener() {}, parent: { postMessage() {} } }, document: { getElementById: element, querySelector: element, querySelectorAll: () => [], addEventListener() {}, body: {}, documentElement: {} }, ResizeObserver: class { observe() {} } });
  vm.runInContext(script, context);
  return { context, element, run: (code: string) => vm.runInContext(code, context) };
}

test('diagnostics match the published previous gateway across identifiers, missing data and latency boundaries', () => {
  const old = legacy();
  for (const value of [row(), row({ order_nsu: 'addon_2', event_id: 'evt_checkout-1', has_email: true, has_phone: true, gbraid: 'gb', fbc: 'fb', fbp: 'browser' }), row({ event_id: 'unknown', paid_at: null }), row({ order_nsu: null, paid_at: '2026-10-29T12:00:00Z' }), row({ paid_at: '2026-10-30T13:00:00Z' }), row({ paid_at: '2026-10-30T11:00:00Z' }), row({ ga4_status: 'failed', ads_status: 'pending', meta_status: 'skipped' })]) {
    old.context.sample = value;
    assert.equal(orderIsFull(value), old.run('orderIsFull(sample)'));
    assert.equal(orderHasFailure(value), old.run('orderHasFailure(sample)'));
    assert.equal(matchQuality(value), old.run('matchQuality(sample).label'));
    assert.deepEqual(matchIdsDetail(value).present, plain(old.run('matchIdsDetail(sample).presentes')));
    assert.equal(eventIdCheck(value).label, old.run('eventIdCheck(sample).label'));
    assert.equal(dispatchLatency(value)?.label ?? null, old.run('dispatchLatency(sample)?.label ?? null'));
  }
});

test('health and attribution use the entire rolling seven-day window, not the visible history page', () => {
  const rows = Array.from({ length: 1100 }, (_, i) => row({ checkout_id: `c${i}`, has_email: true, gclid: i % 2 ? 'g' : null, fbp: 'f', attempts: 2 }));
  rows.push(row({ checkout_id: 'skipped', meta_status: 'skipped' }), row({ checkout_id: 'failed', ads_status: 'failed' }), row({ checkout_id: 'pending', ga4_status: 'pending' }), row({ checkout_id: 'older', created_at: '2026-10-20T00:00:00Z' }));
  const result = trackingSummary(rows, now);
  assert.equal(result.total, 1103); assert.equal(result.full, 1101); assert.equal(result.failing, 2);
  assert.equal(result.channels[2].skipped, 1); assert.equal(result.attribution[0].count, 550);
  const old = legacy(); old.context.samples = rows; old.run('STATE.orders = samples; renderHealth(); renderChannels(); renderAttribution();');
  assert.match(old.element('healthRow').innerHTML, /1\.103/);
  assert.match(old.element('healthRow').innerHTML, /1\.101/);
  assert.match(old.element('attrRow').innerHTML, /550 de 1\.103/);
});

test('empty and skipped sources never become confirmed sends or a fabricated quality percentage', () => {
  assert.equal(trackingSummary([], now).quality, null);
  assert.equal(trackingSummary([], now).attempts, null);
  const result = trackingSummary([row({ ga4_status: 'skipped', ads_status: 'skipped', meta_status: 'skipped' })], now);
  assert.equal(result.full, 1); assert.equal(result.channels[0].success, 0); assert.equal(result.channels[0].skipped, 1);
  assert.equal(channelStatus(row({ ads_status: 'skipped' }), 'ads'), 'skipped');
});

test('filters search beyond the first page and combine status, channel, attribution and BRL amounts', () => {
  const rows = [row(), row({ checkout_id: 'late', event_id: 'target', value: 1799.9, ads_status: 'failed', gbraid: 'click' }), row({ checkout_id: 'skip', meta_status: 'skipped', has_phone: true, has_email: true })];
  assert.equal(filterTracking(rows, { ...emptyTrackingFilters, search: '1.799,90' })[0].checkout_id, 'late');
  assert.equal(filterTracking(rows, { ...emptyTrackingFilters, search: '1799.90' }).length, 0);
  assert.equal(filterTracking(rows, { ...emptyTrackingFilters, search: '1799.9' }).length, 1);
  assert.equal(filterTracking(rows, { ...emptyTrackingFilters, channel: 'ads', attribution: 'com_g', min: '100', status: 'failed' })[0].checkout_id, 'late');
  assert.equal(filterTracking(rows, { ...emptyTrackingFilters, status: 'partial' }).length, 1);
  assert.equal(filterTracking(rows, { ...emptyTrackingFilters, status: 'full' }).length, 2);
});

test('financial comparison matches previous weighted CAC and signed daily gaps, including fractional Ads conversions', () => {
  const revenue = [{ date: '2026-10-29', revenue: 1500, paid_orders: 100 }, { date: '2026-10-30', revenue: 300, paid_orders: 20 }];
  const ads = [{ date: '2026-10-29', cost_micros: '500000000', conversions: '40.5' }, { date: '2026-10-29', cost_micros: 500000000, conversions: 39.5 }, { date: '2026-10-30', cost_micros: 100000000, conversions: 25 }];
  const result = trackingCac(revenue, ads);
  assert.equal(result.realCac, 1100 / 120); assert.equal(result.reportedCac, 1100 / 105);
  assert.ok(result.gap! < 0); assert.ok(result.daily[0].gap! > 0);
  assert.equal(result.daily[1].realCac, 10); assert.equal(result.daily[1].reportedCac, 12.5);
  const old = legacy(); old.context.revenue = revenue; old.context.ads = ads; old.run('renderCac(revenue, ads)');
  assert.match(old.element('cacSummaryRow').innerHTML, /9,17/); assert.match(old.element('cacSummaryRow').innerHTML, /10,48/);
  assert.match(old.element('#cacTable tbody').innerHTML, /-25,0%/);
  assert.equal(trackingCac([], []).realCac, null); assert.equal(trackingCac([], []).gap, null);
});

test('daily graph keeps 7/14/30/90 day buckets and counts each successful channel separately', () => {
  for (const days of [7, 14, 30, 90]) {
    const result = dailyDispatches([row(), row({ checkout_id: 'failed', ga4_status: 'failed', meta_status: 'skipped' })], days, new Date(now));
    assert.equal(result.length, days);
    const bucket = result.find(day => day.date === '2026-10-30')!;
    assert.deepEqual(bucket, { date: '2026-10-30', ga4: 1, ads: 2, meta: 1, fail: 1 });
  }
});

test('90-day reader traverses API pages with a fixed cutoff and stable ordering', async () => {
  const rows = Array.from({ length: 1637 }, (_, i) => row({ checkout_id: `c${i}` })), calls: URLSearchParams[] = [];
  const range = trackingWindows(new Date(now));
  const result = await readTrackingOrders(async query => { calls.push(query); const start = Number(query.get('offset')); return rows.slice(start, start + Number(query.get('limit'))); }, range, new AbortController().signal);
  assert.equal(result.length, rows.length); assert.deepEqual(result.map(r => r.checkout_id), rows.map(r => r.checkout_id));
  assert.equal(calls.length, 5);
  calls.forEach(query => { assert.equal(query.get('order'), 'created_at.desc,checkout_id.asc'); assert.match(query.get('and')!, /created_at.lte.2026-10-31T15:00:00.000Z/); });
});

test('an incomplete or cancelled read is rejected instead of publishing partial health totals', async () => {
  let calls = 0;
  await assert.rejects(readTrackingOrders(async () => { calls++; if (calls > 1) throw new Error('unavailable'); return Array.from({ length: 500 }, () => row()); }, trackingWindows(new Date(now)), new AbortController().signal), /unavailable/);
  const abort = new AbortController();
  await assert.rejects(readTrackingOrders(async () => { abort.abort(); return [row()]; }, trackingWindows(new Date(now)), abort.signal));
});
