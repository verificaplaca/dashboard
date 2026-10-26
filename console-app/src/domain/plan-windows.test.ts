import test from 'node:test';
import assert from 'node:assert/strict';
import { bestPlanStrategies, comparePlanWindow, comparePlanWindows } from './gradual-comparison';
import { buildBaseline, dates, project, type Snapshot } from './simulation';
const calendar = dates('2026-09-26', '2026-10-25');
const snapshot: Snapshot = { start: calendar[0], end: calendar.at(-1)!, loadedAt: '2026-10-26T20:11:33Z', warnings: [], targets: [{ month: '2026-09-01', tax_pct: 8 }],
  revenue: calendar.map((date, i) => ({ date, revenue: (100 + i) * 16, paid_orders: 100 + i })),
  bureau: calendar.map((date, i) => ({ date, custo_bureau: (100 + i) * 2 })),
  ads: calendar.map((date, i) => ({ date, campaign_id: 'a', cost_micros: (1000 + i * 20) * 1e6 })) };
test('four exact historical windows share the full horizon, media budgets and same engine for every condition', () => {
  const before = structuredClone(snapshot), windows = comparePlanWindows(snapshot);
  assert.deepEqual(windows.map(w => [w.id, w.start, w.end, w.data?.base.days]), [
    ['post-change', '2026-10-17', '2026-10-25', 9], ['7', '2026-10-19', '2026-10-25', 7], ['14', '2026-10-12', '2026-10-25', 14], ['30', '2026-09-26', '2026-10-25', 30],
  ]);
  for (const w of windows) {
    assert.equal(w.error, ''); const c = w.data!;
    assert.equal(c.days, 98); assert.equal(c.start, '2026-10-26'); assert.equal(c.end, '2027-01-31');
    assert.equal(c.plannedMedia, 389940); assert.equal(c.directMedia, 490000); assert.equal(c.controlMedia, 273420);
    assert.equal(c.rows.length, 4);
    const base = buildBaseline({ ...snapshot, start: w.start, end: w.end });
    assert.deepEqual(c.base, base);
    for (const r of c.rows) {
      assert.deepEqual(r.result, project(base, snapshot.targets, r.scenario));
      assert.deepEqual(r.directResult, project(base, snapshot.targets, r.directScenario));
      assert.equal(r.result.control.profit, r.directResult.control.profit);
    }
  }
  assert.deepEqual(snapshot, before);
});
test('post-change efficiency excludes the change day and earlier data, while broader bases retain it', () => {
  const altered = structuredClone(snapshot), index = altered.revenue.findIndex(r => r.date === '2026-10-16');
  altered.revenue[index].revenue *= 2;
  const a = comparePlanWindows(snapshot), b = comparePlanWindows(altered);
  assert.deepEqual(a[0].data, b[0].data); assert.deepEqual(a[1].data, b[1].data);
  assert.notEqual(a[2].data!.base.revenue, b[2].data!.base.revenue);
  assert.notEqual(a[3].data!.base.revenue, b[3].data!.base.revenue);
});
test('missing data blocks only affected windows; never shortens 30 days or inserts a zero projection', () => {
  const windows = comparePlanWindows({ ...snapshot, bureau: snapshot.bureau.filter(r => r.date !== '2026-10-01') });
  assert.ok(windows.slice(0, 3).every(w => w.data && !w.error));
  assert.equal(windows[3].data, null); assert.ok(windows[3].error.length > 0);
  const allBroken = comparePlanWindows({ ...snapshot, ads: snapshot.ads.filter(r => r.date !== '2026-10-24') });
  assert.ok(allBroken.every(w => w.data === null && w.error.length > 0));
  const limited = comparePlanWindows({ ...snapshot, start: '2026-10-19' });
  assert.ok(limited[1].data); assert.equal(limited[0].data, null); assert.equal(limited[2].data, null); assert.equal(limited[3].data, null);
});
test('invalid learning and unverified taxes block every window, and zero learning reconciles every strategy', () => {
  assert.ok(comparePlanWindows(snapshot, { days: NaN, lossPct: 6 }).every(w => !w.data && w.error.includes('Aprendizado')));
  assert.ok(comparePlanWindows({ ...snapshot, targets: [] }).every(w => !w.data));
  assert.ok(comparePlanWindows({ ...snapshot, warnings: ['Fonte sem validação'] }).every(w => !w.data));
  for (const learning of [{ days: 0, lossPct: 6 }, { days: 7, lossPct: 0 }]) for (const w of comparePlanWindows(snapshot, learning)) {
    const rows = w.data!.rows;
    for (const [a, b] of [[0, 1], [2, 3]]) {
      assert.equal(rows[a].result.scenario.profit, rows[b].result.scenario.profit);
      assert.equal(rows[a].directResult.scenario.profit, rows[b].directResult.scenario.profit);
    }
  }
});
test('the three-strategy ranking includes keeping, ties, reversals and losses', () => {
  const row = structuredClone(comparePlanWindow(snapshot, '2026-10-19', snapshot.end).rows[0]);
  const rank = (keep: number, gradual: number, direct: number) => {
    row.result.control.profit = keep; row.result.scenario.profit = gradual; row.directResult.scenario.profit = direct;
    return bestPlanStrategies(row);
  };
  assert.deepEqual(rank(100, 80, 90), ['keep']);
  assert.deepEqual(rank(80, 100, 90), ['gradual']);
  assert.deepEqual(rank(80, 90, 100), ['direct']);
  assert.deepEqual(rank(100, 100, 90), ['keep', 'gradual']);
  assert.deepEqual(rank(100, 100, 100), ['keep', 'gradual', 'direct']);
  assert.deepEqual(rank(-10, -20, -30), ['keep']);
});
