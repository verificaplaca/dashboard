import test from 'node:test';
import assert from 'node:assert/strict';
import { compareGradualPlan } from './gradual-comparison';
import { addDays, project, type Snapshot } from './simulation';
const snapshot: Snapshot = { start: '2026-10-21', end: '2026-10-22', loadedAt: '2026-10-23T14:14:53.301Z', warnings: [], targets: [{ month: '2026-09-01', tax_pct: 8 }], revenue: [{ date: '2026-10-21', revenue: 5748.61, paid_orders: 349 }, { date: '2026-10-22', revenue: 5026.67, paid_orders: 323 }], bureau: [{ date: '2026-10-21', custo_bureau: 934.43 }, { date: '2026-10-22', custo_bureau: 722.64 }], ads: [{ date: '2026-10-21', campaign_id: 'a', cost_micros: 3184.05e6 }, { date: '2026-10-22', campaign_id: 'a', cost_micros: 2345.05e6 }] };
test('the four gradual-plan outcomes exactly reproduce the real-data analysis, control and learning costs', () => {
  const before = structuredClone(snapshot), c = compareGradualPlan(snapshot);
  assert.equal(c.days, 98); assert.equal(c.start, '2026-10-26'); assert.equal(c.end, '2027-01-31'); assert.equal(c.controlBudget, 2790);
  assert.equal(c.plannedMedia, 389940); assert.equal(c.controlMedia, 273420); assert.equal(c.extraMedia, 116520);
  assert.deepEqual(c.rows.map(r => [r.result.control.profit, r.result.scenario.profit, r.result.delta, r.learningCost]), [
    [134857.8, 192327.92, 57470.12, 0], [134857.8, 177401.54, 42543.74, 14926.38], [134110.06, 149966.3, 15856.24, 0], [134110.06, 136109.24, 1999.18, 13857.06],
  ]);
  assert.deepEqual(snapshot, before);
  for (const row of c.rows) {
    assert.equal(row.result.scenario.profit, project(c.base, snapshot.targets, row.scenario).scenario.profit);
    assert.ok(Math.abs(row.result.scenario.profit - row.result.control.profit - row.result.delta) < .000001);
  }
});
test('learning is charged once per each of six increases, for exactly seven days, never for the initial budget or control', () => {
  const c = compareGradualPlan(snapshot); assert.equal(c.affectedDays, 42);
  for (const r of c.rows.filter(r => r.withLearning)) {
    const cycles = r.scenario.cycles, losses = cycles.filter(row => row.learningPct > 0);
    assert.equal(losses.length, 6); assert.ok(losses.every(row => row.days === 7)); assert.equal(cycles[0].budget, 2790); assert.equal(cycles[0].days, 6); assert.equal(cycles[0].learningPct, 0);
    assert.equal(cycles.at(-1)!.days, 15); assert.equal(addDays(r.scenario.start, cycles.reduce((n, row) => n + row.days, 0) - 1), '2027-01-31');
    assert.equal(r.result.control.profit, c.rows.find(row => row.elasticity === r.elasticity && !row.withLearning)!.result.control.profit);
    assert.ok(r.result.scenario.cpa! > c.rows.find(row => row.elasticity === r.elasticity && !row.withLearning)!.result.scenario.cpa!);
  }
});
test('zero duration or zero loss reconcile exactly, and a long transition is limited to each stage', () => {
  for (const settings of [{ days: 0, lossPct: 6 }, { days: 7, lossPct: 0 }]) {
    const r = compareGradualPlan(snapshot, settings); assert.equal(r.affectedDays, 0);
    assert.equal(r.rows[0].result.scenario.profit, r.rows[1].result.scenario.profit); assert.equal(r.rows[2].result.scenario.profit, r.rows[3].result.scenario.profit);
  }
  const long = compareGradualPlan(snapshot, { days: 30, lossPct: 6 }); assert.equal(long.affectedDays, 92); assert.equal(long.rows[1].scenario.cycles.at(-1)!.days, 22);
});
test('invalid learning, missing recent sources and pre-change samples never produce projections', () => {
  for (const settings of [{ days: NaN, lossPct: 6 }, { days: 1.5, lossPct: 6 }, { days: -1, lossPct: 6 }, { days: 31, lossPct: 6 }, { days: 7, lossPct: 101 }, { days: 7, lossPct: NaN }]) assert.throws(() => compareGradualPlan(snapshot, settings));
  for (const broken of [{ ...snapshot, ads: snapshot.ads.slice(1) }, { ...snapshot, bureau: [] }, { ...snapshot, revenue: [...snapshot.revenue, snapshot.revenue[0]] }, { ...snapshot, warnings: ['Imposto em fallback'] }, { ...snapshot, targets: [] }, { ...snapshot, end: '2026-10-17' }]) assert.throws(() => compareGradualPlan(broken));
});
test('only the most recent two complete days set efficiency, while actual future tax settings affect both alternatives', () => {
  const older = { ...snapshot, start: '2026-10-20', revenue: [{ date: '2026-10-20', revenue: 99999, paid_orders: 2000 }, ...snapshot.revenue], bureau: [{ date: '2026-10-20', custo_bureau: 1000 }, ...snapshot.bureau], ads: [{ date: '2026-10-20', campaign_id: 'a', cost_micros: 900e6 }, ...snapshot.ads] };
  const normal = compareGradualPlan(snapshot); assert.deepEqual(compareGradualPlan(older).rows, normal.rows);
  const raised = compareGradualPlan({ ...snapshot, targets: [...snapshot.targets, { month: '2027-01-01', tax_pct: 15 }] });
  assert.ok(raised.rows.every((r, i) => r.result.control.profit < normal.rows[i].result.control.profit && r.result.scenario.profit < normal.rows[i].result.scenario.profit));
});
test('severe learning preserves losses and negative gains instead of clamping the charts at zero', () => {
  const c = compareGradualPlan(snapshot, { days: 30, lossPct: 100 });
  assert.ok(c.rows.filter(r => r.withLearning).every(r => r.result.scenario.profit < 0 && r.result.delta < 0));
  assert.ok(c.rows.filter(r => !r.withLearning).every(r => r.result.scenario.profit > 0));
});
test('direct-to-5000 outcomes reconcile with the earlier analysis and the same control', () => {
  const c = compareGradualPlan(snapshot);
  assert.equal(c.directMedia, 490000); assert.equal(c.extraDirectMedia, 100060); assert.equal(c.directBudget, 5000);
  assert.deepEqual(c.rows.map(r => [r.directResult.scenario.profit, r.directAdvantage, r.directLearningCost, r.winner]), [
    [241679.76, 49351.84, 0, 'direct'], [238543.97, 61142.43, 3135.79, 'direct'],
    [159908.56, 9942.26, 0, 'direct'], [157123.19, 21013.95, 2785.37, 'direct'],
  ]);
  for (const r of c.rows) {
    for (const key of ['profit', 'revenue', 'ads', 'bureau', 'tax'] as const) assert.equal(r.directResult.control[key], r.result.control[key]);
    assert.ok(Math.abs(r.directResult.control.orders - r.result.control.orders) < 1e-8);
    assert.deepEqual(r.directResult, project(c.base, snapshot.targets, r.directScenario));
    assert.equal(r.directScenario.start, c.start);
    assert.equal(r.directScenario.cycles.reduce((n, row) => n + row.days, 0), c.days);
    assert.ok(Math.abs(r.directResult.delta - r.result.delta - r.directAdvantage) < 1e-6);
    assert.equal(r.beatsControl, true);
  }
});
test('a direct jump has one learning transition from day one; zero losses reconcile exactly', () => {
  const c = compareGradualPlan(snapshot);
  assert.equal(c.directAffectedDays, 7);
  for (const r of c.rows) {
    assert.equal(r.directAffectedDays, r.withLearning ? 7 : 0);
    assert.ok(r.directScenario.cycles.every(cycle => cycle.budget === 5000));
    if (r.withLearning) {
      assert.deepEqual(r.directScenario.cycles.map(cycle => [cycle.days, cycle.learningPct]), [[7, 6], [91, 0]]);
      assert.equal(r.affectedDays, 42);
    }
  }
  for (const settings of [{ days: 0, lossPct: 6 }, { days: 7, lossPct: 0 }]) {
    const zero = compareGradualPlan(snapshot, settings);
    assert.equal(zero.directAffectedDays, 0);
    assert.equal(zero.rows[0].directResult.scenario.profit, zero.rows[1].directResult.scenario.profit);
    assert.equal(zero.rows[2].directResult.scenario.profit, zero.rows[3].directResult.scenario.profit);
  }
  const long = compareGradualPlan(snapshot, { days: 30, lossPct: 6 });
  assert.equal(long.directAffectedDays, 30); assert.equal(long.affectedDays, 92);
});
test('future taxes also affect direct projections and a lower-profit base can reverse the winner', () => {
  const normal = compareGradualPlan(snapshot);
  const raised = compareGradualPlan({ ...snapshot, targets: [...snapshot.targets, { month: '2027-01-01', tax_pct: 15 }] });
  assert.ok(raised.rows.every((r, i) => r.directResult.scenario.profit < normal.rows[i].directResult.scenario.profit));
  const lowMargin = compareGradualPlan({ ...snapshot, revenue: snapshot.revenue.map(r => ({ ...r, revenue: r.paid_orders * 4 })) });
  assert.ok(lowMargin.rows.every(r => r.winner === 'gradual' && r.directAdvantage < 0 && !r.beatsControl));
  assert.ok(lowMargin.rows.every(r => r.directResult.scenario.profit < 0 && r.result.delta < 0));
});
test('equal net profit is represented as a tie, with no false winner or gain over keeping', () => {
  const zeroMargin: Snapshot = { ...snapshot, targets: [{ month: '2026-10-01', tax_pct: 0 }],
    revenue: snapshot.revenue.map(r => ({ ...r, revenue: 2790, paid_orders: 279 })),
    bureau: snapshot.bureau.map(r => ({ ...r, custo_bureau: 0 })),
    ads: snapshot.ads.map(r => ({ ...r, cost_micros: 2790e6 })) };
  const c = compareGradualPlan(zeroMargin);
  assert.equal(c.rows[0].winner, 'tie'); assert.equal(c.rows[0].directAdvantage, 0);
  assert.equal(c.rows[0].beatsControl, false);
});
