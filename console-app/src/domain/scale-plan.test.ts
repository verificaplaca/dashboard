import test from 'node:test';
import assert from 'node:assert/strict';
import { addDays, buildBaseline, dates, project, type Snapshot } from './simulation';
import { actualScaleWindow, evaluateScalePlan, initialScaleState, recordScaleStage, requiredScaleDays, scaleScenario, SCALE_LEVELS, validateScaleState, type ScaleState } from './scale-plan';
function source(start: string, end: string, spend = 2790, profit = 1000, orders = 350): Snapshot {
  const calendar = dates(start, end);
  return { start, end, loadedAt: '2026-11-15T12:00:00Z', warnings: [], targets: [{ month: '2026-01-01', tax_pct: 8 }], revenue: calendar.map(date => ({ date, revenue: (profit + spend + 700) / .92, paid_orders: orders })), bureau: calendar.map(date => ({ date, custo_bureau: 700 })), ads: calendar.map(date => ({ date, campaign_id: 'a', cost_micros: spend * 1e6 })) };
}
const today = '2026-11-15';
function stage1() {
  const state = initialScaleState(), s = source('2026-10-01', '2026-11-14');
  const next = recordScaleStage(state, s, 1, '2026-11-01', today);
  return next;
}
function setDayProfit(s: Snapshot, date: string, profit: number) { const a = s.ads.filter(r => r.date === date).reduce((sum, r) => sum + r.cost_micros / 1e6, 0), b = s.bureau.find(r => r.date === date)!.custo_bureau; s.revenue.find(r => r.date === date)!.revenue = (profit + a + b) / .92; }
test('initial real six full days stay on hold despite the last two days recovering', () => {
  const s = source('2026-10-01', '2026-10-22', 2975.4, 959.5, 318); s.loadedAt = '2026-10-23T12:00:00Z';
  setDayProfit(s, '2026-10-21', 1170.24); setDayProfit(s, '2026-10-22', 1556.85);
  const r = evaluateScalePlan(s, initialScaleState(), '2026-10-23');
  assert.equal(r.action, 'hold'); assert.equal(r.elapsed, 6); assert.equal(r.required, 14); assert.ok(r.recent!.profitDay > r.current!.profitDay); assert.ok(r.current!.cac! > 9);
  assert.equal(r.cumulative!.start, '2026-10-16'); assert.equal(r.current!.start, '2026-10-17');
});
test('qualifying initial base allows just the next test and freezes an actual 14-day reference when applied', () => {
  const s = source('2026-10-01', '2026-11-14'), r = evaluateScalePlan(s, initialScaleState(), today);
  assert.equal(r.action, 'increase'); assert.equal(r.nextLevel, 1); assert.equal(r.current!.days, 14);
  const state = recordScaleStage(initialScaleState(), s, 1, '2026-11-01', today);
  assert.equal(state.entries.at(-1)!.reference!.window.end, '2026-10-31'); assert.equal(state.entries.at(-1)!.reference!.window.profitDay, 1000); assert.deepEqual(validateScaleState(state), []);
});
test('CAC is a ratio of totals, never the mean of daily CACs', () => {
  const s = source('2026-11-01', '2026-11-02', 100, 100, 10); s.revenue[1].paid_orders = 100;
  assert.ok(Math.abs(actualScaleWindow(s, s.start, s.end)!.cac! - 200 / 110) < 1e-8);
});
test('more investment requires +5% profit, more actual spend and a recovered cumulative transition', () => {
  const state = stage1(), s = source('2026-10-01', '2026-11-14', 3100, 1100, 400);
  // Thirteen full days after Nov 1 is not enough.
  assert.equal(evaluateScalePlan(s, state, today).action, 'hold');
  const full = source('2026-10-01', '2026-11-15', 3100, 1100, 400); full.loadedAt = '2026-11-16T12:00:00Z';
  assert.equal(evaluateScalePlan(full, state, '2026-11-16').action, 'increase');
  const small = source('2026-10-01', '2026-11-15', 3100, 1040, 400); small.loadedAt = full.loadedAt;
  assert.equal(evaluateScalePlan(small, state, '2026-11-16').action, 'hold');
  const noSpend = source('2026-10-01', '2026-11-15', 2790, 1100, 400); noSpend.loadedAt = full.loadedAt;
  assert.equal(evaluateScalePlan(noSpend, state, '2026-11-16').action, 'hold');
  setDayProfit(full, '2026-11-01', -2000);
  const loss = evaluateScalePlan(full, state, '2026-11-16'); assert.equal(loss.action, 'hold'); assert.ok(loss.recovered! < 0); assert.equal(loss.current!.profitDay, 1100);
});
test('profit drop, no gain after extension and seven-day combined CAC/profit signal reduce', () => {
  const state = stage1();
  const drop = source('2026-10-01', '2026-11-15', 3100, 940, 400); drop.loadedAt = '2026-11-16T12:00:00Z';
  const r = evaluateScalePlan(drop, state, '2026-11-16'); assert.equal(r.action, 'reduce'); assert.equal(r.reduceLevel, 0);
  const flat = source('2026-10-01', '2026-11-22', 3100, 1000, 400); flat.loadedAt = '2026-11-23T12:00:00Z';
  assert.ok(evaluateScalePlan(flat, state, '2026-11-23').reasons.some(r => r.includes('extensão')));
  const bad = source('2026-10-01', '2026-11-08', 3100, 990, 310); bad.loadedAt = '2026-11-09T12:00:00Z';
  assert.equal(evaluateScalePlan(bad, state, '2026-11-09').action, 'reduce');
  bad.revenue.forEach(r => setDayProfit(bad, r.date, 1200));
  assert.equal(evaluateScalePlan(bad, state, '2026-11-09').action, 'hold'); // High CAC alone is not a reduction rule.
});
test('three-day cumulative loss contains scale early; explicit zero sales still recognizes real losses', () => {
  const s = source('2026-10-01', '2026-10-19'); s.loadedAt = '2026-10-20T12:00:00Z';
  for (const d of ['2026-10-17', '2026-10-18', '2026-10-19']) { s.revenue.find(r => r.date === d)!.revenue = 0; s.revenue.find(r => r.date === d)!.paid_orders = 0; s.bureau.find(r => r.date === d)!.custo_bureau = 0; }
  const r = evaluateScalePlan(s, initialScaleState(), '2026-10-20'); assert.equal(r.action, 'reduce'); assert.equal(r.current!.cac, null); assert.equal(r.reduceLevel, null); assert.equal(r.current!.days, 3);
});
test('missing sources, duplicates, tax fallback and stale reads cannot produce a green or a fabricated loss', () => {
  const s = source('2026-10-01', '2026-11-14');
  for (const broken of [{ ...s, ads: s.ads.slice(0, -1) }, { ...s, bureau: s.bureau.slice(0, -1) }, { ...s, revenue: [...s.revenue, s.revenue.at(-1)!] }, { ...s, targets: [] }, { ...s, warnings: ['Metas indisponíveis'] }, { ...s, loadedAt: '2026-11-14T12:00:00Z' }]) {
    const r = evaluateScalePlan(broken, initialScaleState(), today); assert.equal(r.action, 'hold'); assert.equal(r.valid, false);
  }
});
test('São Paulo date boundary, current-day exclusion and configurable maturity are respected', () => {
  const s = source('2026-10-01', '2026-11-14'); s.loadedAt = '2026-11-15T01:00:00Z';
  assert.equal(evaluateScalePlan(s, initialScaleState(), today).valid, false);
  s.loadedAt = '2026-11-15T12:00:00Z'; s.end = today;
  assert.equal(evaluateScalePlan(s, initialScaleState(), today).valid, false);
  s.end = '2026-11-14'; const r = evaluateScalePlan(s, { ...initialScaleState(), lagDays: 3, conversionDays: 10 }, today);
  assert.equal(r.eligibleEnd, '2026-11-11'); assert.equal(r.required, 21); assert.equal(r.current!.days, 21);
  assert.equal(requiredScaleDays(60), 126);
});
test('unapproved jumps lack a validated reference and cannot turn green', () => {
  const s = source('2026-10-01', '2026-11-14');
  const state = recordScaleStage(initialScaleState(), s, 3, '2026-10-20', today);
  const r = evaluateScalePlan(s, state, today); assert.equal(r.action, 'hold'); assert.equal(r.previous, null);
  assert.throws(() => recordScaleStage(state, s, 4, '2026-11-16', today)); assert.throws(() => recordScaleStage(state, s, 4, '2026-10-20', today));
});
test('a rollback uses the previously validated same-level baseline and revalidates before retesting', () => {
  const state = stage1(), s = source('2026-10-01', '2026-11-30'); s.loadedAt = '2026-12-01T12:00:00Z';
  const reduced = recordScaleStage(state, s, 0, '2026-11-16', '2026-12-01');
  const r = evaluateScalePlan(s, reduced, '2026-12-01'); assert.equal(r.previous!.level, 0); assert.equal(r.action, 'increase'); assert.equal(r.nextLevel, 1);
});
test('top level holds instead of recommending above R$ 5,000 and malformed local state is rejected', () => {
  const state = stage1(), ref = state.entries.at(-1)!.reference!;
  const cap: ScaleState = { ...state, entries: [{ ...state.entries.at(-1)!, level: 6, reference: { ...ref, level: 5 } }] };
  const s = source('2026-10-01', '2026-11-15', 5000, 1200, 650); s.loadedAt = '2026-11-16T12:00:00Z';
  assert.equal(evaluateScalePlan(s, cap, '2026-11-16').action, 'hold'); assert.equal(evaluateScalePlan(s, cap, '2026-11-16').nextLevel, null); assert.equal(SCALE_LEVELS.at(-1), 5000);
  assert.ok(validateScaleState({ ...state, lagDays: NaN }).length);
  const corrupt = structuredClone(state); corrupt.entries.at(-1)!.reference!.window.profit = 999999;
  assert.ok(validateScaleState(corrupt).length); assert.equal(evaluateScalePlan(s, corrupt, '2026-11-16').action, 'hold');
});

test('applying the gradual calendar to the existing engine reconciles 98 days and R$ 389,940 of planned media', () => {
  const s = source('2026-10-01', '2026-10-22'), b = buildBaseline(s), scenario = scaleScenario(b), r = project(b, s.targets, scenario);
  assert.equal(r.scenario.days, 98); assert.equal(r.scenario.ads, 389940); assert.equal(r.control.ads, 273420); assert.equal(addDays(scenario.start, 97), '2027-01-31'); assert.ok(scenario.cycles.every(c => c.learningPct === 0)); assert.equal(scenario.elasticity, .8); assert.equal(scenario.ticket, b.ticket);
});
test('exact gain boundary can advance; changed maturity and unexplained spend cannot', () => {
  const state = stage1(), s = source('2026-10-01', '2026-11-15', 3100, 1050, 400); s.loadedAt = '2026-11-16T12:00:00Z';
  assert.equal(evaluateScalePlan(s, state, '2026-11-16').action, 'increase');
  const high = source('2026-10-01', '2026-11-15', 5000, 1200, 650); high.loadedAt = s.loadedAt;
  assert.equal(evaluateScalePlan(high, state, '2026-11-16').action, 'hold');
  const longer = source('2026-10-01', '2026-11-22', 3100, 1200, 400); longer.loadedAt = '2026-11-23T12:00:00Z';
  const r = evaluateScalePlan(longer, { ...state, conversionDays: 10 }, '2026-11-23');
  assert.equal(r.current!.days, 21); assert.equal(r.previous!.window.days, 14); assert.equal(r.action, 'hold');
});
