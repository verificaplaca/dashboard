import test from 'node:test';
import assert from 'node:assert/strict';
import { addDays, buildBaseline, dates, defaultScenario, project, taxAt, validateScenario, type Snapshot } from './simulation';
const snapshot: Snapshot = { start: '2026-09-01', end: '2026-09-02', loadedAt: '', warnings: [], targets: [{ month: '2026-09-01', tax_pct: 8 }], revenue: [{ date: '2026-09-01', revenue: 2000, paid_orders: 100 }, { date: '2026-09-02', revenue: 2000, paid_orders: 100 }], bureau: [{ date: '2026-09-01', custo_bureau: 200 }, { date: '2026-09-02', custo_bureau: 200 }], ads: [{ date: '2026-09-01', campaign_id: 'a', cost_micros: 1000e6 }, { date: '2026-09-02', campaign_id: 'a', cost_micros: 1000e6 }] };
const base = buildBaseline(snapshot);
const initial = { ...defaultScenario(base), baselineBudget: 1000, ticket: 20, bureauUnit: 2, cycles: [{ id: 'C1', days: 10, budget: 1000, seasonality: 1, learningPct: 0 }] };
test('real financial base reconciles and includes spend-only days', () => {
  assert.equal(base.profit, 1280); assert.equal(base.adsDay, 1000); assert.equal(base.ticket, 20); assert.equal(base.bureauUnit, 2);
  const b = buildBaseline({ ...snapshot, bureau: [...snapshot.bureau, { date: '2026-09-02', custo_bureau: 50 }] }); assert.equal(b.bureau, 450); assert.ok(b.issues.some(i => i.includes('duplicados')));
  const noSales = structuredClone(snapshot); noSales.revenue[1] = { date: '2026-09-02', revenue: 0, paid_orders: 0 }; assert.equal(buildBaseline(noSales).ads, 2000); assert.equal(buildBaseline(noSales).profit, -560);
});
test('equal spend and unchanged assumptions yield zero incremental profit', () => {
  const r = project(base, snapshot.targets, initial); assert.equal(r.delta, 0); assert.equal(r.scenario.profit, 6400); assert.equal(r.scenario.orders, 1000); assert.equal(r.marginalCpa, null);
});
test('linear doubling has exact financial delta and marginal CPA', () => {
  const r = project(base, snapshot.targets, { ...initial, elasticity: 1, cycles: [{ ...initial.cycles[0], budget: 2000 }] });
  assert.equal(r.extraOrders, 1000); assert.equal(r.extraAds, 10000); assert.equal(r.delta, 6400); assert.equal(r.marginalCpa, 10); assert.ok(Math.abs(r.breakEvenCpa - 16.4) < 1e-10);
});
test('no media-responsive orders yields higher spend with no extra sales', () => {
  const r = project(base, snapshot.targets, { ...initial, responsivePct: 0, cycles: [{ ...initial.cycles[0], budget: 2000 }] });
  assert.equal(r.extraOrders, 0); assert.equal(r.delta, -10000); assert.equal(r.marginalCpa, null); assert.ok(r.extraRevenueNeeded! > 12000);
});
test('seasonality affects both alternatives; learning and conversion affect only scenario', () => {
  const r = project(base, snapshot.targets, { ...initial, conversionPct: 25, cycles: [{ ...initial.cycles[0], seasonality: 1.2, learningPct: 20 }] });
  assert.equal(r.control.orders, 1200); assert.ok(Math.abs(r.scenario.orders - 1200) < 1e-8); assert.equal(r.delta, 0);
});
test('tax inheritance is per field, never from future rows; zero is a real rate', () => {
  const t = [{ month: '2026-09-01', tax_pct: 8 }, { month: '2026-10-01', tax_pct: null }, { month: '2026-11-01', tax_pct: 0 }];
  assert.deepEqual(taxAt(t, '2026-08-31'), { value: 8, fallback: true }); assert.equal(taxAt(t, '2026-10-31').value, 8); assert.equal(taxAt(t, '2026-11-01').value, 0);
  const r = project(base, t, { ...initial, start: '2026-10-31', cycles: [{ ...initial.cycles[0], days: 2 }] }); assert.equal(r.control.tax, 160); assert.equal(r.control.profit, 1440);
});
test('extra fees and fixed costs reconcile separately and use same assumptions', () => {
  const r = project(base, snapshot.targets, { ...initial, extraFeePct: 2, fixedDay: 100 });
  assert.equal(r.scenario.panelProfit, 6400); assert.equal(r.scenario.fees, 400); assert.equal(r.scenario.fixed, 1000); assert.equal(r.scenario.profit, 5000); assert.equal(r.delta, 0);
});
test('zero scenario spend is valid and CPA unavailable; unit losses have no break-even revenue', () => {
  const zero = project(base, snapshot.targets, { ...initial, elasticity: .8, cycles: [{ ...initial.cycles[0], budget: 0 }] }); assert.equal(zero.scenario.orders, 0); assert.equal(zero.scenario.cpa, null);
  const loss = project(base, snapshot.targets, { ...initial, bureauUnit: 30 }); assert.ok(loss.breakEvenCpa < 0); assert.equal(loss.extraRevenueNeeded, null);
});
test('invalid inputs, reversed periods, incomplete and duplicate financial bases block projection', () => {
  assert.throws(() => dates('2026-02-30', '2026-03-01')); assert.throws(() => dates('2026-10-22', '2026-10-01'));
  assert.ok(validateScenario({ ...initial, elasticity: NaN }).length); assert.ok(validateScenario({ ...initial, cycles: [{ ...initial.cycles[0], days: 1.5 }] }).length);
  const incomplete = buildBaseline({ ...snapshot, revenue: snapshot.revenue.slice(0, 1) }); assert.ok(incomplete.issues.length); assert.throws(() => project(incomplete, snapshot.targets, initial));
});
test('PDF daily budgets preserve 98 days and exact investment; dates cross year safely', () => {
  const s = defaultScenario(base); const r = project(base, snapshot.targets, { ...s, baselineBudget: 2790 });
  assert.equal(r.scenario.days, 98); assert.equal(r.scenario.ads, 350560); assert.equal(r.control.ads, 273420); assert.equal(r.cycles.at(-1)!.end, '2027-01-31'); assert.equal(addDays('2026-12-31', 1), '2027-01-01');
});

test('unchanged fractional unit economics preserve zero delta without rounding the averages', () => {
  const fractional = structuredClone(snapshot); fractional.revenue[0].revenue = 2000.13; fractional.bureau[0].custo_bureau = 200.07;
  const b = buildBaseline(fractional); const s = defaultScenario(b);
  const r = project(b, fractional.targets, { ...s, cycles: s.cycles.map(c => ({ ...c, budget: s.baselineBudget })) });
  assert.equal(r.delta, 0);
});
