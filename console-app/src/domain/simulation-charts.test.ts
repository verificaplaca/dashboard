import test from 'node:test';
import assert from 'node:assert/strict';
import { buildProfitCharts } from './simulation-charts';
import { buildBaseline, defaultScenario, project, type Snapshot } from './simulation';
const data: Snapshot = { start: '2026-09-01', end: '2026-09-01', loadedAt: '', warnings: [], targets: [{ month: '2026-09-01', tax_pct: 8 }], revenue: [{ date: '2026-09-01', revenue: 2000, paid_orders: 100 }], bureau: [{ date: '2026-09-01', custo_bureau: 200 }], ads: [{ date: '2026-09-01', campaign_id: 'a', cost_micros: 1000e6 }] };
const base = buildBaseline(data), scenario = defaultScenario(base);
test('all chart series use net profit after informed fees and fixed costs, and cumulative totals reconcile', () => {
  const s = { ...scenario, extraFeePct: 2, fixedDay: 100 }, r = project(base, data.targets, s), charts = buildProfitCharts(base, data.targets, s);
  assert.equal(charts.metric, 'profit'); assert.equal(charts.current.y, r.scenario.profit); assert.notEqual(charts.current.y, r.scenario.panelProfit);
  assert.equal(charts.control, r.control.profit);
  for (const series of charts.series) {
    assert.deepEqual(series.points[0], { x: 0, y: 0 });
    assert.equal(series.points.at(-1)!.y, series.total); assert.ok(Math.abs(series.cycleProfits.reduce((a, p) => a + p, 0) - series.total) < .01);
  }
});
test('budget curves hold a single control reference and pass through the current plan', () => {
  const charts = buildProfitCharts(base, data.targets, { ...scenario, cycles: scenario.cycles.map(c => ({ ...c, budget: c.budget + .005 })) }), actual = charts.budgetSeries.find(s => s.id === 'current')!;
  assert.deepEqual(actual.points[5], charts.current);
  const control = charts.budgetSeries.find(s => s.id === 'control')!;
  assert.ok(control.points.every(p => p.y === charts.control));
  for (const series of charts.budgetSeries.filter(s => s.id !== 'control')) assert.ok(series.points.every((p, i) => !i || p.x > series.points[i - 1].x));
});
test('negative profit stays negative on the charts and zero spend has no misleading budget curve', () => {
  const loss = buildProfitCharts(base, data.targets, { ...scenario, bureauUnit: 100 });
  assert.ok(loss.current.y < 0); assert.ok(loss.series.find(s => s.id === 'current')!.points.at(-1)!.y < 0);
  const zero = buildProfitCharts(base, data.targets, { ...scenario, cycles: scenario.cycles.map(c => ({ ...c, budget: 0 })) });
  assert.equal(zero.budgetSeries.length, 0); assert.equal(zero.current.x, 0); assert.ok(zero.series.length > 1);
});
test('custom elasticity remains visible alongside presets without duplicating matching presets', () => {
  const custom = buildProfitCharts(base, data.targets, { ...scenario, elasticity: .75 });
  assert.equal(custom.series.length, 5); assert.ok(custom.series.some(s => s.id === 'current' && s.label.includes('0,75')));
  const linear = buildProfitCharts(base, data.targets, { ...scenario, elasticity: 1 });
  assert.equal(linear.series.filter(s => s.id === 'current').length, 1); assert.equal(linear.series.some(s => s.id === 'linear'), false);
});
