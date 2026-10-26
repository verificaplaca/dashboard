import test from 'node:test';
import assert from 'node:assert/strict';
import { buildBaseline, defaultScenario, type Snapshot } from './simulation';
import { optimizationSettings } from './investment-optimizer';
import { compareCacScenarios } from './cac-comparison';
const snapshot: Snapshot = { start: '2026-10-21', end: '2026-10-22', loadedAt: '', warnings: [], targets: [{ month: '2026-09-01', tax_pct: 8 }], revenue: [{ date: '2026-10-21', revenue: 5748.61, paid_orders: 349 }, { date: '2026-10-22', revenue: 5026.67, paid_orders: 323 }], bureau: [{ date: '2026-10-21', custo_bureau: 934.43 }, { date: '2026-10-22', custo_bureau: 722.64 }], ads: [{ date: '2026-10-21', campaign_id: 'a', cost_micros: 3184.05e6 }, { date: '2026-10-22', campaign_id: 'a', cost_micros: 2345.05e6 }] };
const base = buildBaseline(snapshot), settings = optimizationSettings(defaultScenario(base));
test('the twelve CAC comparisons reproduce the financial projections already shown to the user', () => {
  const rows = compareCacScenarios(base, snapshot.targets, settings);
  assert.deepEqual(rows.map(r => r.budget), [3400, 4100, 5000]);
  assert.deepEqual(rows.map(r => [r.constant.scenario.profit, r.constantLearning.scenario.profit, r.elastic.scenario.profit, r.elasticLearning.scenario.profit]), [
    [164342.08, 162209.74, 144174.66, 142128.77],
    [198177.56, 195606.25, 152702.62, 150326.12],
    [241679.76, 238543.97, 159908.56, 157123.19],
  ]);
  for (const row of rows) for (const r of [row.constant, row.constantLearning, row.elastic, row.elasticLearning]) {
    assert.equal(r.control.profit, 133627.9);
    assert.ok(Math.abs(r.scenario.profit - r.control.profit - r.delta.profit) < .000001);
  }
});
test('fixed comparison hypotheses preserve CAC regardless of the general optimizer elasticity or media share', () => {
  const p = { ...settings, elasticity: .6, responsivePct: 25 }, copy = structuredClone(p);
  const rows = compareCacScenarios(base, snapshot.targets, p);
  for (const r of rows) {
    assert.ok(Math.abs(r.constant.scenario.cpa! - base.ads / base.orders) < 1e-10);
    assert.ok(r.elastic.scenario.cpa! > r.constant.scenario.cpa!);
    assert.ok(r.constantLearning.scenario.cpa! > r.constant.scenario.cpa!);
    assert.equal(r.elastic.scenario.profit, compareCacScenarios(base, snapshot.targets, settings).find(x => x.budget === r.budget)!.elastic.scenario.profit);
  }
  assert.deepEqual(p, copy);
});
test('a manual reference uses the respective hypothesis control and keeping avoids a second learning loss', () => {
  const rows = compareCacScenarios(base, snapshot.targets, { ...settings, controlBudget: 3400 });
  const keep = rows[0];
  assert.equal(keep.constantLearning.delta.profit, 0); assert.equal(keep.elasticLearning.delta.profit, 0);
  assert.equal(keep.constantLearning.learningCost, 0); assert.equal(keep.elasticLearning.learningCost, 0);
  assert.notEqual(keep.constant.control.profit, keep.elastic.control.profit);
  assert.ok(rows[2].constantLearning.delta.profit > 0);
  assert.throws(() => compareCacScenarios({ ...base, issues: ['Fonte incompleta'] }, snapshot.targets, settings));
});
