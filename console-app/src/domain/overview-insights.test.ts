import test from 'node:test';
import assert from 'node:assert/strict';
import { buildMetricDetail, buildPeriodSummary, type InsightSnapshot } from './overview-insights';
const snapshot: InsightSnapshot = { totals: { rev: 10000, cost: 3000, costTotal: 5000, conv: 300, pft: 5000, netPftOpt: 4200, cac: 10, margin: .5, upsell: 25 }, previous: { conv: 200, cost: 2400, cac: 12, rev: 8000, margin: .45, upsell: 30 }, targets: { cpa: 9, upsell: 32, netProfit: 5000 }, months: 1, real: true, dayCount: 12, upsellOrders: 75 };
test('summary separates positive result, missed target and improving acquisition cost', () => {
  assert.equal(buildPeriodSummary(snapshot), 'Resultado positivo, abaixo da meta mensal. CAC melhorou em relação ao período anterior.');
  assert.match(buildPeriodSummary({ ...snapshot, dayCount: 0 }), /^Sem dados/);
  assert.match(buildPeriodSummary({ ...snapshot, real: false }), /^Dados de exemplo/);
});
test('CAC uses BRL not cents, lower is better, and target is independent of comparison', () => {
  const detail = buildMetricDetail('cac', snapshot);
  assert.match(detail.value, /10,00/);
  assert.match(detail.comparison, /Redução.*2,00.*melhora/);
  assert.match(detail.target, /1,00 acima do teto/);
  assert.equal(detail.tone, 'negative');
});
test('margin and upsell comparisons use percentage points with their respective input units', () => {
  assert.match(buildMetricDetail('margin', snapshot).comparison, /5 p.p./);
  assert.match(buildMetricDetail('upsell', snapshot).comparison, /5 p.p.*piora/);
  assert.match(buildMetricDetail('upsell', snapshot).target, /7 p.p. abaixo/);
  assert.equal(buildMetricDetail('upsell', snapshot).rows[1].value, '75');
});
test('empty denominators and absent previous periods never become claims of improvement', () => {
  const empty = { ...snapshot, totals: { ...snapshot.totals, conv: 0, rev: 0 }, previous: null };
  for (const key of ['cac', 'margin', 'upsell'] as const) {
    assert.equal(buildMetricDetail(key, empty).value, '—');
    assert.equal(buildMetricDetail(key, empty).tone, 'neutral');
  }
  assert.match(buildMetricDetail('cac', { ...snapshot, previous: null }).comparison, /indisponível/);
});
test('cross-month scope does not compare consolidated metrics against a single monthly target', () => {
  const multi = { ...snapshot, months: 2 };
  assert.doesNotMatch(buildPeriodSummary(multi), /meta mensal/);
  assert.match(buildMetricDetail('cac', multi).target, /variam por mês/);
  assert.equal(buildMetricDetail('upsell', multi).tone, 'neutral');
});
test('demo upsell preserves original calculation and never invents numerator', () => {
  const detail = buildMetricDetail('upsell', { ...snapshot, real: false, upsellOrders: null });
  assert.match(detail.formula, /Média/);
  assert.match(detail.rows[1].value, /Não disponível/);
});
