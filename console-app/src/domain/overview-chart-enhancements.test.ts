import test from 'node:test';
import assert from 'node:assert/strict';
import { comparisonSeries, enhanceOverviewChart, targetDistance } from '../overview-chart-enhancements';

test('comparison shifts by calendar duration, preserves gaps and never extends predictions into prior data', () => {
  const current = [{ dateStr: '2026-03-01' }, { dateStr: '2026-03-03' }];
  const result = comparisonSeries(current, [{ dateStr: '2026-02-26', revenue: 0 }, { dateStr: '2026-02-27', revenue: 20 }], 'revenue', 4);
  assert.deepEqual(result.dates, ['2026-02-26', '2026-02-28']);
  assert.deepEqual(result.data, [0, null, null, null]);
});
test('comparisons include losses and use the equal previous window across years', () => {
  assert.deepEqual(comparisonSeries([{ dateStr: '2026-01-01' }], [{ dateStr: '2025-12-31', netProfit: -12 }], 'netProfit').data, [-12]);
});
test('target distances distinguish acquisition ceilings from upsell floors', () => {
  assert.equal(targetDistance(9.54, 9, true)?.tone, 'negative');
  assert.equal(targetDistance(8, 9, true)?.tone, 'positive');
  assert.equal(targetDistance(26.1, 32, false, true)?.text, '5,9 p.p. abaixo da meta');
  assert.equal(targetDistance(null, 32), null);
});
test('comparison augments only supported charts and preserves original observations and formatter', () => {
  const dataset = { data: [12], label: 'Receita' }, formatter = () => 'original';
  const config: any = { data: { labels: ['1/1'], datasets: [dataset] }, options: { plugins: { tooltip: { callbacks: { label: formatter } } } } };
  const snapshot = { current: [{ dateStr: '2026-01-01', revenue: 12 }], allDays: [{ dateStr: '2025-12-31', revenue: 10 }], compare: true, reducedMotion: true };
  enhanceOverviewChart('revenueChart', config, snapshot);
  assert.strictEqual(config.data.datasets[0], dataset);
  assert.strictEqual(config.options.plugins.tooltip.callbacks.label, formatter);
  assert.deepEqual(config.data.datasets[1].data, [10]);
  assert.equal(config.options.animation, false);
  const mini: any = { data: { datasets: [dataset] }, options: {} };
  enhanceOverviewChart('ks_r', mini, snapshot);
  assert.deepEqual(mini, { data: { datasets: [dataset] }, options: {} });
});
test('unavailable previous window adds no series', () => {
  const config: any = { data: { labels: ['1/1'], datasets: [] }, options: {} };
  enhanceOverviewChart('profitChart', config, { current: [{ dateStr: '2026-01-01' }], allDays: [], compare: true, reducedMotion: false });
  assert.equal(config.data.datasets.length, 0);
});
test('tooltip avoids undefined percentage from zero baseline and reports missing equivalent dates', () => {
  const config: any = { data: { labels: ['1/1', '2/1'], datasets: [] }, options: {} };
  enhanceOverviewChart('revenueChart', config, { current: [{ dateStr: '2026-01-01', revenue: 12 }, { dateStr: '2026-01-02', revenue: 14 }], allDays: [{ dateStr: '2025-12-30', revenue: 0 }], compare: true, reducedMotion: false });
  const callback = config.options.plugins.tooltip.callbacks.afterBody;
  assert.equal(callback([{ dataIndex: 0 }]).some((line: string) => /Infinity|NaN|%/.test(line)), false);
  assert.match(callback([{ dataIndex: 1 }]).join(' '), /sem dados/);
});
