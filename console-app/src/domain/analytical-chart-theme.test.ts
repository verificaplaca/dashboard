import test from 'node:test';
import assert from 'node:assert/strict';
import { styleAnalyticalChart } from '../analytical-chart-theme';

const chart = (data: any[], colors: string[] = []) => ({ type: 'bar', data: { labels: ['A', 'B', 'C'], datasets: [{ data, backgroundColor: colors }] }, options: { plugins: { datalabels: { display: true }, tooltip: { callbacks: { label: (value: unknown) => String(value) } } }, scales: { y: { min: -10, max: 40, ticks: {} }, x: { ticks: {} } } } });

test('profit styling preserves signed observations, missing values, labels and formatting callbacks', () => {
  const config = chart([-12, 0, null]);
  const values = config.data.datasets[0].data;
  const labels = config.data.labels;
  const format = config.options.plugins.tooltip.callbacks.label;
  styleAnalyticalChart('profitChart', config);
  assert.strictEqual(config.data.datasets[0].data, values);
  assert.strictEqual(config.data.labels, labels);
  assert.strictEqual(config.options.plugins.tooltip.callbacks.label, format);
  assert.deepEqual(config.data.datasets[0].backgroundColor, ['#fb8792', '#a8dc69', 'transparent']);
  assert.equal(config.options.scales.y.min, -10);
});

test('CAC presentation respects the original threshold classifications instead of recalculating them', () => {
  const config = chart([8, 10, 14], ['rgba(22,163,74,.7)', 'rgba(217,119,6,.7)', 'rgba(220,38,38,.7)']);
  styleAnalyticalChart('cacChart', config);
  assert.deepEqual(config.data.datasets[0].backgroundColor, ['#a8dc69', '#edc75f', '#fb8792']);
  assert.deepEqual(config.data.datasets[0].data, [8, 10, 14]);
});

test('upsell preserves source classifications and expands the ceiling when observations exceed 40%', () => {
  const config = chart([52, 20], ['rgba(124,58,237,.7)', 'rgba(124,58,237,.35)']);
  styleAnalyticalChart('upsellChart', config);
  assert.deepEqual(config.data.datasets[0].backgroundColor, ['#a8dc69', '#fb8792']);
  assert.ok(config.options.scales.y.max >= 52);
  assert.equal(config.options.scales.y.min, -10);
  assert.deepEqual(config.data.datasets[0].data, [52, 20]);
});

test('donut presentation preserves all campaign categories and volumes', () => {
  const config = chart([5, 0, 9]);
  styleAnalyticalChart('convChart', config);
  assert.deepEqual(config.data.labels, ['A', 'B', 'C']);
  assert.deepEqual(config.data.datasets[0].data, [5, 0, 9]);
});

test('unrelated charts are untouched', () => {
  const config = chart([1, 2]);
  const before = JSON.stringify(config);
  styleAnalyticalChart('unknownChart', config);
  assert.equal(JSON.stringify(config), before);
});
