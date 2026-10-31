import test from 'node:test';
import assert from 'node:assert/strict';
import { balanceChartScales, paddedChartRange } from './chart-scales';

test('SVG histories retain losses, references and finite ranges for flat or missing data', () => {
  const range = paddedChartRange([-100, 200, null, NaN, 300]);
  assert.ok(range.min < -100);
  assert.ok(range.max > 300);
  assert.deepEqual(paddedChartRange([0, 0, null]), { min: 0, max: 1 });
  assert.deepEqual(paddedChartRange([]), { min: 0, max: 1 });
  assert.equal(paddedChartRange([-100]).max, 0);
  assert.equal(paddedChartRange([100]).min, 0);
});

test('unbounded financial axes keep zero and reserve space for positive and negative results', () => {
  const options = { scales: { y: { ticks: {} }, x: { ticks: { maxTicksLimit: 10 } } } } as any;
  balanceChartScales(options, 1200);
  assert.equal(options.scales.y.beginAtZero, true);
  assert.equal(options.scales.y.grace, '20%');
  assert.equal(options.scales.y.min, undefined);
  assert.equal(options.scales.y.max, undefined);
});

test('dual axes keep their distinct currency and order formatters', () => {
  const money = (v: number) => `R$ ${v}`, orders = (v: number) => v;
  const options = { scales: { y: { ticks: { callback: money } }, y1: { beginAtZero: true, ticks: { callback: orders, precision: 0 } } } } as any;
  balanceChartScales(options, 390);
  assert.equal(options.scales.y.ticks.callback, money);
  assert.equal(options.scales.y1.ticks.callback, orders);
  assert.equal(options.scales.y1.ticks.precision, 0);
  assert.equal(options.scales.y1.grace, '20%');
});

test('percentage limits and intentional explicit bounds stay fixed', () => {
  const options = { scales: { y: { min: 0, max: 100, ticks: {} }, y1: { min: -10, max: 40, ticks: {} } } } as any;
  balanceChartScales(options, 390);
  assert.equal(options.scales.y.min, 0);
  assert.equal(options.scales.y.max, 100);
  assert.equal(options.scales.y.grace, undefined);
  assert.equal(options.scales.y1.min, -10);
  assert.equal(options.scales.y1.max, 40);
});

test('mobile reduces horizontal labels without changing the selected budget range', () => {
  const options = { scales: { x: { type: 'linear', min: 0, max: 10000, ticks: { maxTicksLimit: 6 } } } } as any;
  balanceChartScales(options, 390);
  assert.equal(options.scales.x.ticks.maxTicksLimit, 4);
  assert.equal(options.scales.x.min, 0);
  assert.equal(options.scales.x.max, 10000);
  balanceChartScales(options, 1200);
  assert.equal(options.scales.x.ticks.maxTicksLimit, 6);
});

test('custom cycle labels, miniatures, nonlinear and radial charts keep their scales', () => {
  const cycleTicks = () => {};
  const options = { scales: { x: { afterBuildTicks: cycleTicks, ticks: { maxTicksLimit: 12 } }, y: { display: false, min: -1, max: 2, grace: 0 }, yLog: { type: 'logarithmic', ticks: {} }, r: { min: 0, max: 100 } } } as any;
  balanceChartScales(options, 320);
  assert.equal(options.scales.x.afterBuildTicks, cycleTicks);
  assert.equal(options.scales.x.ticks.maxTicksLimit, 12);
  assert.equal(options.scales.y.grace, 0);
  assert.equal(options.scales.yLog.beginAtZero, undefined);
  assert.deepEqual(options.scales.r, { min: 0, max: 100 });
});
