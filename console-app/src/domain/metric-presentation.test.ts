import test from 'node:test';
import assert from 'node:assert/strict';
import { comparisonStatus, goalProgress, resultStatus, targetStatus } from './metric-presentation';

test('comparison direction respects cost per order without treating all spend growth as loss', () => {
  assert.equal(comparisonStatus(-2, true).tone, 'positive');
  assert.equal(comparisonStatus(-2000).tone, 'negative');
  assert.equal(comparisonStatus(1000, false, true).tone, 'neutral');
  assert.equal(comparisonStatus(null).text, 'Base indisponível');
  assert.equal(comparisonStatus(0).text, 'Sem mudança');
});

test('positive profit below target and a loss above a negative previous result remain distinct', () => {
  assert.equal(resultStatus(50400).tone, 'positive');
  assert.equal(targetStatus(50400, 55000).tone, 'negative');
  assert.equal(resultStatus(-100).tone, 'negative');
  assert.equal(resultStatus(0).tone, 'neutral');
});
test('acquisition ceilings reverse target direction; unavailable targets are never called failures', () => {
  assert.equal(targetStatus(6, 5.5, true).tone, 'negative');
  assert.equal(targetStatus(5.5, 5.5, true).tone, 'positive');
  assert.equal(targetStatus(null, 5.5, true).text, 'Sem dados');
  assert.equal(targetStatus(6, null, true).text, 'Meta não definida');
  assert.equal(targetStatus(6, 0, true).tone, 'negative');
  assert.equal(targetStatus(0, 0, true).tone, 'positive');
  assert.equal(targetStatus(NaN, 5.5).tone, 'neutral');
});
test('goal labels retain losses and overachievement while only the decorative track is clamped', () => {
  const normal = goalProgress(50400, 55000)!;
  assert.equal(normal.remaining, 4600);
  assert.ok(Math.abs(normal.percent - 91.636363636) < .000001);
  assert.deepEqual(goalProgress(-20, 100), { percent: -20, track: 0, remaining: 120 });
  assert.deepEqual(goalProgress(120, 100), { percent: 120, track: 100, remaining: -20 });
  assert.equal(goalProgress(10, 0), null);
  assert.equal(goalProgress(Infinity, 100), null);
});
