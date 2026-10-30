import { test } from 'node:test';
import assert from 'node:assert/strict';
import { netMarginPercent, costValue } from './overview-series';
test('net margin uses revenue denominator, preserves loss and rejects missing base', () => {
  assert.equal(netMarginPercent(100, 20), 20);
  assert.equal(netMarginPercent(100, -15), -15);
  assert.equal(netMarginPercent(0, 20), null);
  assert.equal(netMarginPercent(100, NaN), null);
});
test('cost modes select source fields without converting missing values to zero', () => {
  const day = {cost:80,costBureau:20,costTotal:100};
  assert.equal(costValue(day,'total'),100);
  assert.equal(costValue(day,'ads'),80);
  assert.equal(costValue(day,'bureau'),20);
  assert.equal(costValue({cost:0},'ads'),0);
  assert.equal(costValue({},'bureau'),null);
});
