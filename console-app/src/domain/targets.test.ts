import test from 'node:test';
import assert from 'node:assert/strict';
import { targetPayload } from './targets';
test('blank target fields preserve monthly inheritance rather than becoming zero', () => {
  const row = targetPayload('2026-10', { target_cpa: '0', tax_pct: '' });
  assert.equal(row.month, '2026-10-01'); assert.equal(row.target_cpa, 0); assert.equal(row.tax_pct, null); assert.equal(row.revenue_target, null);
});
test('rejects invalid periods and values before a production write', () => {
  for (const month of ['2026-00', '2026-13', '2026-10-01', 'bad']) assert.throws(() => targetPayload(month, {}));
  for (const value of ['-1', '101', 'NaN', 'Infinity']) assert.throws(() => targetPayload('2026-10', { tax_pct: value }));
  assert.throws(() => targetPayload('2026-10', { target_upsell_pct: '101' }));
  assert.equal(targetPayload('2026-10', { tax_pct: '8' }).tax_pct, 8);
});
