import test from 'node:test';
import assert from 'node:assert/strict';
import type { SupabaseClient } from '@supabase/supabase-js';
import { saveTarget } from './save-target';
import { targetPayload } from './targets';
function fake(authed: boolean, rows: { month: string }[] | null, denied = false) {
  const calls: unknown[][] = [];
  const query = { update: (row: unknown) => { calls.push(['update', row]); return query; }, insert: (row: unknown) => { calls.push(['insert', row]); return query; }, eq: (...args: unknown[]) => { calls.push(['eq', ...args]); return query; }, is: (...args: unknown[]) => { calls.push(['is', ...args]); return query; }, select: async () => ({ data: rows, error: denied ? new Error('RLS denied') : null }) };
  const client = { auth: { getUser: async () => ({ data: { user: authed ? { id: 'real-auth-id' } : null }, error: null }) }, from: (table: string) => { calls.push(['from', table]); return query; } } as unknown as SupabaseClient;
  return { client, calls };
}
test('expired or missing identity never sends a monthly target write', async () => {
  const { client, calls } = fake(false, []);
  await assert.rejects(saveTarget(client, targetPayload('2026-10', {})), /Entre novamente/); assert.deepEqual(calls, []);
});
test('a stale version or RLS rejection is a failure, never a success notification', async () => {
  const row = targetPayload('2026-10', {});
  for (const denied of [false, true]) {
    const { client, calls } = fake(true, [], denied);
    await assert.rejects(saveTarget(client, row, { ...row, updated_at: 'previous-version' }), /não foi salva/);
    assert.ok(calls.some(call => call[0] === 'eq' && call[1] === 'updated_at' && call[2] === 'previous-version'));
  }
});
test('inserts a new month without upserting over someone else’s newly created row', async () => {
  const row = targetPayload('2026-10', { tax_pct: '8' });
  const { client, calls } = fake(true, [{ month: row.month }]);
  await saveTarget(client, row); assert.ok(calls.some(call => call[0] === 'insert')); assert.ok(!calls.some(call => call[0] === 'update'));
});
