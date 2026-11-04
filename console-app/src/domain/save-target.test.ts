import test from 'node:test';
import assert from 'node:assert/strict';
import type { SupabaseClient } from '@supabase/supabase-js';
import { saveTarget } from './save-target';
import { targetPayload } from './targets';
function fake(authed: boolean, rows: { month: string }[] | null, denied = false, role = 'administrator', active = true) {
  const calls: unknown[][] = [];
  const query = { update: (row: unknown) => { calls.push(['update', row]); return query; }, insert: (row: unknown) => { calls.push(['insert', row]); return query; }, eq: (...args: unknown[]) => { calls.push(['eq', ...args]); return query; }, is: (...args: unknown[]) => { calls.push(['is', ...args]); return query; }, select: async () => ({ data: rows, error: denied ? new Error('RLS denied') : null }) };
  const client = { auth: { getUser: async () => ({ data: { user: authed ? { id: 'real-auth-id' } : null }, error: null }) }, from: (table: string) => { calls.push(['from', table]); return table==='console_users'?{select:()=>({eq:()=>({maybeSingle:async()=>({data:{role,active},error:null})})})}:query; } } as unknown as SupabaseClient;
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

test('active managers can insert targets and update a version while leaving tax untouched',async()=>{
 const row=targetPayload('2026-10',{revenue_target:'15000',tax_pct:'99'});
 for(const previous of [undefined,{...row,tax_pct:8,updated_at:'observed-version'}]){
  const {client,calls}=fake(true,[{month:row.month}],false,'manager');
  await saveTarget(client,row,previous);
  const mutation=calls.find(call=>call[0]===(previous?'update':'insert'))!;
  assert.equal((mutation[1] as Record<string,unknown>).revenue_target,15000);
  assert.equal(Object.hasOwn(mutation[1] as object,'tax_pct'),false);
  if(previous)assert.ok(calls.some(call=>call[0]==='eq'&&call[1]==='updated_at'&&call[2]==='observed-version'));
 }
});
test('inactive administrators, inactive managers and unknown profiles never submit a target mutation',async()=>{
 for(const [role,active] of [['manager',false],['administrator',false],['unsupported',true]] as const){
  const {client,calls}=fake(true,[],false,role,active);
  await assert.rejects(saveTarget(client,targetPayload('2026-10',{})),/administradores e gerentes ativos/);
  assert.ok(!calls.some(call=>call[0]==='insert'||call[0]==='update'));
 }
});
test('administrators retain tax editing and managers cannot overwrite a stale month',async()=>{
 const row=targetPayload('2026-10',{tax_pct:'9'}),admin=fake(true,[{month:row.month}]);
 await saveTarget(admin.client,row);
 assert.equal((admin.calls.find(call=>call[0]==='insert')![1] as typeof row).tax_pct,9);
 const manager=fake(true,[],false,'manager');
 await assert.rejects(saveTarget(manager.client,row,{...row,updated_at:'stale'}),/não foi salva/);
});
