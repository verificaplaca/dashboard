import test from 'node:test';
import assert from 'node:assert/strict';
import { createSeed } from './engine';
import { saoPauloClock, greeting, individualProgress, progressGoals, inboxScope, canCompleteTask } from './home';

test('greeting uses all four São Paulo boundaries and the local date', () => {
  const cases = [['2026-10-07T03:00:00Z', 'Boa madrugada'], ['2026-10-07T07:59:00Z', 'Boa madrugada'], ['2026-10-07T08:00:00Z', 'Bom dia'], ['2026-10-07T14:59:00Z', 'Bom dia'], ['2026-10-07T15:00:00Z', 'Boa tarde'], ['2026-10-07T20:59:00Z', 'Boa tarde'], ['2026-10-07T21:00:00Z', 'Boa noite']];
  for (const [at, expected] of cases) assert.equal(greeting(saoPauloClock(new Date(at)).hour), expected);
  assert.equal(saoPauloClock(new Date('2026-10-07T02:59:00Z')).day, '2026-10-06');
});
test('level excludes pending, reverted and other people; boundaries never divide by zero', () => {
  const state = createSeed();
  state.credits = [{ id: 'one', eventId: 'one', userId: 'USR-04', teamId: 'team-sales', label: 'Confirmado', points: 250, status: 'confirmed', at: '' }, { id: 'two', eventId: 'two', userId: 'USR-04', teamId: 'team-sales', label: 'Pendente', points: 1500, status: 'pending', at: '' }, { id: 'three', eventId: 'three', userId: 'USR-04', teamId: 'team-sales', label: 'Estornado', points: 3000, status: 'reverted', at: '' }];
  assert.deepEqual(individualProgress(state, 'USR-04'), { confirmed: 250, pending: 1500, level: 2, next: 750, progress: 0 });
  assert.equal(individualProgress(state, 'USR-01').confirmed, 0);
  state.credits[0].points = 3200;
  assert.deepEqual(individualProgress(state, 'USR-04'), { confirmed: 3200, pending: 1500, level: 5, next: null, progress: 100 });
});
test('progress respects permissions, own individual goals and team boundaries', () => {
  const state = createSeed(), sales = state.users.find(u => u.role === 'sales')!, support = state.users.find(u => u.role === 'support')!;
  assert.ok(progressGoals(state, sales).some(g => g.scope === 'team'));
  assert.ok(progressGoals(state, sales).every(g => g.unit !== 'currency' && (g.scope === 'organization' || g.ownerId === sales.id || g.ownerId === sales.teamId)));
  assert.ok(!progressGoals(state, support).some(g => g.ownerId === 'team-sales'));
  state.rolePermissions = { sales: [] };
  assert.deepEqual(progressGoals(state, sales), []);
  assert.deepEqual(inboxScope(state, sales), []);
  assert.equal(canCompleteTask(state, sales, sales.id), false);
});
test('assigned conversations stay available across channels; tasks cannot be completed across scope', () => {
  const state = createSeed(), sales = state.users.find(u => u.id === 'USR-04')!, manager = state.users.find(u => u.role === 'manager')!;
  assert.ok(inboxScope(state, sales).some(c => c.id === 'CON-1003'));
  assert.ok(!inboxScope(state, sales).some(c => c.ownerId === 'USR-06'));
  assert.equal(canCompleteTask(state, sales, 'USR-05'), false);
  assert.equal(canCompleteTask(state, manager, 'USR-05'), true);
  assert.equal(canCompleteTask(state, manager, 'USR-06'), false);
});
