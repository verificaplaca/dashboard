import { can, type State, type User } from './engine';

export { saoPauloClock, greeting } from './clock';
export function individualProgress(state: State, userId: string) {
  const credits = state.credits.filter(c => c.userId === userId);
  const confirmed = credits.filter(c => c.status === 'confirmed').reduce((sum, c) => sum + c.points, 0);
  const pending = credits.filter(c => c.status === 'pending').reduce((sum, c) => sum + c.points, 0);
  const thresholds = [0, 250, 750, 1500, 3000];
  const level = thresholds.filter(t => confirmed >= t).length;
  const floor = thresholds[level - 1], next = thresholds[level];
  return { confirmed, pending, level, next: next ?? null, progress: next === undefined ? 100 : Math.max(0, Math.min(100, (confirmed - floor) / (next - floor) * 100)) };
}
export function progressGoals(state: State, user: User) {
  if (!user.active || !can(user.role, 'view.evolution', state)) return [];
  const global = ['owner', 'admin'].includes(user.role);
  return state.goals.filter(g => g.unit !== 'currency' && (g.scope === 'organization' || (g.scope === 'team' && (global || g.ownerId === user.teamId)) || (g.scope === 'individual' && g.ownerId === user.id)));
}
export function inboxScope(state: State, user: User) {
  if (!user.active || !can(user.role, 'view.inbox', state)) return [];
  return state.conversations.filter(c => ['owner', 'admin'].includes(user.role) || c.ownerId === user.id || state.users.find(u => u.id === c.ownerId)?.teamId === user.teamId || (!c.ownerId && state.instances.find(i => i.id === c.instanceId)?.teamId === user.teamId));
}
export function canCompleteTask(state: State, user: User, ownerId: string) {
  return user.active && can(user.role, 'tasks.write', state) && (['owner', 'admin'].includes(user.role) || ownerId === user.id || (user.role === 'manager' && state.users.find(u => u.id === ownerId)?.teamId === user.teamId));
}
