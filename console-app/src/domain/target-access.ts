import type { ConsoleRole } from './users';

export function canEditTargets(role: ConsoleRole | null) {
  return role === 'administrator' || role === 'manager';
}
