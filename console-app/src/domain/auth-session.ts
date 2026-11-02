import type { SupabaseClient, User } from '@supabase/supabase-js';
import { readConsoleAccess, type ConsoleRole } from './users';

type SessionAccess = {
  status: 'signed-out' | 'ready' | 'denied' | 'error';
  user: User | null;
  role: ConsoleRole | null;
};

export async function readSessionAccess(client: SupabaseClient): Promise<SessionAccess> {
  const { data, error } = await client.auth.getUser();
  if (error) {
    const ended = error.name === 'AuthSessionMissingError'
      || ['session_not_found', 'session_expired', 'refresh_token_not_found', 'refresh_token_already_used', 'bad_jwt', 'user_not_found'].includes(error.code ?? '')
      || error.status === 401 || error.status === 403;
    return { status: ended ? 'signed-out' : 'error', user: null, role: null };
  }
  if (!data.user) return { status: 'signed-out', user: null, role: null };
  try {
    const member = await readConsoleAccess(client, data.user.id);
    return { status: member?.active ? 'ready' : 'denied', user: data.user, role: member?.active ? member.role : null };
  } catch {
    return { status: 'error', user: data.user, role: null };
  }
}
