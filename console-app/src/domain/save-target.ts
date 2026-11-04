import type { SupabaseClient } from '@supabase/supabase-js';
import type { Target } from './targets';
import { readConsoleAccess } from './users';
import { canEditTargets } from './target-access';
/** Match the existing auth/RLS contract and reject concurrent replacement. */
export async function saveTarget(client: SupabaseClient, row: Target, previous?: Target) {
  const { data: identity, error: authError } = await client.auth.getUser();
  if (authError || !identity.user) throw new Error('Entre novamente para salvar as metas.');
  const access=await readConsoleAccess(client,identity.user.id);
  if(!access?.active||!canEditTargets(access.role)) throw new Error('Somente administradores e gerentes ativos podem alterar as metas da empresa.');
  if (previous && previous.month !== row.month) throw new Error('O mês foi alterado. Revise os dados antes de salvar.');
  // Managers edit business targets; tax remains an administrator setting.
  const { tax_pct: _tax, ...targets } = row;
  const payload = access.role === 'administrator' ? row : targets;
  let request;
  if (previous) {
    const update = client.from('monthly_targets').update(payload).eq('month', previous.month);
    request = previous.updated_at ? update.eq('updated_at', previous.updated_at).select('month') : update.is('updated_at', null).select('month');
  } else request = client.from('monthly_targets').insert(payload).select('month');
  const { data, error } = await request;
  if (error || data?.length !== 1 || data[0].month !== row.month) throw new Error('A alteração não foi salva. Sua sessão pode não ter permissão ou este mês foi alterado por outra pessoa. Atualize as metas antes de tentar novamente.');
}
