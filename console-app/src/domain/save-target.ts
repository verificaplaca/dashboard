import type { SupabaseClient } from '@supabase/supabase-js';
import type { Target } from './targets';
/** Match the existing auth/RLS contract and reject concurrent replacement. */
export async function saveTarget(client: SupabaseClient, row: Target, previous?: Target) {
  const { data: identity, error: authError } = await client.auth.getUser();
  if (authError || !identity.user) throw new Error('Entre novamente para salvar as metas.');
  if (previous && previous.month !== row.month) throw new Error('O mês foi alterado. Revise os dados antes de salvar.');
  let request;
  if (previous) {
    const update = client.from('monthly_targets').update(row).eq('month', previous.month);
    request = previous.updated_at ? update.eq('updated_at', previous.updated_at).select('month') : update.is('updated_at', null).select('month');
  } else request = client.from('monthly_targets').insert(row).select('month');
  const { data, error } = await request;
  if (error || data?.length !== 1 || data[0].month !== row.month) throw new Error('A alteração não foi salva. Sua sessão pode não ter permissão ou este mês foi alterado por outra pessoa. Atualize as metas antes de tentar novamente.');
}
