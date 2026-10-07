import config from './production-config.json';
import type { SupabaseClient } from '@supabase/supabase-js';
let client: Promise<SupabaseClient> | undefined;
export function getAuthClient() {
  return client ||= import('@supabase/supabase-js').then(({ createClient }) => createClient(config.url, config.key, {
    auth: { storage: sessionStorage, storageKey: 'vp-production-auth', persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
  })).catch(error => { client = undefined; throw error; });
}
export async function readPublic<T>(table: 'monthly_targets' | 'conversion_dispatches_public', query: URLSearchParams, signal?: AbortSignal): Promise<T[]> {
  const response = await fetch(`${config.url}/rest/v1/${table}?${query}`, { headers: { apikey: config.key, Authorization: `Bearer ${config.key}` }, signal });
  if (!response.ok) throw new Error(`Fonte indisponível (HTTP ${response.status}). Tente atualizar.`);
  return response.json();
}
