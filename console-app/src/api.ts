import { readPublicRows } from './domain/public-read';
import config from './production-config.json';
import type { SupabaseClient } from '@supabase/supabase-js';
import { AUTH_STORAGE_KEY, createAuthStorage } from './lib/auth-storage';
let authStorage: ReturnType<typeof createAuthStorage> | undefined;
export function getSessionStorage() {
  return authStorage ||= createAuthStorage(localStorage, sessionStorage);
}
let client: Promise<SupabaseClient> | undefined;
export function getAuthClient() {
  return client ||= import('@supabase/supabase-js').then(({ createClient }) => createClient(config.url, config.key, {
    auth: { storage: getSessionStorage(), storageKey: AUTH_STORAGE_KEY, persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
  })).catch(error => { client = undefined; throw error; });
}
export async function readPublic<T>(table: 'monthly_targets' | 'conversion_dispatches_public' | 'revenue_daily' | 'google_ads_campaign_daily', query: URLSearchParams, signal?: AbortSignal): Promise<T[]> {
  return readPublicRows<T>(`${config.url}/rest/v1/${table}?${query}`,{apikey:config.key,Authorization:`Bearer ${config.key}`},signal);
}
