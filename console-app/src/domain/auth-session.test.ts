import test from 'node:test';
import assert from 'node:assert/strict';
import { AuthApiError, AuthRetryableFetchError, AuthSessionMissingError, type SupabaseClient } from '@supabase/supabase-js';
import { AUTH_STORAGE_KEY as key, createAuthStorage } from '../lib/auth-storage';
import { readSessionAccess } from './auth-session';

function store() {
  const values = new Map<string, string>();
  return {getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => { values.set(key, value); }, removeItem: (key: string) => { values.delete(key); }};
}

test('remembered login survives a new tab and SDK token renewal stays persistent', () => {
  const disk = store(), tab = store();
  const storage = createAuthStorage(disk, tab);
  storage.setRemember(true); storage.setItem(key, 'session-one');
  assert.equal(tab.getItem(key), null);
  const reopened = createAuthStorage(disk, store());
  assert.equal(reopened.getItem(key), 'session-one');
  reopened.setItem(key, 'refreshed-session');
  assert.equal(disk.getItem(key), 'refreshed-session');
});
test('unchecked login stays in the tab and clears the previous persistent session', () => {
  const disk = store(), tab = store(); disk.setItem(key, 'old-session');
  const storage = createAuthStorage(disk, tab);
  storage.setRemember(false); storage.setItem(key, 'tab-session');
  storage.setItem(key, 'renewed-tab-session');
  assert.equal(disk.getItem(key), null);
  assert.equal(tab.getItem(key), 'renewed-tab-session');
  assert.equal(createAuthStorage(disk, store()).getItem(key), null);
});
test('existing tab session is restored without automatically making it persistent', () => {
  const disk = store(), tab = store(); tab.setItem(key, 'legacy-session');
  const storage = createAuthStorage(disk, tab);
  assert.equal(storage.getItem(key), 'legacy-session');
  storage.setItem(key, 'renewed-session');
  assert.equal(disk.getItem(key), null);
});
test('a tab opened before remembered sign-in in another tab preserves its persistence', () => {
  const disk = store(), tab = store(), storage = createAuthStorage(disk, tab);
  disk.setItem(key, 'other-tab-login');
  storage.setItem(key, 'renewed-session');
  assert.equal(disk.getItem(key), 'renewed-session');
  assert.equal(tab.getItem(key), null);
});
test('SDK sign-out removes both stores so remembered sessions cannot reappear', () => {
  const disk = store(), tab = store(); disk.setItem(key, 'one'); tab.setItem(key, 'two');
  createAuthStorage(disk, tab).removeItem(key);
  assert.equal(disk.getItem(key), null); assert.equal(tab.getItem(key), null);
});

function client(error: unknown, member: unknown = {active: true, role: 'manager'}, memberError: unknown = null) {
  const query = {eq() {return this;}, maybeSingle: async () => ({data: member, error: memberError})};
  return {auth: {getUser: async () => ({data: {user: error ? null : {id: 'operator'}}, error})}, from: () => ({select: () => query})} as unknown as SupabaseClient;
}
test('temporary network/server/rate-limit errors offer retry instead of requesting new login', async () => {
  for (const error of [new AuthRetryableFetchError('offline', 0), new AuthApiError('unavailable', 503, undefined), new AuthApiError('rate limit', 429, 'over_request_rate_limit')]) {
    assert.deepEqual(await readSessionAccess(client(error)), {status: 'error', user: null, role: null});
  }
});
test('missing, expired or revoked session still requires sign-in', async () => {
  for (const error of [new AuthSessionMissingError(), new AuthApiError('revoked', 400, 'session_not_found'), new AuthApiError('invalid', 401, 'bad_jwt')]) {
    assert.equal((await readSessionAccess(client(error))).status, 'signed-out');
  }
});
test('restored sessions still require active server membership and fail closed on lookup error', async () => {
  assert.equal((await readSessionAccess(client(null))).role, 'manager');
  assert.equal((await readSessionAccess(client(null, {active: false, role: 'administrator'}))).status, 'denied');
  assert.equal((await readSessionAccess(client(null, null))).status, 'denied');
  const unavailable = await readSessionAccess(client(null, null, {code: '503'}));
  assert.equal(unavailable.status, 'error'); assert.equal(unavailable.role, null);
});
