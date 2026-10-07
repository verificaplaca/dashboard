import { createContext, useContext, useEffect, useRef, useState, type ReactNode } from 'react';
import type { User } from '@supabase/supabase-js';
import { getAuthClient } from './api';
import { useUI } from './ui';
import { Dialog } from './components';
type Auth = { user: User | null; openLogin: () => void; signOut: () => Promise<void>; refresh: () => Promise<void> };
const Context = createContext<Auth | null>(null);
export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null), [login, setLogin] = useState(false);
  const [requested, setRequested] = useState(() => { try { return !!sessionStorage.getItem('vp-production-auth'); } catch { return false; } });
  const { notify } = useUI();
  const revision = useRef(0);
  const refresh = async () => {
    const request = ++revision.current;
    const client = await getAuthClient();
    const { data, error } = await client.auth.getUser();
    if (request === revision.current) setUser(error ? null : data.user);
  };
  useEffect(() => {
    if (!requested) return;
    let active = true, unsubscribe = () => {};
    getAuthClient().then(client => {
      if (!active) return;
      const { data } = client.auth.onAuthStateChange(event => {
        if (event === 'SIGNED_OUT') { revision.current++; setUser(null); }
        else { queueMicrotask(() => { if (active) void refresh().catch(() => setUser(null)); }); }
      });
      unsubscribe = () => data.subscription.unsubscribe();
    }).catch(() => notify('Não foi possível verificar a sessão. A leitura permanece disponível.'));
    return () => { active = false; revision.current++; unsubscribe(); };
  }, [requested]);
  const signOut = async () => {
    const client = await getAuthClient();
    const { error } = await client.auth.signOut({ scope: 'local' });
    if (error) { notify('Não foi possível encerrar a sessão. Tente novamente.'); return; }
    revision.current++; setUser(null); notify('Sessão encerrada.');
  };
  return <Context.Provider value={{ user, refresh, openLogin: () => { setRequested(true); setLogin(true); }, signOut }}>{children}{login && <Login close={() => setLogin(false)} refresh={refresh} />}</Context.Provider>;
}
export function useAuth() { const value = useContext(Context); if (!value) throw new Error('AuthProvider ausente'); return value; }
function Login({ close, refresh }: { close: () => void; refresh: () => Promise<void> }) {
  const [busy, setBusy] = useState(false), [error, setError] = useState('');
  return <Dialog title="Entrar no Verifica Placa" close={close}><form className="live-auth-form" onSubmit={async event => {
    event.preventDefault(); if (busy) return;
    const form = event.currentTarget, values = new FormData(form);
    setBusy(true); setError('');
    try {
      const client = await getAuthClient();
      const { error } = await client.auth.signInWithPassword({ email: String(values.get('email')).trim(), password: String(values.get('password')) });
      if (error) throw new Error('Não foi possível entrar. Confira seu email e senha.');
      await refresh(); form.reset(); close();
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Falha ao entrar. Tente novamente.'); }
    finally { setBusy(false); }
  }}><div className="form-fields live-form"><p>Use a mesma conta que edita as metas no dashboard atual.</p><label>Email<input autoFocus name="email" type="email" autoComplete="username" required /></label><label>Senha<input name="password" type="password" autoComplete="current-password" required /></label>{error && <p role="alert">{error}</p>}</div><div className="dialog-footer"><button type="button" className="button secondary" onClick={close}>Cancelar</button><button className="button primary" disabled={busy}>{busy ? 'Entrando…' : 'Entrar'}</button></div></form></Dialog>;
}
