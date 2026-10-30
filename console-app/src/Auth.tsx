import { createContext, useContext, useEffect, useRef, useState, type ReactNode } from 'react';
import type { User } from '@supabase/supabase-js';
import { getAuthClient } from './api';
import { useUI } from './ui';
import { Dialog } from './components';
import { readConsoleAccess, type ConsoleRole } from './domain/users';
type AccessStatus = 'loading' | 'signed-out' | 'ready' | 'denied' | 'error';
type Auth = { role: ConsoleRole | null; status: AccessStatus; user: User | null; openLogin: () => void; signOut: () => Promise<void>; refresh: () => Promise<void> };
const Context = createContext<Auth | null>(null);
export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null), [login, setLogin] = useState(false);
  const [requested, setRequested] = useState(() => { try { return !!sessionStorage.getItem('vp-production-auth'); } catch { return false; } });
  const [role, setRole] = useState<ConsoleRole | null>(null), [status, setStatus] = useState<AccessStatus>(requested ? 'loading' : 'signed-out');
  const { notify } = useUI();
  const revision = useRef(0);
  const refresh = async () => {
    const request = ++revision.current;
    const client = await getAuthClient();
    const { data, error } = await client.auth.getUser();
    if (request !== revision.current) return;
    if(error || !data.user) { setUser(null); setRole(null); setStatus('signed-out'); return; }
    try {
      const member = await readConsoleAccess(client, data.user.id);
      if(request !== revision.current) return;
      setUser(data.user); setRole(member?.active ? member.role : null); setStatus(member?.active ? 'ready' : 'denied');
    } catch { if(request === revision.current) { setUser(data.user); setRole(null); setStatus('error'); } }
  };
  useEffect(() => {
    if (!requested) return;
    let active = true, unsubscribe = () => {};
    getAuthClient().then(client => {
      if (!active) return;
      const { data } = client.auth.onAuthStateChange(event => {
        if (event === 'SIGNED_OUT') { revision.current++; setUser(null); setRole(null); setStatus('signed-out'); }
        else { queueMicrotask(() => { if (active) void refresh().catch(() => { setRole(null); setStatus('error'); }); }); }
      });
      unsubscribe = () => data.subscription.unsubscribe();
    }).catch(() => { setRole(null); setStatus('error'); });
    return () => { active = false; revision.current++; unsubscribe(); };
  }, [requested]);
  useEffect(() => {
    if(!requested) return;
    const check = () => { if(document.visibilityState==='visible') void refresh().catch(() => { setRole(null); setStatus('error'); }); };
    const timer=setInterval(check,30_000);
    window.addEventListener('focus',check);
    return () => { clearInterval(timer); window.removeEventListener('focus',check); };
  },[requested]);
  const signOut = async () => {
    const client = await getAuthClient();
    const { error } = await client.auth.signOut({ scope: 'local' });
    if (error) { notify('Não foi possível encerrar a sessão. Tente novamente.'); return; }
    revision.current++; setUser(null); setRole(null); setStatus('signed-out'); notify('Sessão encerrada.');
  };
  return <Context.Provider value={{ user, role, status, refresh, openLogin: () => { setRequested(true); setLogin(true); }, signOut }}>{children}{login && <Login close={() => setLogin(false)} refresh={refresh} />}</Context.Provider>;
}
export function useOptionalAuth() { return useContext(Context); }
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
  }}><div className="form-fields live-form"><p>Entre com sua conta de gerente ou administrador. Novos acessos são criados por um administrador.</p><label>Email<input autoFocus name="email" type="email" autoComplete="username" required /></label><label>Senha<input name="password" type="password" autoComplete="current-password" required /></label>{error && <p role="alert">{error}</p>}</div><div className="dialog-footer"><button type="button" className="button secondary" onClick={close}>Cancelar</button><button className="button primary" disabled={busy}>{busy ? 'Entrando…' : 'Entrar'}</button></div></form></Dialog>;
}
