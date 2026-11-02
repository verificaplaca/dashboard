import { createContext, useContext, useEffect, useRef, useState, type ReactNode } from 'react';
import type { User } from '@supabase/supabase-js';
import { getAuthClient, getSessionStorage } from './api';
import { AUTH_STORAGE_KEY } from './lib/auth-storage';
import { readSessionAccess } from './domain/auth-session';
import { useUI } from './ui';
import { Dialog } from './components';
import { LoginPage } from '@/components/ui/sign-in-page';
import type { ConsoleRole } from './domain/users';
type AccessStatus = 'loading' | 'signed-out' | 'ready' | 'denied' | 'error';
type Auth = { role: ConsoleRole | null; status: AccessStatus; user: User | null; openLogin: () => void; signOut: () => Promise<void>; refresh: () => Promise<void>; signIn: (email: string, password: string, remember: boolean) => Promise<void> };
const Context = createContext<Auth | null>(null);
export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null), [login, setLogin] = useState(false);
  const [requested, setRequested] = useState(() => { try { return !!getSessionStorage().getItem(AUTH_STORAGE_KEY); } catch { return false; } });
  const [role, setRole] = useState<ConsoleRole | null>(null), [status, setStatus] = useState<AccessStatus>(requested ? 'loading' : 'signed-out');
  const { notify } = useUI();
  const revision = useRef(0);
  const refresh = async () => {
    const request = ++revision.current;
    const client = await getAuthClient();
    const access = await readSessionAccess(client);
    if (request !== revision.current) return;
    setUser(access.user); setRole(access.role); setStatus(access.status);
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
    window.addEventListener('online',check);
    return () => { clearInterval(timer); window.removeEventListener('focus',check); window.removeEventListener('online',check); };
  },[requested]);
  const signOut = async () => {
    const client = await getAuthClient();
    const { error } = await client.auth.signOut({ scope: 'local' });
    if (error) { notify('Não foi possível encerrar a sessão. Tente novamente.'); return; }
    revision.current++; setUser(null); setRole(null); setStatus('signed-out'); notify('Sessão encerrada.');
  };
  const signIn = async (email: string, password: string, remember: boolean) => {
    const client = await getAuthClient();
    getSessionStorage().setRemember(remember);
    setRequested(true);
    const { error } = await client.auth.signInWithPassword({ email, password });
    if (error) throw new Error('Não foi possível entrar. Confira seu email e senha.');
    await refresh();
    setLogin(false);
  };
  return <Context.Provider value={{ user, role, status, refresh, signIn, openLogin: () => { setRequested(true); setLogin(true); }, signOut }}>{children}{login && <Dialog title="Entrar no Verifica Placa" close={() => setLogin(false)}><LoginPage compact onSignIn={signIn} /></Dialog>}</Context.Provider>;
}
export function useOptionalAuth() { return useContext(Context); }
export function useAuth() { const value = useContext(Context); if (!value) throw new Error('AuthProvider ausente'); return value; }
