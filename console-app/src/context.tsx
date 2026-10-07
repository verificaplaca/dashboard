import React, { createContext, useContext, useEffect, useRef, useState } from 'react';
import { UIProvider } from './ui';
import { createSeed, reducer, can, type State } from './domain/engine';

const STORAGE = 'verifica-placa-console-v1';
type Context = { state: State; run: (type: string, payload?: any, message?: string) => State | undefined; can: (permission: string) => boolean; user: State['users'][number]; toast: string; notify: (message: string) => void };
const DemoContext = createContext<Context | null>(null);
export * from './format';

function readState(): State {
  try {
    const stored = localStorage.getItem(STORAGE);
    if (stored) {
      const parsed = JSON.parse(stored);
      const seed = createSeed();
      if (parsed.schemaVersion === seed.schemaVersion && Array.isArray(parsed.orders) && Array.isArray(parsed.users) && parsed.users.some((u: any) => u.id === parsed.currentUserId)) return parsed;
      localStorage.setItem(`${STORAGE}-backup-${Date.now()}`, stored);
    }
  } catch { /* Invalid demo storage is retained; a new session can still open. */ }
  return createSeed();
}
export function DemoProvider({ children }: { children: React.ReactNode }) {
  const [state, setState] = useState<State>(readState);
  const stateRef = useRef(state);
  stateRef.current = state;
  const [toast, setToast] = useState('');
  useEffect(() => { try { localStorage.setItem(STORAGE, JSON.stringify(state)); } catch { setToast('O navegador não permitiu salvar. Suas alterações duram nesta sessão.'); } }, [state]);
  useEffect(() => { if (!toast) return; const timer = setTimeout(() => setToast(''), 4500); return () => clearTimeout(timer); }, [toast]);
  const user = state.users.find(u => u.id === state.currentUserId) || state.users[0];
  const run = (type: string, payload?: any, message = 'Alteração salva na demonstração.') => {
    try { const next = reducer(stateRef.current, { type, payload }); stateRef.current = next; setState(next); setToast(message); return next; }
    catch (e) { setToast(e instanceof Error ? e.message : 'Não foi possível concluir a ação.'); return undefined; }
  };
  return <UIProvider theme={state.settings.theme} setTheme={theme => run('settings.update', { theme })} notify={setToast}><DemoContext.Provider value={{ state, run, can: p => user.active && can(user.role, p, state), user, toast, notify: setToast }}>{children}</DemoContext.Provider></UIProvider>;
}
export function useDemo() { const context = useContext(DemoContext); if (!context) throw new Error('DemoProvider ausente'); return context; }
