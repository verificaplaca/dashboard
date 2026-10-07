import { createContext, useContext, type ReactNode } from 'react';
type UI = { theme: string; setTheme: (theme: string) => void; notify: (message: string) => void };
const Context = createContext<UI | null>(null);
export function UIProvider({ children, ...value }: UI & { children: ReactNode }) { return <Context.Provider value={value}>{children}</Context.Provider>; }
export function useUI() { const value = useContext(Context); if (!value) throw new Error('UIProvider ausente'); return value; }
