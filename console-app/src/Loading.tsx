import { createContext, useContext, useLayoutEffect, useRef, useState, useSyncExternalStore, type ReactNode } from 'react';
import { GooeyLoader } from '@/components/ui/loader-10';
import { createLoadingTracker } from './domain/loading';
import { lockBodyScroll } from './lib/scroll-lock';
import './loading.css';
const Context=createContext<ReturnType<typeof createLoadingTracker>|null>(null);
export function useGlobalLoading(loading: boolean) {
 const tracker=useContext(Context);
 useLayoutEffect(()=>{if(loading&&tracker)return tracker.begin();},[loading,tracker]);
}
export function LoadingFallback() { useGlobalLoading(true); return null; }
export function LoadingScreen({ exiting = false }: { exiting?: boolean }) {
 const ref=useRef<HTMLDivElement>(null);
 useLayoutEffect(()=>{
  const before=document.activeElement as HTMLElement|null,releaseScroll=lockBodyScroll();
  ref.current?.focus({preventScroll:true});
  return()=>{releaseScroll();queueMicrotask(()=>{if(before?.isConnected&&!before.closest('[inert]'))before.focus({preventScroll:true});});};
 },[]);
 return <div ref={ref} tabIndex={-1} onKeyDown={event=>{event.stopPropagation();if(event.key==='Tab')event.preventDefault();}} className={`global-loading-overlay${exiting ? ' is-exiting' : ''}`} role="status" aria-live="polite" aria-label="Carregando dados do painel">
  <div className="global-loading-center"><GooeyLoader role="presentation" aria-label={undefined} primaryColor="#a0f644" secondaryColor="#d2ffad" borderColor="var(--border)"/><p>Carregando dados<span>Aguarde um instante.</span></p></div>
 </div>;
}
export function LoadingProvider({children}:{children:ReactNode}) {
 const [tracker]=useState(createLoadingTracker);
 return <Context.Provider value={tracker}><LoadingLayer tracker={tracker}>{children}</LoadingLayer></Context.Provider>;
}
function LoadingLayer({tracker,children}:{tracker:ReturnType<typeof createLoadingTracker>;children:ReactNode}) {
 const pending=useSyncExternalStore(tracker.subscribe,tracker.getSnapshot,tracker.getSnapshot),[phase,setPhase]=useState<'hidden'|'settling'|'exiting'>('hidden');
 const visible=pending||phase!=='hidden';
 useLayoutEffect(()=>{
  if(pending){setPhase('settling');return;}
  if(phase==='hidden')return;
  let exitTimer: ReturnType<typeof setTimeout> | undefined;
  // Finish rendering before fading away; retain focus/scroll locks until the
  // fade ends. A new read cancels both timers and restores opacity immediately.
  const settleTimer=setTimeout(()=>{
   if(matchMedia('(prefers-reduced-motion: reduce)').matches){setPhase('hidden');return;}
   setPhase('exiting');
   exitTimer=setTimeout(()=>setPhase('hidden'),320);
  },160);
  return()=>{clearTimeout(settleTimer);clearTimeout(exitTimer);};
 },[pending]);
 return <><div className="global-loading-content" inert={visible} aria-hidden={visible||undefined}>{children}</div>{visible&&<LoadingScreen exiting={!pending&&phase==='exiting'}/>}</>;
}
