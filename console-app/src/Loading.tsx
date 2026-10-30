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
export function LoadingScreen() {
 const ref=useRef<HTMLDivElement>(null);
 useLayoutEffect(()=>{
  const before=document.activeElement as HTMLElement|null,releaseScroll=lockBodyScroll();
  ref.current?.focus({preventScroll:true});
  return()=>{releaseScroll();queueMicrotask(()=>{if(before?.isConnected&&!before.closest('[inert]'))before.focus({preventScroll:true});});};
 },[]);
 return <div ref={ref} tabIndex={-1} onKeyDown={event=>{event.stopPropagation();if(event.key==='Tab')event.preventDefault();}} className="global-loading-overlay" role="status" aria-live="polite" aria-label="Carregando dados do painel">
  <div className="global-loading-center"><GooeyLoader role="presentation" aria-label={undefined} primaryColor="#a0f644" secondaryColor="#d2ffad" borderColor="var(--border)"/><p>Carregando dados<span>Aguarde um instante.</span></p></div>
 </div>;
}
export function LoadingProvider({children}:{children:ReactNode}) {
 const [tracker]=useState(createLoadingTracker);
 return <Context.Provider value={tracker}><LoadingLayer tracker={tracker}>{children}</LoadingLayer></Context.Provider>;
}
function LoadingLayer({tracker,children}:{tracker:ReturnType<typeof createLoadingTracker>;children:ReactNode}) {
 const pending=useSyncExternalStore(tracker.subscribe,tracker.getSnapshot,tracker.getSnapshot),[settling,setSettling]=useState(false);
 const visible=pending||settling;
 useLayoutEffect(()=>{
  if(pending){setSettling(true);return;}
  // Allow the final data render (and chained reads) to settle before revealing it.
  const timer=setTimeout(()=>setSettling(false),160);
  return()=>clearTimeout(timer);
 },[pending]);
 return <><div className="global-loading-content" inert={visible} aria-hidden={visible||undefined}>{children}</div>{visible&&<LoadingScreen/>}</>;
}
