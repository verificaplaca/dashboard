import { createContext, useContext, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore, type ReactNode } from 'react';
import { GooeyLoader } from '@/components/ui/loader-10';
import { createLoadingTracker, estimateLoadingProgress } from './domain/loading';
import { lockBodyScroll } from './lib/scroll-lock';
import './loading.css';
const Context=createContext<ReturnType<typeof createLoadingTracker>|null>(null);
export function useGlobalLoading(loading: boolean) {
 const tracker=useContext(Context);
 useLayoutEffect(()=>{if(loading&&tracker)return tracker.begin();},[loading,tracker]);
}
export function LoadingFallback() { useGlobalLoading(true); return null; }
export function LoadingScreen({ exiting = false, progress = 0 }: { exiting?: boolean; progress?: number }) {
 const ref=useRef<HTMLDivElement>(null);
 useLayoutEffect(()=>{
  const before=document.activeElement as HTMLElement|null,releaseScroll=lockBodyScroll();
  ref.current?.focus({preventScroll:true});
  return()=>{releaseScroll();queueMicrotask(()=>{if(before?.isConnected&&!before.closest('[inert]'))before.focus({preventScroll:true});});};
 },[]);
 const percent=Math.min(100,Math.max(0,Math.floor(progress)));
 return <div ref={ref} tabIndex={-1} onKeyDown={event=>{event.stopPropagation();if(event.key==='Tab')event.preventDefault();}} className={`global-loading-overlay${exiting ? ' is-exiting' : ''}`} aria-label="Carregando dados do painel">
  <div className="global-loading-center"><GooeyLoader role="presentation" aria-label={undefined} primaryColor="#a0f644" secondaryColor="#d2ffad" borderColor="var(--border)"/><div className="global-loading-progress"><p>Preparando seu painel</p><strong className="global-loading-percent" aria-hidden="true">{percent}<span>%</span></strong><div className="global-loading-track" role="progressbar" aria-label="Carregamento do painel" aria-valuemin={0} aria-valuemax={100} aria-valuenow={percent} aria-valuetext={`${percent}%${percent<100?' — progresso estimado':''}`}><div className="global-loading-fill" style={{transform:`scaleX(${percent/100})`}}/></div><span className="global-loading-caption">{percent===100?'Carregamento concluído':'Progresso estimado'}</span></div><span className="global-loading-announcement" role="status">{percent===100?'Carregamento concluído.':'Carregando o painel. Aguarde um instante.'}</span></div>
 </div>;
}
export function LoadingProvider({children}:{children:ReactNode}) {
 const [tracker]=useState(createLoadingTracker);
 return <Context.Provider value={tracker}><LoadingLayer tracker={tracker}>{children}</LoadingLayer></Context.Provider>;
}
function LoadingLayer({tracker,children}:{tracker:ReturnType<typeof createLoadingTracker>;children:ReactNode}) {
 const snapshot=useSyncExternalStore(tracker.subscribe,tracker.getProgressSnapshot,tracker.getProgressSnapshot),pending=snapshot.pending;
 const [phase,setPhase]=useState<'hidden'|'settling'|'complete'|'exiting'>('hidden'),[progress,setProgress]=useState(0);
 const started=useRef(0),currentSnapshot=useRef(snapshot);currentSnapshot.current=snapshot;
 const visible=pending||phase!=='hidden';
 useEffect(()=>{
  if(!pending)return;
  const update=()=>{const state=currentSnapshot.current;setProgress(value=>estimateLoadingProgress(value,performance.now()-started.current,state.completed,state.total));};
  update();const timer=setInterval(update,100);return()=>clearInterval(timer);
 },[pending]);
 useLayoutEffect(()=>{
  if(pending){if(phase==='hidden'||phase==='complete'||phase==='exiting'){started.current=performance.now();setProgress(0);}setPhase('settling');return;}
  if(phase==='hidden')return;
  let exitTimer: ReturnType<typeof setTimeout> | undefined;
  let completeTimer: ReturnType<typeof setTimeout> | undefined;
  // Finish rendering before fading away; retain focus/scroll locks until the
  // fade ends. A new read cancels both timers and restores opacity immediately.
  const settleTimer=setTimeout(()=>{
   setProgress(100);setPhase('complete');
   completeTimer=setTimeout(()=>{
    if(matchMedia('(prefers-reduced-motion: reduce)').matches){setPhase('hidden');return;}
    setPhase('exiting');
    exitTimer=setTimeout(()=>setPhase('hidden'),320);
   },240);
  },160);
  return()=>{clearTimeout(settleTimer);clearTimeout(completeTimer);clearTimeout(exitTimer);};
 },[pending]);
 return <><div className="global-loading-content" inert={visible} aria-hidden={visible||undefined}>{children}</div>{visible&&<LoadingScreen progress={progress} exiting={!pending&&phase==='exiting'}/>}</>;
}
