import React, { lazy, Suspense } from 'react';
import {parseRecoveryEntry} from './domain/password-recovery';
import ReactDOM from 'react-dom/client';
import './tailwind.css';
import './styles.css';
import { LoadingProvider, LoadingFallback } from './Loading';
import './layout.css';
import './live.css';
import './chart-tooltips.css';
const Production = lazy(() => import('./Production'));
const Demo = lazy(() => import('./DemoEntry'));
const Recovery=lazy(()=>import('./PasswordRecovery'));
const recovery=parseRecoveryEntry(location.search,location.hash);
if(recovery){history.replaceState(null,'',`${location.pathname}?recovery=1`);document.querySelector('meta[name=referrer]')?.remove();const meta=document.createElement('meta');meta.name='referrer';meta.content='no-referrer';document.head.append(meta);}
const demo = new URLSearchParams(location.search).get('demo') === '1';
if(!demo){try{document.documentElement.dataset.theme=localStorage.getItem('vp-production-theme')||'light';}catch{document.documentElement.dataset.theme='light';}}
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><LoadingProvider><Suspense fallback={<LoadingFallback/>}>{recovery ? <Recovery entry={recovery}/> : demo ? <Demo /> : <Production />}</Suspense></LoadingProvider></React.StrictMode>);
