import React, { lazy, Suspense } from 'react';
import ReactDOM from 'react-dom/client';
import './tailwind.css';
import './styles.css';
import { LoadingProvider, LoadingFallback } from './Loading';
import './layout.css';
import './live.css';
const Production = lazy(() => import('./Production'));
const Demo = lazy(() => import('./DemoEntry'));
const demo = new URLSearchParams(location.search).get('demo') === '1';
if(!demo){try{document.documentElement.dataset.theme=localStorage.getItem('vp-production-theme')||'light';}catch{document.documentElement.dataset.theme='light';}}
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><LoadingProvider><Suspense fallback={<LoadingFallback/>}>{demo ? <Demo /> : <Production />}</Suspense></LoadingProvider></React.StrictMode>);
