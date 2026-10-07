import React, { lazy, Suspense } from 'react';
import ReactDOM from 'react-dom/client';
import './styles.css';
import './layout.css';
import './live.css';
const Production = lazy(() => import('./Production'));
const Demo = lazy(() => import('./DemoEntry'));
const demo = new URLSearchParams(location.search).get('demo') === '1';
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><Suspense fallback={<div className="boot-screen" role="status">Carregando Verifica Placa…</div>}>{demo ? <Demo /> : <Production />}</Suspense></React.StrictMode>);
