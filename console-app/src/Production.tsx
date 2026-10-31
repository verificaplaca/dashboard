import { mobileLayoutQuery } from './lib/responsive';
import { lazy, Suspense, useEffect, useRef, useState } from 'react';
import { Activity, ArrowUpRight, CheckCircle2, ChevronsLeft, CreditCard, Crosshair, Wallet, Gauge, Headphones, Home, LayoutDashboard, LogIn, Megaphone, MessageSquare, Moon, PanelLeftClose, Search, Settings2, ShieldCheck, Sparkles, Sun, Target, TrendingUp, Users, Workflow, X } from 'lucide-react';
import { Avatar, Card, Dialog, Empty, PageHeading } from './components';
import { go, useRoute } from './format';
import { AuthProvider, useAuth } from './Auth';
import { LoadingFallback, useGlobalLoading } from './Loading';
import { UIProvider, useUI } from './ui';
import { observeEntrances } from './motion';
import './home.css';
import './users.css';
import './redesign.css';
import './modules-redesign.css';
import './control-alignment.css';
import './mobile-ux.css';
import {lockBodyScroll} from './lib/scroll-lock';
const UserManagement = lazy(() => import('./UserManagement'));
const Finance = lazy(() => import('./Finance'));
const Simulation = lazy(() => import('./Simulation'));
const Dashboard = lazy(() => import('./Dashboard'));
const OperationsHome = lazy(() => import('./OperationsHome'));
const Campaigns = lazy(() => import('./LiveCampaigns'));
const Targets = lazy(() => import('./Targets'));
const Tracking = lazy(() => import('./LiveTracking'));
const Profile = lazy(() => import('./LiveProfile'));
const Tasks = lazy(() => import('./LiveTasks'));
const groups = [
  { label: 'PRINCIPAL', items: [{ path: '/inicio', label: 'Início', icon: Home, live: true }, { path: '/dashboard', label: 'Visão Geral', icon: LayoutDashboard, live: true }, { path: '/financeiro', label: 'Financeiro', icon: Wallet, live: true }, { path: '/simulador', label: 'Simulador', icon: Gauge, live: true }] },
  { label: 'RESULTADOS', items: [{ path: '/marketing', label: 'Campanhas', icon: Megaphone, live: true }, { path: '/tracking', label: 'Tracking', icon: TrendingUp, live: true }, { path: '/metas', label: 'Metas mensais', icon: Target, live: true }] },
  { label: 'ADMINISTRAÇÃO', items: [{ path: '/usuarios', label: 'Usuários', icon: Users, live: true }] },
  { label: 'EXECUÇÃO', items: [{ path: '/tarefas', label: 'Tarefas', icon: CheckCircle2, live: true }] },
  { label: 'OPERAÇÃO · PRÓXIMA ETAPA', items: [{ path: '/crm', label: 'CRM e funis', icon: Workflow, live: false }, { path: '/clientes', label: 'Clientes', icon: Users, live: false }, { path: '/vendas', label: 'Vendas', icon: CreditCard, live: false }, { path: '/consultas', label: 'Consultas', icon: Crosshair, live: false }, { path: '/atendimento', label: 'Caixa compartilhada', icon: MessageSquare, live: false }, { path: '/whatsapp', label: 'WhatsApp', icon: Activity, live: false }, { path: '/evolucao', label: 'Evolução da equipe', icon: Sparkles, live: false }, { path: '/suporte', label: 'Suporte interno', icon: Headphones, live: false }, { path: '/administracao', label: 'Administração', icon: Settings2, live: false }] },
];
const items = [...groups.flatMap(group => group.items), { path: '/perfil', label: 'Meu perfil', icon: Users, live: true }];
function demoLink() { return `${import.meta.env.BASE_URL}?demo=1#/inicio`; }

export default function Production() {
  const [theme, setTheme] = useState(() => { try { return localStorage.getItem('vp-production-theme') || 'light'; } catch { return 'light'; } });
  const [toast, notify] = useState('');
  useEffect(() => { document.documentElement.dataset.theme = theme; try { localStorage.setItem('vp-production-theme', theme); } catch {} }, [theme]);
  useEffect(() => { if (!toast) return; const timer = setTimeout(() => notify(''), 5000); return () => clearTimeout(timer); }, [toast]);
  return <UIProvider theme={theme} setTheme={setTheme} notify={notify}><AuthProvider><AccessGate />{toast && <div className="toast" role="status"><CheckCircle2 size={18} /><span>{toast}</span><button aria-label="Fechar aviso" onClick={() => notify('')}><X size={14} /></button></div>}</AuthProvider></UIProvider>;
}
function AccessGate() {
  const {status,openLogin,refresh,signOut,user}=useAuth();
  useGlobalLoading(status==='loading');
  if(status==='ready') return <Shell />;
  return <main className="access-gate"><Card title="Verifica Placa · Acesso ao painel">
    {status==='loading'?<p role="status">Verificando seu acesso…</p>:<>
    <p role={status==='error'?'alert':undefined}>{status==='denied'?'Sua conta não tem acesso ativo ao painel. Procure um administrador.':status==='error'?'Não foi possível verificar sua permissão. Tente novamente.':'Entre com sua conta para acessar o painel.'}</p>
    {status==='signed-out'?<button className="button primary" onClick={openLogin}>Entrar no painel</button>:<button className="button primary" onClick={()=>void refresh()}>Verificar novamente</button>}
    {user&&<button className="button secondary" onClick={()=>void signOut()}>Sair da conta</button>}</>}
  </Card></main>;
}
function Shell() {
  const route = useRoute().split('?')[0], { theme, setTheme } = useUI(), { user, role, openLogin } = useAuth();
  const [collapsed, setCollapsed] = useState(true), [mobile, setMobile] = useState(false), [search, setSearch] = useState(false), [query, setQuery] = useState('');
  const [smallScreen,setSmallScreen]=useState(()=>matchMedia(mobileLayoutQuery).matches);
  const menuRef=useRef<HTMLElement>(null);
  useEffect(()=>{
    const media=matchMedia(mobileLayoutQuery);
    const closeWide=()=>{setSmallScreen(media.matches);if(!media.matches)setMobile(false);};
    media.addEventListener('change',closeWide);return()=>media.removeEventListener('change',closeWide);
  },[]);
  useEffect(()=>{
    if(!mobile)return;
    const release=lockBodyScroll(),content=document.querySelector<HTMLElement>('.app-content'),previous=document.activeElement as HTMLElement|null;
    const wasInert=content?.inert;if(content)content.inert=true;
    const focusables=()=>Array.from(menuRef.current?.querySelectorAll<HTMLElement>('button:not([disabled]),a[href]')||[]).filter(element=>element.getClientRects().length);
    const keydown=(event:KeyboardEvent)=>{
      if(event.key==='Escape'){event.preventDefault();setMobile(false);}
      if(event.key==='Tab'){const nodes=focusables(),first=nodes[0],last=nodes.at(-1);if(!first)return;if(!menuRef.current?.contains(document.activeElement)){event.preventDefault();first.focus();}else if(event.shiftKey&&document.activeElement===first){event.preventDefault();last?.focus();}else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus();}}
    };
    document.addEventListener('keydown',keydown);focusables()[0]?.focus();
    return()=>{document.removeEventListener('keydown',keydown);if(content)content.inert=wasInert??false;release();if(previous?.isConnected)previous.focus({preventScroll:true});};
  },[mobile]);
  useEffect(()=>{
    if(!smallScreen)return;
    const labelTables=()=>document.querySelectorAll<HTMLElement>('.table-scroll').forEach(wrapper=>{
      if(wrapper.hasAttribute('tabindex'))return;
      wrapper.tabIndex=0;wrapper.setAttribute('role','region');
      const title=wrapper.querySelector('caption')?.textContent?.trim()||wrapper.closest('.card')?.querySelector('h2,h3')?.textContent?.trim()||'Tabela de dados';
      if(!wrapper.hasAttribute('aria-label'))wrapper.setAttribute('aria-label',title+' · deslize para consultar as colunas');
    });
    labelTables();const observer=new MutationObserver(labelTables);
    observer.observe(document.body,{childList:true,subtree:true});return()=>observer.disconnect();
  },[smallScreen]);
  const allowed = (item: {path:string}) => item.path !== '/usuarios' || role === 'administrator';
  const searchResults=items.filter(allowed).filter(item=>item.label.toLocaleLowerCase('pt-BR').includes(query.toLocaleLowerCase('pt-BR')));
  const current = items.find(item => item.path === route);
  const [overviewSource, setOverviewSource] = useState('CARREGANDO FONTE');
  useEffect(() => {
    setOverviewSource('CARREGANDO FONTE');
    const update = (event: Event) => setOverviewSource((event as CustomEvent<string>).detail);
    window.addEventListener('vp-overview-source', update);
    return () => window.removeEventListener('vp-overview-source', update);
  }, [route]);
  useEffect(() => { setMobile(false); window.scrollTo({ top: 0 }); const main = document.querySelector<HTMLElement>('main'); if (main) return observeEntrances(main); }, [route]);
  useEffect(() => { const handler = (event: KeyboardEvent) => { if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') { event.preventDefault(); setMobile(false); setSearch(true); } }; window.addEventListener('keydown', handler); return () => window.removeEventListener('keydown', handler); }, []);
  const name = typeof user?.user_metadata?.full_name === 'string' ? user.user_metadata.full_name : 'Minha conta';
  return <div className={`app-shell premium-console module-design ${collapsed ? 'collapsed' : ''} ${mobile ? 'mobile-open' : ''}`}>
    <button className="skip-link" onClick={() => document.getElementById('main')?.focus()}>Pular para o conteúdo</button>
    {mobile && <button className="sidebar-shade" aria-label="Fechar menu" onClick={() => setMobile(false)} />}
    <aside ref={menuRef} id="operation-menu" className="sidebar" inert={smallScreen&&!mobile} aria-hidden={smallScreen&&!mobile||undefined} role={mobile?'dialog':undefined} aria-modal={mobile||undefined} aria-label={mobile?'Menu da operação':undefined}><div className="mobile-menu-heading"><strong>Verifica Placa</strong><button className="icon-button" aria-label="Fechar menu" onClick={()=>setMobile(false)}><X size={20}/></button></div><button className="brand" onClick={() => go('/inicio')} title="Verifica Placa" aria-label="Verifica Placa · Início"><span className="brand-mark"><img src={`${import.meta.env.BASE_URL}verifica-placa-logo.png`} alt="" /></span></button><div className="workspace-label"><ShieldCheck size={18} /><span>Verifica Placa<small>Dados da operação</small></span></div>
      <nav aria-label="Navegação principal">{groups.filter(group => group.items.some(item => item.live && allowed(item))).map(group => <div className="nav-group" key={group.label}><div className="nav-label">{group.label}</div>{group.items.filter(item => item.live && allowed(item)).map(item => <button key={item.path} onClick={() => {setMobile(false);go(item.path);}} title={`${item.label}${item.live ? '' : ' · em integração'}`} aria-current={route === item.path ? 'page' : undefined} className={`nav-item ${route === item.path ? 'active' : ''} ${item.live ? '' : 'planned-module'}`}><item.icon size={19} /><span>{item.label}</span>{!item.live && <i className="planned-dot" />}</button>)}</div>)}</nav>
      <div className="sidebar-bottom"><button className="rail-more" aria-label="Todos os módulos" title="Todos os módulos" onClick={() => {setMobile(false);setSearch(true);}}><Settings2 size={19} /></button><button className="rail-account" aria-label={user ? 'Meu perfil' : 'Entrar na conta'} title={user ? 'Meu perfil' : 'Entrar na conta'} onClick={() => {setMobile(false);user ? go('/perfil') : openLogin();}}><Avatar name={name} /></button><button className="collapse-button" onClick={() => setCollapsed(!collapsed)} aria-label={collapsed ? 'Expandir menu' : 'Recolher menu'}><ChevronsLeft size={17} /><span>Recolher menu</span></button></div>
    </aside>
    <div className="app-content"><header className="topbar"><div className="breadcrumb"><button className="icon-button mobile-menu" aria-label="Abrir menu" aria-expanded={mobile} aria-controls="operation-menu" onClick={() => setMobile(true)}><PanelLeftClose size={19} /></button><span className="mobile-route">{current?.label || 'Página'}</span></div><nav className="primary-tabs" aria-label="Áreas principais">{items.filter(item => ['/dashboard', '/financeiro', '/simulador', '/marketing'].includes(item.path)).map(item => <button key={item.path} onClick={() => go(item.path)} aria-current={route === item.path ? 'page' : undefined}>{item.label}</button>)}</nav><div className="topbar-actions"><button className="global-search" aria-label="Buscar módulo" onClick={() => setSearch(true)}><Search size={15} /><span>Buscar módulo</span><kbd>⌘ K</kbd></button><span className="demo-chip" role="status"><i />{route === '/dashboard' ? overviewSource : current?.live ? 'FONTES DA OPERAÇÃO' : 'EM INTEGRAÇÃO'}</span><button className="icon-button" aria-label={theme === 'dark' ? 'Ativar tema claro' : 'Ativar tema escuro'} onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}>{theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}</button>{user ? <button className="profile-trigger" aria-label="Meu perfil" onClick={() => go('/perfil')}><Avatar name={name} /><span>{name.split(' ')[0]}<small>{role === 'administrator' ? 'Administrador' : 'Gerente'}</small></span></button> : <button className="button secondary" onClick={openLogin}><LogIn size={15} />Entrar</button>}</div></header>
      <main id="main" tabIndex={-1} className="main-content" data-module={route.slice(1)}><Suspense fallback={<LoadingFallback/>}>{route === '/usuarios' ? role === 'administrator' ? <UserManagement /> : <Empty title="Acesso restrito" description="Somente administradores podem gerenciar usuários." /> : route === '/inicio' ? <OperationsHome /> : route === '/dashboard' ? <Dashboard /> : route === '/financeiro' ? <Finance /> : route === '/simulador' ? <Simulation /> : route === '/marketing' ? <Campaigns /> : route === '/tracking' ? <Tracking /> : route === '/metas' ? <Targets /> : route === '/perfil' ? <Profile /> : route === '/tarefas' ? <Tasks /> : current ? <IntegrationStage title={current.label} /> : <Empty title="Página não encontrada" description="Escolha um módulo no menu."><button className="button primary" onClick={() => go('/inicio')}>Voltar ao início</button></Empty>}</Suspense>
        <footer className="app-footer"><span><ShieldCheck size={13} />Verifica Placa · Console</span><span>{current?.live ? 'Fontes reais · Supabase existente' : 'Módulo em integração'}</span><a href={demoLink()}>Explorar demonstração<ArrowUpRight size={12} /></a></footer>
      </main>
    </div>
    {search && <Dialog title="Buscar módulo" close={() => setSearch(false)}><label className="search-field"><Search size={18} /><input autoFocus value={query} onChange={event => setQuery(event.target.value)} placeholder="Visão Geral, campanhas, metas…" /></label><div className="search-results">{searchResults.map(item => <button key={item.path} onClick={() => { go(item.path); setSearch(false); setQuery(''); }}><item.icon size={17} /><span>{item.label}<small>{item.live ? 'Dados reais' : 'Em integração'}</small></span><ArrowUpRight size={15} /></button>)}{!searchResults.length&&<p className="dialog-description" role="status">Nenhum módulo encontrado. Tente outro nome.</p>}</div></Dialog>}
  </div>;
}
function IntegrationStage({ title }: { title: string }) {
  return <><PageHeading eyebrow="OPERAÇÃO · PRÓXIMA ETAPA" title={title} description="A experiência está pronta para demonstração. A operação real deste módulo ainda precisa ser conectada." /><Card><Empty title="Integração em preparação" description="Os fluxos desta área estão disponíveis com dados fictícios na demonstração."><a className="button primary" href={demoLink()}>Explorar demonstração<ArrowUpRight size={15} /></a><button className="button secondary" onClick={() => go('/dashboard')}>Ver dados da operação</button></Empty></Card></>;
}
