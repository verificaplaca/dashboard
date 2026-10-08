import { lazy, Suspense, useEffect, useState } from 'react';
import { Activity, ArrowUpRight, CalendarDays, CheckCircle2, ChevronsLeft, CreditCard, Crosshair, Headphones, Home, LayoutDashboard, LogIn, Megaphone, MessageSquare, Moon, PanelLeftClose, Search, Settings2, ShieldCheck, Sparkles, Sun, Target, TrendingUp, Users, Workflow, X } from 'lucide-react';
import { Avatar, Card, Dialog, Empty, PageHeading } from './components';
import { go, useRoute } from './format';
import { AuthProvider, useAuth } from './Auth';
import { UIProvider, useUI } from './ui';
import { observeEntrances } from './motion';
import { greeting, saoPauloClock } from './domain/clock';
import './home.css';
const Dashboard = lazy(() => import('./Dashboard'));
const ExecutiveSummary = lazy(() => import('./ExecutiveSummary'));
const Campaigns = lazy(() => import('./LiveCampaigns'));
const Targets = lazy(() => import('./Targets'));
const Tracking = lazy(() => import('./LiveTracking'));
const Profile = lazy(() => import('./LiveProfile'));
const Tasks = lazy(() => import('./LiveTasks'));
const HomeTasks = lazy(() => import('./TaskHomeSummary'));
const groups = [
  { label: 'PRINCIPAL', items: [{ path: '/inicio', label: 'Início', icon: Home, live: true }, { path: '/dashboard', label: 'Visão Geral', icon: LayoutDashboard, live: true }] },
  { label: 'RESULTADOS', items: [{ path: '/marketing', label: 'Campanhas', icon: Megaphone, live: true }, { path: '/tracking', label: 'Tracking', icon: TrendingUp, live: true }, { path: '/metas', label: 'Metas mensais', icon: Target, live: true }] },
  { label: 'EXECUÇÃO', items: [{ path: '/tarefas', label: 'Tarefas', icon: CheckCircle2, live: true }] },
  { label: 'OPERAÇÃO · PRÓXIMA ETAPA', items: [{ path: '/crm', label: 'CRM e funis', icon: Workflow, live: false }, { path: '/clientes', label: 'Clientes', icon: Users, live: false }, { path: '/vendas', label: 'Vendas', icon: CreditCard, live: false }, { path: '/consultas', label: 'Consultas', icon: Crosshair, live: false }, { path: '/atendimento', label: 'Caixa compartilhada', icon: MessageSquare, live: false }, { path: '/whatsapp', label: 'WhatsApp', icon: Activity, live: false }, { path: '/evolucao', label: 'Evolução da equipe', icon: Sparkles, live: false }, { path: '/suporte', label: 'Suporte interno', icon: Headphones, live: false }, { path: '/administracao', label: 'Administração', icon: Settings2, live: false }] },
];
const items = [...groups.flatMap(group => group.items), { path: '/perfil', label: 'Meu perfil', icon: Users, live: true }];
function demoLink() { return `${import.meta.env.BASE_URL}?demo=1#/inicio`; }

export default function Production() {
  const [theme, setTheme] = useState(() => { try { return localStorage.getItem('vp-production-theme') || 'dark'; } catch { return 'dark'; } });
  const [toast, notify] = useState('');
  useEffect(() => { document.documentElement.dataset.theme = theme; try { localStorage.setItem('vp-production-theme', theme); } catch {} }, [theme]);
  useEffect(() => { if (!toast) return; const timer = setTimeout(() => notify(''), 5000); return () => clearTimeout(timer); }, [toast]);
  return <UIProvider theme={theme} setTheme={setTheme} notify={notify}><AuthProvider><Shell />{toast && <div className="toast" role="status"><CheckCircle2 size={18} /><span>{toast}</span><button aria-label="Fechar aviso" onClick={() => notify('')}><X size={14} /></button></div>}</AuthProvider></UIProvider>;
}
function Shell() {
  const route = useRoute().split('?')[0], { theme, setTheme } = useUI(), { user, openLogin } = useAuth();
  const [collapsed, setCollapsed] = useState(false), [mobile, setMobile] = useState(false), [search, setSearch] = useState(false), [query, setQuery] = useState('');
  const current = items.find(item => item.path === route);
  useEffect(() => { setMobile(false); window.scrollTo({ top: 0 }); const main = document.querySelector<HTMLElement>('main'); if (main) return observeEntrances(main); }, [route]);
  useEffect(() => { const handler = (event: KeyboardEvent) => { if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') { event.preventDefault(); setSearch(true); } }; window.addEventListener('keydown', handler); return () => window.removeEventListener('keydown', handler); }, []);
  const name = typeof user?.user_metadata?.full_name === 'string' ? user.user_metadata.full_name : 'Minha conta';
  return <div className={`app-shell ${collapsed ? 'collapsed' : ''} ${mobile ? 'mobile-open' : ''}`}>
    <button className="skip-link" onClick={() => document.getElementById('main')?.focus()}>Pular para o conteúdo</button>
    {mobile && <button className="sidebar-shade" aria-label="Fechar menu" onClick={() => setMobile(false)} />}
    <aside className="sidebar"><button className="brand" onClick={() => go('/inicio')} title="Verifica Placa"><span className="brand-mark"><img src={`${import.meta.env.BASE_URL}verifica-placa-icon.jpg`} alt="" /></span><span className="brand-text">verifica<span>placa</span><small>BUSINESS CONSOLE</small></span></button><div className="workspace-label"><ShieldCheck size={18} /><span>Verifica Placa<small>Dados da operação</small></span></div>
      <nav aria-label="Navegação principal">{groups.map(group => <div className="nav-group" key={group.label}><div className="nav-label">{group.label}</div>{group.items.map(item => <button key={item.path} onClick={() => go(item.path)} title={`${item.label}${item.live ? '' : ' · em integração'}`} aria-current={route === item.path ? 'page' : undefined} className={`nav-item ${route === item.path ? 'active' : ''} ${item.live ? '' : 'planned-module'}`}><item.icon size={18} /><span>{item.label}</span>{!item.live && <i className="planned-dot" />}</button>)}</div>)}</nav>
      <div className="sidebar-bottom"><button className="collapse-button" onClick={() => setCollapsed(!collapsed)} aria-label={collapsed ? 'Expandir menu' : 'Recolher menu'}><ChevronsLeft size={17} /><span>Recolher menu</span></button></div>
    </aside>
    <div className="app-content"><header className="topbar"><div className="breadcrumb"><button className="icon-button mobile-menu" aria-label="Abrir menu" onClick={() => setMobile(true)}><PanelLeftClose size={19} /></button><span>{current?.label || 'Página'}</span></div><div className="topbar-actions"><button className="global-search" aria-label="Buscar módulo" onClick={() => setSearch(true)}><Search size={15} /><span>Buscar módulo</span><kbd>⌘ K</kbd></button><span className="demo-chip"><i />DADOS REAIS</span><button className="icon-button" aria-label={theme === 'dark' ? 'Ativar tema claro' : 'Ativar tema escuro'} onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}>{theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}</button>{user ? <button className="profile-trigger" aria-label="Meu perfil" onClick={() => go('/perfil')}><Avatar name={name} /><span>{name.split(' ')[0]}<small>Conta autenticada</small></span></button> : <button className="button secondary" onClick={openLogin}><LogIn size={15} />Entrar</button>}</div></header>
      <main id="main" tabIndex={-1} className="main-content"><Suspense fallback={<Card><p role="status">Carregando módulo…</p></Card>}>{route === '/inicio' ? <LiveHome /> : route === '/dashboard' ? <Dashboard /> : route === '/marketing' ? <Campaigns /> : route === '/tracking' ? <Tracking /> : route === '/metas' ? <Targets /> : route === '/perfil' ? <Profile /> : route === '/tarefas' ? <Tasks /> : current ? <IntegrationStage title={current.label} /> : <Empty title="Página não encontrada" description="Escolha um módulo no menu."><button className="button primary" onClick={() => go('/inicio')}>Voltar ao início</button></Empty>}</Suspense>
        <footer className="app-footer"><span><ShieldCheck size={13} />Verifica Placa · Console</span><span>{current?.live ? 'Fontes reais · Supabase existente' : 'Módulo em integração'}</span><a href={demoLink()}>Explorar demonstração<ArrowUpRight size={12} /></a></footer>
      </main>
    </div>
    {search && <Dialog title="Buscar módulo" close={() => setSearch(false)}><label className="search-field"><Search size={18} /><input autoFocus value={query} onChange={event => setQuery(event.target.value)} placeholder="Visão Geral, campanhas, metas…" /></label><div className="search-results">{items.filter(item => item.label.toLocaleLowerCase('pt-BR').includes(query.toLocaleLowerCase('pt-BR'))).map(item => <button key={item.path} onClick={() => { go(item.path); setSearch(false); setQuery(''); }}><item.icon size={17} /><span>{item.label}<small>{item.live ? 'Dados reais' : 'Em integração'}</small></span><ArrowUpRight size={15} /></button>)}</div></Dialog>}
  </div>;
}
function LiveHome() {
  const { user } = useAuth();
  const [now, setNow] = useState(() => new Date());
  useEffect(() => { const timer = setInterval(() => setNow(new Date()), 60_000); return () => clearInterval(timer); }, []);
  const clock = saoPauloClock(now), name = typeof user?.user_metadata?.full_name === 'string' ? user.user_metadata.full_name.split(' ')[0] : '';
  const GreetingIcon = clock.hour < 5 || clock.hour >= 18 ? Moon : Sun;
  return <div className="home-page live-home"><div className="live-welcome card"><div><span className="eyebrow"><GreetingIcon size={16} />{new Intl.DateTimeFormat('pt-BR', { timeZone: 'America/Sao_Paulo', weekday: 'long', day: 'numeric', month: 'long' }).format(now)}</span><h1>{greeting(clock.hour)}{name ? `, ${name}` : ''}.</h1><p>Acompanhe o resultado da operação e prepare os próximos passos.</p><div className="live-welcome-actions"><button className="button primary" onClick={() => go('/dashboard')}>Abrir Visão Geral<ArrowUpRight size={15} /></button><button className="button secondary" onClick={() => go('/metas')}><CalendarDays size={15} />Revisar metas</button></div></div><div className="live-orbit" aria-hidden="true"><ShieldCheck size={42} /></div></div>
    <ExecutiveSummary />
    <HomeTasks />
    <div className="live-quick-grid">{[['/marketing', 'Campanhas', 'Investimento e performance de aquisição.'], ['/tracking', 'Tracking', 'Entregas de eventos aos canais de mídia.'], ['/metas', 'Metas mensais', 'Objetivos e imposto vigentes por mês.']].map(([path, title, description]) => <Card key={path} title={title}><p>{description}</p><button className="text-button" onClick={() => go(path)}>Acompanhar<ArrowUpRight size={15} /></button></Card>)}</div>
    <Card title="Seu próximo ciclo de evolução" subtitle="Metas individuais, desafios e conquistas fazem parte da próxima etapa da operação."><p>Explore os fluxos de equipe na demonstração enquanto conectamos as vendas, entregas e responsabilidades reais.</p><a className="button secondary" href={demoLink()}><Sparkles size={15} />Explorar metas e evolução</a></Card>
  </div>;
}
function IntegrationStage({ title }: { title: string }) {
  return <><PageHeading eyebrow="OPERAÇÃO · PRÓXIMA ETAPA" title={title} description="A experiência está pronta para demonstração. A operação real deste módulo ainda precisa ser conectada." /><Card><Empty title="Integração em preparação" description="Os fluxos desta área estão disponíveis com dados fictícios na demonstração."><a className="button primary" href={demoLink()}>Explorar demonstração<ArrowUpRight size={15} /></a><button className="button secondary" onClick={() => go('/dashboard')}>Ver dados da operação</button></Empty></Card></>;
}
