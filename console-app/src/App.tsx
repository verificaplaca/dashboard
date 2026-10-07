import { observeEntrances } from './motion';
import { useEffect, useRef, useState } from 'react';
import { Activity, ArrowUpRight, Bell, BookOpen, ChevronDown, ChevronRight, ChevronsLeft, Command, CreditCard, Crosshair, Gauge, Headphones, LayoutDashboard, LogOut, Megaphone, MessageSquare, Moon, PanelLeftClose, Plus, Search, Settings2, Home as HomeIcon, ShieldCheck, Sparkles, Sun, Target, TrendingUp, Users, Workflow, X, CheckCircle2 } from 'lucide-react';
import { Avatar, Badge, Dialog, Drawer, Empty, FormDialog, Progress } from './components';
import { date, go, roleNames, useDemo, useRoute } from './context';
import Dashboard from './Dashboard';
import Home from './Home';
import Inbox from './Inbox';
import Workspaces from './Workspaces';
import { scopeOwner, scopedClients } from './Commerce';

const groups = [
  { label: 'PRINCIPAL', items: [{ path: '/inicio', label: 'Início', icon: HomeIcon, permission: '' }, { path: '/dashboard', label: 'Visão Geral', icon: LayoutDashboard, permission: 'view.dashboard' }] },
  { label: 'COMERCIAL', items: [{ path: '/crm', label: 'CRM e funis', icon: Workflow, permission: 'view.crm' }, { path: '/clientes', label: 'Clientes', icon: Users, permission: 'view.clients' }, { path: '/vendas', label: 'Vendas', icon: CreditCard, permission: 'view.sales' }, { path: '/consultas', label: 'Consultas', icon: Crosshair, permission: 'view.consultations' }] },
  { label: 'ATENDIMENTO', items: [{ path: '/atendimento', label: 'Caixa compartilhada', icon: MessageSquare, permission: 'view.inbox' }, { path: '/whatsapp', label: 'WhatsApp', icon: Activity, permission: 'view.whatsapp' }] },
  { label: 'CRESCIMENTO', items: [{ path: '/marketing', label: 'Marketing', icon: Megaphone, permission: 'view.marketing' }, { path: '/tracking', label: 'Tracking', icon: TrendingUp, permission: 'view.tracking' }] },
  { label: 'EQUIPE', items: [{ path: '/tarefas', label: 'Tarefas', icon: CheckCircle2, permission: 'view.tasks' }, { path: '/evolucao', label: 'Metas e evolução', icon: Target, permission: 'view.evolution' }, { path: '/suporte', label: 'Suporte interno', icon: Headphones, permission: 'view.support' }] },
  { label: 'GESTÃO', items: [{ path: '/relatorios', label: 'Relatórios', icon: BookOpen, permission: 'view.reports' }, { path: '/administracao', label: 'Administração', icon: Settings2, permission: 'view.admin' }] }
];
const items = groups.flatMap(g => g.items);

export default function App() {
  const { state, user, can, run, toast, notify } = useDemo();
  const route = useRoute();
  const [pathname, queryString] = route.split('?');
  const parts = pathname.split('/').filter(Boolean), root = `/${parts[0] || 'inicio'}`;
  useEffect(() => { const content = document.querySelector<HTMLElement>('main'); if (content) return observeEntrances(content); }, [root]);
  const current = items.find(i => i.path === root);
  const [collapsed, setCollapsed] = useState(() => localStorage.getItem('vp-sidebar') === 'collapsed');
  const [mobile, setMobile] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [notifications, setNotifications] = useState(false);
  const [profileOpen, setProfileOpen] = useState(false);
  const profileMenuRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!profileOpen) return;
    const dismiss = (event: PointerEvent) => { if (!profileMenuRef.current?.contains(event.target as Node)) setProfileOpen(false); };
    const escape = (event: KeyboardEvent) => { if (event.key === 'Escape') { setProfileOpen(false); profileMenuRef.current?.querySelector<HTMLButtonElement>('button')?.focus(); } };
    document.addEventListener('pointerdown', dismiss); document.addEventListener('keydown', escape);
    return () => { document.removeEventListener('pointerdown', dismiss); document.removeEventListener('keydown', escape); };
  }, [profileOpen]);
  const [quickTask, setQuickTask] = useState(false);
  useEffect(() => { document.documentElement.dataset.theme = state.settings.theme; document.documentElement.dataset.compact = String(state.settings.compact); }, [state.settings]);
  useEffect(() => { localStorage.setItem('vp-sidebar', collapsed ? 'collapsed' : 'expanded'); }, [collapsed]);
  useEffect(() => { setMobile(false); setProfileOpen(false); window.scrollTo({ top: 0 }); }, [root]);
  useEffect(() => { const handler = (e: KeyboardEvent) => { if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); setSearchOpen(true); } }; window.addEventListener('keydown', handler); return () => window.removeEventListener('keydown', handler); }, []);
  const count = state.conversations.filter(c => c.unread && c.status !== 'resolved' && (['owner','admin'].includes(user.role) || c.ownerId === user.id || state.users.find(u => u.id === c.ownerId)?.teamId === user.teamId || (!c.ownerId && state.instances.find(i => i.id === c.instanceId)?.teamId === user.teamId))).length;
  const hasAccess = (permission: string) => !permission || can(permission);
  const allowed = root === '/perfil' || (current && hasAccess(current.permission));
  const results = [
    ...items.filter(i => hasAccess(i.permission)).map(i => ({ label: i.label, sub: 'Módulo', path: i.path })),
    ...(can('view.clients') ? scopedClients(state, user).map(c => ({ label: c.name, sub: `Cliente · ${c.company}`, path: `/clientes/${c.id}` })) : []),
    ...(can('view.sales') ? state.orders.filter(o => scopeOwner(state, user, o.assistedBy)).map(o => ({ label: o.id, sub: `Pedido · ${state.clients.find(c => c.id === o.clientId)?.name}`, path: `/vendas/${o.id}` })) : [])
  ].filter(r => `${r.label} ${r.sub}`.toLowerCase().includes(query.toLowerCase())).slice(0, 9);
  const nav = (path: string) => { go(path); setSearchOpen(false); setQuery(''); };
  const activeGoals = state.goals.filter(g => g.scope === 'organization');
  const orgGoal = activeGoals[0];

  return <div className={`app-shell ${collapsed ? 'collapsed' : ''} ${mobile ? 'mobile-open' : ''}`}>
    <button className="skip-link" onClick={() => document.getElementById('main')?.focus()}>Pular para o conteúdo</button>
    {mobile && <button className="sidebar-shade" aria-label="Fechar menu" onClick={() => setMobile(false)} />}
    <aside className="sidebar">
      <button className="brand" onClick={() => go('/inicio')} title="Verifica Placa"><span className="brand-mark"><img src={`${import.meta.env.BASE_URL}verifica-placa-icon.jpg`} alt="" /></span><span className="brand-text">verifica<span>placa</span><small>BUSINESS CONSOLE</small></span></button>
      <div className="workspace-label"><span className="workspace-icon"><img src={`${import.meta.env.BASE_URL}verifica-placa-icon.jpg`} alt="" /></span><span>Verifica Placa<small>Workspace principal</small></span><ChevronDown size={14} /></div>
      <nav aria-label="Navegação principal">{groups.map(group => group.items.some(i => hasAccess(i.permission)) && <div className="nav-group" key={group.label}><div className="nav-label">{group.label}</div>{group.items.filter(i => hasAccess(i.permission)).map(item => <button key={item.path} onClick={() => go(item.path)} title={item.label} aria-current={root === item.path ? 'page' : undefined} className={`nav-item ${root === item.path ? 'active' : ''}`}><item.icon size={18} /><span>{item.label}</span>{item.path === '/atendimento' && count > 0 && <em>{count}</em>}{item.path === '/evolucao' && <Sparkles className="nav-spark" size={13} />}</button>)}</div>)}</nav>
      <div className="sidebar-bottom">{root !== '/dashboard' && root !== '/inicio' && <div className="sidebar-goal"><span><Target size={15} />FOCO DA EQUIPE</span><strong>{orgGoal?.label || 'Evoluir juntos'}</strong><div><small>Progresso do ciclo</small><b>{orgGoal ? Math.round(orgGoal.current / orgGoal.target * 100) : 0}%</b></div><Progress value={orgGoal ? orgGoal.current / orgGoal.target * 100 : 0} />{can('view.evolution') && <button onClick={() => go('/evolucao')}>Ver metas <ArrowUpRight size={14} /></button>}</div>}<button className="collapse-button" onClick={() => setCollapsed(!collapsed)} aria-label={collapsed ? 'Expandir menu' : 'Recolher menu'}><ChevronsLeft size={17} /><span>Recolher menu</span></button></div>
    </aside>
    <div className="app-content">
      <header className="topbar"><div className="breadcrumb"><button className="icon-button mobile-menu" aria-label="Abrir menu" onClick={() => setMobile(true)}><PanelLeftClose size={19} /></button><span className="breadcrumb-workspace">Workspace</span><ChevronRight size={13} /><span>{current?.label || (root === '/perfil' ? 'Meu perfil' : 'Página')}</span></div><div className="topbar-actions"><button className="global-search" aria-label="Buscar no workspace" onClick={() => setSearchOpen(true)}><Search size={15} /><span>Buscar no workspace</span><kbd>⌘ K</kbd></button><span className="demo-chip"><i />{root === '/dashboard' ? 'FONTES ORIGINAIS' : root === '/inicio' ? 'WORKSPACE' : 'DEMONSTRAÇÃO'}</span><button className="icon-button" aria-label={state.settings.theme === 'dark' ? 'Ativar tema claro' : 'Ativar tema escuro'} onClick={() => run('settings.update', { theme: state.settings.theme === 'dark' ? 'light' : 'dark' }, 'Tema atualizado.')} >{state.settings.theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}</button><button className="icon-button notification-button" aria-label="Abrir notificações" onClick={() => setNotifications(true)}><Bell size={18} /><i /></button><div className="profile-menu-wrap" ref={profileMenuRef}><button className="profile-trigger" aria-label="Menu do perfil e simulação de acesso" aria-expanded={profileOpen} onClick={() => setProfileOpen(!profileOpen)}><Avatar name={user.name} src={user.avatar} /><span>{user.name.split(' ')[0]}<small>{roleNames[user.role]}</small></span><ChevronDown size={14} /></button>{profileOpen && <div className="profile-menu"><button onClick={() => { setProfileOpen(false); go('/perfil'); }}>Meu perfil <Users size={15} /></button><div className="profile-menu-label">SIMULAR PERFIL</div>{state.users.map(u => <button key={u.id} disabled={!u.active} className={user.id === u.id ? 'selected' : ''} onClick={() => { run('demo.switchUser', { id: u.id }, `Perfil demonstrativo: ${u.name}`); setProfileOpen(false); go('/inicio'); }}>{u.name}<small>{roleNames[u.role]}{!u.active ? ' · Suspenso' : ''}</small></button>)}</div>}</div></div></header>
      <main id="main" tabIndex={-1} className={`main-content ${root === '/atendimento' ? 'inbox-main' : ''}`}>
        {allowed ? root === '/inicio' ? <Home /> : root === '/dashboard' ? <Dashboard /> : root === '/atendimento' ? <Inbox conversationId={parts[1]} /> : <Workspaces page={root.slice(1)} id={parts[1]} query={new URLSearchParams(queryString)} /> : <Empty title={current ? 'Acesso restrito ao seu perfil' : 'Página não encontrada'} description={current ? 'Escolha um módulo autorizado no menu ou simule outro perfil no canto superior.' : 'Este endereço não corresponde a uma tela do workspace.'}><button className="button secondary" onClick={() => go('/inicio')}>Voltar ao início</button></Empty>}
        <footer className="app-footer"><a href={`${import.meta.env.BASE_URL}#/inicio`}>Sair da demonstração</a><span><ShieldCheck size={13} />Verifica Placa · Business Console</span><span>{root === '/dashboard' ? 'Visão geral · leitura das fontes originais' : root === '/inicio' ? 'Financeiro: fontes originais · Operação: demonstração' : 'Dados fictícios · Referência: 06 out 2026'}</span>{can('view.support') && <button onClick={() => go('/suporte')}>Suporte interno <ArrowUpRight size={12} /></button>}</footer>
      </main>
    </div>
    {toast && <div className="toast" role="status"><CheckCircle2 size={18} /><span>{toast}</span><button aria-label="Fechar aviso" onClick={() => notify('')}><X size={14} /></button></div>}
    {searchOpen && <Dialog title="Buscar no workspace" close={() => setSearchOpen(false)}><label className="search-field search-dialog-input"><Search size={20} /><input autoFocus value={query} onChange={e => setQuery(e.target.value)} placeholder="Cliente, pedido ou módulo..." /></label><div className="search-results">{results.map(r => <button key={r.path} onClick={() => nav(r.path)}><Search size={16} /><span>{r.label}<small>{r.sub}</small></span><ArrowUpRight size={16} /></button>)}{results.length === 0 && <Empty description="Tente buscar pelo nome de um cliente ou código de pedido." />}</div></Dialog>}
    {notifications && <Drawer title="Central de notificações" close={() => setNotifications(false)}><p className="drawer-description">Pendências que precisam da sua atenção.</p>{can('view.consultations') && state.orders.filter(o => o.consultationStatus === 'failed').slice(0, 3).map(o => <button className="notification-row" key={o.id} onClick={() => { setNotifications(false); go(`/consultas/${o.id}`); }}><Crosshair size={18} /><span>Consulta precisa de revisão<small>{o.id} · {o.plate}</small></span><ChevronRight size={15} /></button>)}{can('view.whatsapp') && state.instances.filter(i => i.status === 'disconnected').map(i => <button className="notification-row" key={i.id} onClick={() => { setNotifications(false); go('/whatsapp'); }}><Activity size={18} /><span>{i.name} desconectada<small>Verificar conexão e fila de envio</small></span><ChevronRight size={15} /></button>)}{can('view.tasks') && <button className="notification-row" onClick={() => { setNotifications(false); go('/tarefas'); }}><CheckCircle2 size={18} /><span>{state.tasks.filter(t => t.status !== 'done' && scopeOwner(state,user,t.ownerId)).length} tarefas em aberto<small>Organize os próximos passos</small></span><ChevronRight size={15} /></button>}</Drawer>}
    {quickTask && <FormDialog title="Nova tarefa" close={() => setQuickTask(false)} fields={[{ name: 'title', label: 'Título', required: true }, { name: 'dueDate', label: 'Prazo', type: 'date', value: '2026-10-07', required: true }]} onSubmit={v => !!run('task.create', { ...v, clientId: null, ownerId: user.id, priority: 'normal' })} />}
  </div>;
}
