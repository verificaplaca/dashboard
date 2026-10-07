import { useEffect, useState } from 'react';
import { Activity, ArrowRight, ArrowUpRight, BookOpen, CalendarDays, Check, CheckCircle2, ChevronRight, Clock3, CreditCard, Crosshair, Headphones, LayoutDashboard, Megaphone, MessageSquare, Moon, Settings2, ShieldCheck, Sunrise, Sun, Target, TrendingUp, Users, WifiOff, Workflow, type LucideIcon } from 'lucide-react';
import { date, go, roleNames, useDemo } from './context';
import { scopeOwner, scopedClients } from './Commerce';
import ExecutiveSummary from './ExecutiveSummary';
import { canCompleteTask, greeting, inboxScope, individualProgress, progressGoals, saoPauloClock } from './domain/home';
import './home.css';

type Shortcut = { path: string; permission: string; label: string; detail: string; icon: LucideIcon; count?: number };
type PriorityItem = { key: string; label: string; detail: string; count: number; path: string; icon: LucideIcon; urgent?: boolean };
const focusChoices = [
  { id: 'plan', label: 'Organizar os próximos passos', permission: 'view.tasks', path: '/tarefas', description: 'Comece pelos prazos e dê um próximo passo a cada pendência.' },
  { id: 'relationship', label: 'Cuidar do relacionamento', permission: 'view.inbox', path: '/atendimento', description: 'Retome conversas com o contexto de cada cliente.' },
  { id: 'pipeline', label: 'Avançar as oportunidades', permission: 'view.crm', path: '/crm', description: 'Transforme cada oportunidade em uma ação comercial clara.' },
  { id: 'quality', label: 'Garantir a qualidade da entrega', permission: 'view.consultations', path: '/consultas', description: 'Revise falhas e acompanhe o caminho até a entrega.' },
  { id: 'growth', label: 'Acompanhar a aquisição', permission: 'view.marketing', path: '/marketing', description: 'Confira campanhas e os segmentos que precisam de atenção.' },
  { id: 'executive', label: 'Acompanhar os resultados', permission: 'view.dashboard', path: '/dashboard', description: 'Use a visão geral para orientar as próximas decisões.' },
];
function readFocus(userId: string) { try { return localStorage.getItem(`vp-home-focus-${userId}`) || ''; } catch { return ''; } }
function Bar({ value, label }: { value: number; label: string }) { return <div className="home-progress" role="progressbar" aria-label={label} aria-valuenow={Math.round(Math.max(0, Math.min(100, value)))} aria-valuemin={0} aria-valuemax={100}><span style={{ width: `${Math.max(0, Math.min(100, value))}%` }} /></div>; }

export default function Home() {
  const { state, user, can, run, notify } = useDemo();
  const [now, setNow] = useState(() => new Date());
  const [focusStore, setFocusStore] = useState(() => ({ userId: user.id, id: readFocus(user.id) }));
  useEffect(() => {
    const refresh = () => setNow(new Date());
    const onVisible = () => { if (document.visibilityState === 'visible') refresh(); };
    const timer = window.setInterval(refresh, 60_000);
    document.addEventListener('visibilitychange', onVisible);
    return () => { window.clearInterval(timer); document.removeEventListener('visibilitychange', onVisible); };
  }, []);
  useEffect(() => { setFocusStore({ userId: user.id, id: readFocus(user.id) }); }, [user.id]);
  const clock = saoPauloClock(now), salutation = greeting(clock.hour);
  const GreetingIcon = clock.hour < 5 || clock.hour >= 18 ? Moon : clock.hour < 12 ? Sunrise : Sun;
  const firstName = user.name.split(' ')[0];
  const lead = ['owner', 'admin', 'manager'].includes(user.role);
  const global = ['owner', 'admin'].includes(user.role);
  const clients = scopedClients(state, user);
  const clientIds = new Set(clients.map(c => c.id));
  const tasks = can('view.tasks') ? state.tasks.filter(t => scopeOwner(state, user, t.ownerId)) : [];
  const openTasks = tasks.filter(t => t.status !== 'done');
  const late = openTasks.filter(t => t.dueDate < clock.day);
  const due = openTasks.filter(t => t.dueDate === clock.day);
  const myTasks = tasks.filter(t => t.ownerId === user.id);
  const taskCompletion = myTasks.length ? myTasks.filter(t => t.status === 'done').length / myTasks.length * 100 : 0;
  const conversations = inboxScope(state, user);
  const pendingConversations = conversations.filter(c => c.status !== 'resolved' && (c.unread > 0 || c.status === 'waiting' || c.messages.some(m => m.kind === 'outgoing' && m.status === 'pending')));
  const orders = can('view.sales') ? state.orders.filter(o => scopeOwner(state, user, o.assistedBy) || (user.role === 'support' && clientIds.has(o.clientId))) : [];
  const consultations = can('view.consultations') ? state.orders.filter(o => o.paidAt && (scopeOwner(state, user, o.assistedBy) || user.role === 'support')) : [];
  const failed = consultations.filter(o => o.consultationStatus === 'failed');
  const instances = can('view.whatsapp') ? state.instances.filter(i => global || i.teamId === user.teamId) : [];
  const disconnected = instances.filter(i => i.status === 'disconnected');
  const opportunities = can('view.crm') ? state.opportunities.filter(o => scopeOwner(state, user, o.ownerId)) : [];
  const awaiting = opportunities.filter(o => o.stage === 'payment');
  const tickets = can('view.support') ? state.tickets.filter(t => scopeOwner(state, user, t.ownerId)) : [];
  const goals = progressGoals(state, user);
  const goalCards = [...goals].sort((a, b) => Number(a.scope === 'organization') - Number(b.scope === 'organization')).slice(0, 3);
  const personal = individualProgress(state, user.id);
  const teamGoal = goals.find(g => g.scope === 'team' && g.ownerId === user.teamId) || goals.find(g => g.scope === 'organization');
  const teamProgress = teamGoal && teamGoal.target > 0 ? Math.min(100, teamGoal.current / teamGoal.target * 100) : 0;
  const leadership = lead && personal.confirmed === 0;
  const focusOptions = focusChoices.filter(f => can(f.permission));
  const savedFocus = focusStore.userId === user.id ? focusStore.id : readFocus(user.id);
  const focus = focusOptions.find(f => f.id === savedFocus) || focusOptions[0];
  const setFocus = (id: string) => { setFocusStore({ userId: user.id, id }); try { localStorage.setItem(`vp-home-focus-${user.id}`, id); } catch { notify('O foco foi atualizado nesta sessão. O navegador não permitiu salvar.'); } };
  const priorities: PriorityItem[] = [
    { key: 'late', label: 'Prazos que merecem atenção', detail: 'Tarefas com prazo anterior a hoje.', count: late.length, path: '/tarefas', icon: Clock3, urgent: true },
    { key: 'failed', label: 'Revisar consultas com falha', detail: 'Confira o contexto antes de uma nova tentativa.', count: failed.length, path: failed[0] ? `/consultas/${failed[0].id}` : '/consultas', icon: Crosshair, urgent: true },
    { key: 'offline', label: 'Canais desconectados', detail: 'Acompanhe a conexão e a fila de envio.', count: disconnected.length, path: '/whatsapp', icon: WifiOff, urgent: true },
    { key: 'conversations', label: 'Conversas esperando um próximo passo', detail: 'Mensagens não lidas, espera ou envio pendente.', count: pendingConversations.length, path: pendingConversations[0] ? `/atendimento/${pendingConversations[0].id}` : '/atendimento', icon: MessageSquare },
    { key: 'today', label: 'Compromissos de hoje', detail: 'Tarefas com prazo no dia atual.', count: due.length, path: '/tarefas', icon: CalendarDays },
    { key: 'payment', label: 'Oportunidades aguardando pagamento', detail: 'Acompanhe o próximo passo do cliente.', count: awaiting.length, path: awaiting[0] ? `/crm/${awaiting[0].id}` : '/crm', icon: Workflow },
    { key: 'orders', label: 'Pedidos aguardando pagamento', detail: 'Acompanhe os pedidos ainda sem confirmação.', count: !can('view.crm') ? orders.filter(o => o.status === 'pending').length : 0, path: '/vendas?status=pending', icon: CreditCard },
  ].filter(p => p.count > 0);
  const shortcuts: Shortcut[] = [
    { path: '/dashboard', permission: 'view.dashboard', label: 'Visão geral', detail: 'Indicadores executivos', icon: LayoutDashboard },
    { path: '/crm', permission: 'view.crm', label: 'CRM e funis', detail: 'Oportunidades em aberto', icon: Workflow, count: opportunities.filter(o => !['won', 'lost'].includes(o.stage)).length },
    { path: '/clientes', permission: 'view.clients', label: 'Clientes', detail: 'Relacionamento e histórico', icon: Users, count: clients.length },
    { path: '/vendas', permission: 'view.sales', label: 'Vendas', detail: 'Pedidos aguardando pagamento', icon: CreditCard, count: orders.filter(o => o.status === 'pending').length },
    { path: '/consultas', permission: 'view.consultations', label: 'Consultas', detail: 'Entregas em processamento', icon: Crosshair, count: consultations.filter(o => o.consultationStatus === 'processing').length },
    { path: '/atendimento', permission: 'view.inbox', label: 'Atendimento', detail: 'Conversas em andamento', icon: MessageSquare, count: conversations.filter(c => c.status !== 'resolved').length },
    { path: '/whatsapp', permission: 'view.whatsapp', label: 'WhatsApp', detail: 'Instâncias disponíveis', icon: Activity, count: instances.length },
    { path: '/marketing', permission: 'view.marketing', label: 'Marketing', detail: 'Campanhas ativas da amostra', icon: Megaphone, count: state.campaigns.filter(c => c.status === 'active').length },
    { path: '/tracking', permission: 'view.tracking', label: 'Tracking', detail: 'Eventos recebidos na amostra', icon: TrendingUp, count: state.trackingEvents.filter(e => e.status === 'received' && (user.role !== 'manager' || (e.clientId && clientIds.has(e.clientId)))).length },
    { path: '/tarefas', permission: 'view.tasks', label: 'Tarefas', detail: 'Próximos passos em aberto', icon: CheckCircle2, count: openTasks.length },
    { path: '/evolucao', permission: 'view.evolution', label: 'Metas e evolução', detail: 'Metas operacionais visíveis', icon: Target, count: goals.length },
    { path: '/suporte', permission: 'view.support', label: 'Suporte interno', detail: 'Chamados em andamento', icon: Headphones, count: tickets.filter(t => t.status !== 'resolved').length },
    { path: '/relatorios', permission: 'view.reports', label: 'Relatórios', detail: 'Análises da demonstração', icon: BookOpen },
    { path: '/administracao', permission: 'view.admin', label: 'Administração', detail: 'Pessoas com acesso ativo', icon: Settings2, count: state.users.filter(u => u.active).length },
  ].filter(s => can(s.permission));
  const agenda = [...openTasks].sort((a, b) => a.dueDate.localeCompare(b.dueDate) || Number(b.priority === 'high') - Number(a.priority === 'high')).slice(0, 4);
  const mainAction = priorities[0];
  const orbProgress = can('view.evolution') ? leadership ? teamProgress : personal.progress : 0;
  const formattedNow = new Intl.DateTimeFormat('pt-BR', { timeZone: 'America/Sao_Paulo', weekday: 'long', day: 'numeric', month: 'long' }).format(now);

  return <div className="home-page">
    <section className="home-hero" aria-label="Boas-vindas">
      <div className="home-hero-copy">
        <div className="home-hero-meta"><span className="home-date"><GreetingIcon size={15} />{formattedNow}</span><span className="home-demo-label">Operação demonstrativa</span></div>
        <h1><span key={salutation} className="home-greeting">{salutation},</span><br />{firstName}<span className="home-greeting-dot">.</span></h1>
        <p className="home-hero-description">Tudo conectado ao seu próximo passo.<br /><span>{priorities.length ? `${priorities.length} prioridades para acompanhar hoje.` : 'Nenhuma pendência urgente por aqui. Escolha seu foco.'}</span></p>
        <div className="home-hero-actions">{mainAction ? <button className="home-primary" onClick={() => go(mainAction.path)}>{mainAction.key === 'late' ? 'Organizar prioridades' : mainAction.key === 'failed' ? 'Revisar primeira consulta' : mainAction.key === 'offline' ? 'Verificar conexões' : mainAction.key === 'conversations' ? 'Retomar atendimento' : mainAction.key === 'payment' ? 'Acompanhar oportunidade' : mainAction.key === 'orders' ? 'Acompanhar pagamentos' : 'Ver tarefas de hoje'}<ArrowRight size={17} /></button> : focus && <button className="home-primary" onClick={() => go(focus.path)}>Começar pelo meu foco<ArrowRight size={17} /></button>}{can('view.evolution') && <button className="home-secondary" onClick={() => go('/evolucao')}>Ver minha evolução<ArrowUpRight size={16} /></button>}</div>
        <div className="home-focus"><Target size={16} /><div><label htmlFor="home-focus">Meu foco agora</label><select id="home-focus" value={focus?.id || ''} onChange={e => setFocus(e.target.value)}>{focusOptions.length ? focusOptions.map(f => <option key={f.id} value={f.id}>{f.label}</option>) : <option value="">Acompanhar meu perfil</option>}</select><p>{focus?.description || 'Seus módulos aparecem conforme as permissões do perfil.'}</p></div>{focus && <button className="home-focus-go" onClick={() => go(focus.path)} aria-label={`Abrir foco: ${focus.label}`}><ArrowUpRight size={19} /></button>}</div>
      </div>
      <div className="home-hero-progress">
        <div className="home-progress-heading"><span>{lead ? 'VISÃO DE LIDERANÇA' : 'PROGRESSO COM PROPÓSITO'}</span><ShieldCheck size={17} /></div>
        <div className="home-orbit-wrap"><div className="home-orbit" aria-hidden="true"><span className="home-orbit-dot" /><svg viewBox="0 0 180 180"><circle className="home-orbit-track" cx="90" cy="90" r="75" /><circle className="home-orbit-fill" cx="90" cy="90" r="75" pathLength="100" strokeDasharray={`${orbProgress} 100`} /></svg></div><div className="home-orbit-content">{lead && personal.confirmed === 0 ? <><Users size={28} /><strong>Equipe</strong><span>Liderança que dá direção</span></> : can('view.evolution') ? <><span>NÍVEL</span><strong>{personal.level.toString().padStart(2, '0')}</strong><span>{personal.confirmed.toLocaleString('pt-BR')} créditos confirmados</span></> : <><ShieldCheck size={28} /><strong>Seu espaço</strong><span>{roleNames[user.role]}</span></>}</div></div>
        <div className="home-hero-bars">
          {can('view.evolution') && !leadership && <div><div className="home-bar-label"><span>Próxima faixa</span><b>{personal.next === null ? 'Faixa máxima' : `${personal.confirmed} / ${personal.next}`}</b></div><Bar value={personal.progress} label="Progresso individual para a próxima faixa" /></div>}
          {can('view.tasks') && myTasks.length > 0 && <div><div className="home-bar-label"><span>Minhas tarefas concluídas</span><b>{myTasks.filter(t => t.status === 'done').length} / {myTasks.length}</b></div><Bar value={taskCompletion} label="Tarefas individuais concluídas" /></div>}
          {teamGoal && <div><div className="home-bar-label"><span>{teamGoal.scope === 'team' ? 'Meta da equipe' : 'Meta coletiva'}</span><b>{Math.round(teamProgress)}%</b></div><Bar value={teamProgress} label={teamGoal.label} /></div>}
        </div>
        <p className="home-progress-note">{can('view.evolution') ? `${personal.pending} créditos pendentes de validação. ${leadership ? 'Seu papel é acompanhar o progresso da equipe.' : 'Só resultados confirmados compõem sua evolução.'}` : 'Acesso e informações seguem as permissões do seu perfil.'}</p>
      </div>
    </section>

    {can('view.dashboard') && <section className="home-executive" aria-label="Resumo executivo"><ExecutiveSummary /></section>}

    <section className="home-demo" aria-labelledby="home-demo-title">
      <div className="home-section-heading home-demo-heading"><div><span className="home-kicker">CONTINUIDADE DA OPERAÇÃO</span><h2 id="home-demo-title">Seu próximo passo, em contexto</h2><p>Dados demonstrativos · {roleNames[user.role]} · prazos comparados com a data atual em São Paulo.</p></div><span className="home-demo-label">Demo local</span></div>
      <div className="home-work-grid">
        <section className="home-panel home-priorities" aria-labelledby="home-priority-title"><div className="home-panel-heading"><div><h3 id="home-priority-title">Onde concentrar atenção</h3><p>Prioridades com uma ação clara.</p></div><span className="home-panel-count">{priorities.length}</span></div><div className="home-priority-list">{priorities.length ? priorities.slice(0, 5).map(p => <button key={p.key} className={`home-priority-row ${p.urgent ? 'home-priority-urgent' : ''}`} onClick={() => go(p.path)}><span className="home-row-icon"><p.icon size={18} /></span><span className="home-row-copy"><strong>{p.label}</strong><small>{p.detail}</small></span><span className="home-row-count">{p.count}</span><ChevronRight size={16} /></button>) : <div className="home-empty"><CheckCircle2 size={26} /><strong>Tudo encaminhado por aqui</strong><p>Nenhuma prioridade aberta para este perfil. Seus atalhos continuam disponíveis abaixo.</p></div>}</div></section>
        {can('view.tasks') ? <section className="home-panel home-agenda" aria-labelledby="home-agenda-title"><div className="home-panel-heading"><div><h3 id="home-agenda-title">Agenda de próximos passos</h3><p>{lead ? 'Pendências no seu escopo de gestão.' : 'Suas tarefas em ordem de prazo.'}</p></div><button className="home-text-link" onClick={() => go('/tarefas')}>Ver todas<ArrowUpRight size={14} /></button></div>{agenda.length ? <div className="home-agenda-list">{agenda.map(t => <div className="home-task-row" key={t.id}>{canCompleteTask(state, user, t.ownerId) ? <button className="home-task-check" aria-label={`Concluir tarefa: ${t.title}`} onClick={() => run('task.update', { id: t.id, status: 'done' }, 'Tarefa concluída. Créditos dependem de resultados verificados.')}><Check size={14} /></button> : <span className="home-task-marker"><Clock3 size={15} /></span>}<div><button className="home-task-title" onClick={() => go('/tarefas')}>{t.title}</button><small><span className={t.dueDate < clock.day ? 'home-overdue' : ''}>{t.dueDate < clock.day ? 'Atrasada · ' : t.dueDate === clock.day ? 'Hoje · ' : ''}{date(t.dueDate)}</span> · {state.users.find(u => u.id === t.ownerId)?.name.split(' ')[0] || 'Equipe'}{t.priority === 'high' && ' · Alta prioridade'}</small></div></div>)}</div> : <div className="home-empty"><CalendarDays size={26} /><strong>Agenda livre no seu escopo</strong><p>Organize a próxima atividade na área de tarefas.</p></div>}<p className="home-panel-footnote">Concluir uma tarefa organiza a agenda. Pontos exigem resultados e validação.</p></section> : <section className="home-panel home-focus-panel"><div className="home-panel-heading"><div><h3>Espaço para uma decisão melhor</h3><p>Concentre a navegação no que você precisa agora.</p></div><Target size={20} /></div><h4>{focus?.label || 'Seus acessos, reunidos'}</h4><p>{focus?.description || 'Explore os módulos autorizados logo abaixo.'}</p>{focus && <button className="home-secondary" onClick={() => go(focus.path)}>Abrir meu foco<ArrowRight size={16} /></button>}<small>O foco fica salvo neste navegador. Não altera metas nem gera créditos.</small></section>}
      </div>
      {can('view.evolution') && <section className="home-goals" aria-labelledby="home-goals-title"><div className="home-section-heading"><div><span className="home-kicker">RESULTADOS VERIFICÁVEIS</span><h2 id="home-goals-title">Evoluir com consistência</h2><p>Metas operacionais da demonstração. Créditos pendentes ficam fora do progresso individual.</p></div><button className="home-text-link" onClick={() => go('/evolucao')}>Metas e critérios<ArrowUpRight size={15} /></button></div><div className="home-goal-grid">{goalCards.length ? goalCards.map(g => { const percentage = g.target > 0 ? Math.min(100, g.current / g.target * 100) : 0; const format = (value: number) => `${value.toLocaleString('pt-BR', { maximumFractionDigits: 1 })}${g.unit === 'percent' ? '%' : ''}`; return <button key={g.id} className="home-goal-card" onClick={() => go('/evolucao')}><div className="home-goal-top"><span>{g.scope === 'organization' ? 'COLETIVA' : g.scope === 'team' ? 'EQUIPE' : 'INDIVIDUAL'}</span><Target size={17} /></div><h3>{g.label}</h3><div className="home-goal-value"><strong>{format(g.current)}</strong><span>de {format(g.target)}</span><b>{Math.round(percentage)}%</b></div><Bar value={percentage} label={g.label} /><p>{g.description}</p><span className="home-goal-bottom">{percentage >= 100 ? 'Objetivo atingido · consulte os critérios' : 'Continuar com qualidade'}<ArrowRight size={14} /></span></button>; }) : <div className="home-panel home-empty"><Target size={26} /><strong>Seu progresso começa com um resultado</strong><p>As metas operacionais aparecerão quando estiverem vinculadas ao seu escopo.</p></div>}</div></section>}
      <section className="home-shortcuts" aria-labelledby="home-shortcuts-title"><div className="home-section-heading"><div><span className="home-kicker">TODO O WORKSPACE</span><h2 id="home-shortcuts-title">Atalhos para seguir em frente</h2><p>Somente os módulos autorizados. As contagens abaixo usam os registros demonstrativos.</p></div><span className="home-module-total">{shortcuts.length} módulos</span></div><div className="home-shortcut-grid">{shortcuts.map(s => <button className="home-shortcut" key={s.path} onClick={() => go(s.path)}><span className="home-shortcut-icon"><s.icon size={19} /></span><div><strong>{s.label}</strong><small>{s.detail}</small></div>{s.count !== undefined ? <span className="home-shortcut-count">{s.count.toLocaleString('pt-BR')}</span> : <ArrowUpRight size={16} className="home-shortcut-arrow" />}</button>)}</div></section>
    </section>
  </div>;
}
