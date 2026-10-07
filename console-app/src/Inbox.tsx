import { useEffect, useRef, useState } from 'react';
import { ArrowLeft, Check, CheckCheck, CheckCircle2, ChevronDown, Clock3, Info, MessageSquare, Paperclip, Plus, Search, Send, ShieldCheck, StickyNote, UserRound, X } from 'lucide-react';
import { useDemo, money, date, go, statusNames, stageNames } from './context';
import { Avatar, Badge, Empty, FormDialog, InfoNote, LinkButton, PageHeading } from './components';

const DEMO_NOW = new Date('2026-10-06T15:00:00Z').getTime();
const templates = [
  { id: 'retomar-atendimento', label: 'Retomar atendimento', text: 'Olá! Podemos retomar seu atendimento no Verifica Placa? Responda a esta mensagem para continuar.' },
  { id: 'relatorio-disponivel', label: 'Atualização do pedido', text: 'Olá! Temos uma atualização do seu pedido no Verifica Placa. Responda para falar com nossa equipe.' },
];
const time = (value: string) => new Date(value).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
const messageDate = (value: string) => new Date(value).toLocaleString('pt-BR', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
type Draft = { text: string; mode: 'reply' | 'note'; templateId: string };
type DialogType = 'assign' | 'task' | 'ticket' | 'opportunity' | 'order' | null;

export default function Inbox({ conversationId }: { conversationId?: string }) {
  const { state, run, can, user, notify } = useDemo();
  const [search, setSearch] = useState('');
  const [tab, setTab] = useState(['owner', 'admin', 'manager'].includes(user.role) ? 'all' : 'mine');
  const [queue, setQueue] = useState('all');
  const [status, setStatus] = useState('all');
  const [drafts, setDrafts] = useState<Record<string, Draft>>({});
  const [dialog, setDialog] = useState<DialogType>(null);
  const [showContext, setShowContext] = useState(false);
  const [mobileList, setMobileList] = useState(!conversationId);
  const bottom = useRef<HTMLDivElement>(null);
  const attachment = useRef<HTMLInputElement>(null);
  const isAdmin = ['owner', 'admin'].includes(user.role);
  const teamUsers = state.users.filter(u => u.active && (isAdmin || u.teamId === user.teamId));
  const assignUsers = teamUsers.filter(u => ['owner', 'admin', 'manager', 'sales', 'support'].includes(u.role));
  const accessible = state.conversations.filter(c => {
    if (isAdmin) return true;
    const instance = state.instances.find(i => i.id === c.instanceId);
    const owner = state.users.find(u => u.id === c.ownerId);
    return c.ownerId === user.id || owner?.teamId === user.teamId || (!owner && instance?.teamId === user.teamId);
  });
  const findClient = (id: string) => state.clients.find(c => c.id === id);
  const filtered = accessible.filter(c => {
    const client = findClient(c.clientId);
    const matchesTab = tab === 'resolved' ? c.status === 'resolved' : c.status !== 'resolved' && (tab !== 'mine' || c.ownerId === user.id);
    const needle = search.toLocaleLowerCase('pt-BR').trim();
    const haystack = [client?.name, client?.phone, client?.email, ...c.messages.map(m => m.text)].join(' ').toLocaleLowerCase('pt-BR');
    return matchesTab && (queue === 'all' || c.queue === queue) && (status === 'all' || c.status === status) && (!needle || haystack.includes(needle));
  }).sort((a, b) => Number(b.unread > 0) - Number(a.unread > 0) || new Date(b.messages.at(-1)?.at || 0).getTime() - new Date(a.messages.at(-1)?.at || 0).getTime());
  const selected = conversationId ? accessible.find(c => c.id === conversationId) : undefined;
  const client = selected ? findClient(selected.clientId) : undefined;
  const instance = state.instances.find(i => i.id === selected?.instanceId);
  const owner = state.users.find(u => u.id === selected?.ownerId);
  const draft: Draft = selected && drafts[selected.id] || { text: '', mode: 'reply', templateId: '' };
  const inbound = selected?.messages.filter(m => m.kind === 'incoming').at(-1);
  const expiresAt = inbound ? new Date(inbound.at).getTime() + 24 * 60 * 60 * 1000 : 0;
  const windowClosed = instance?.type === 'official' && (!expiresAt || DEMO_NOW >= expiresAt);
  const offline = instance?.status === 'disconnected';
  const isNote = draft.mode === 'note';
  const resolved = selected?.status === 'resolved';
  const allowed = can('inbox.write');
  const selectedTemplate = templates.find(t => t.id === draft.templateId);
  const orders = state.orders.filter(o => o.clientId === client?.id).sort((a, b) => new Date(b.paidAt || 0).getTime() - new Date(a.paidAt || 0).getTime()).slice(0, 3);
  const tasks = state.tasks.filter(t => t.clientId === client?.id);
  const tickets = state.tickets.filter(t => t.clientId === client?.id);
  const opportunities = state.opportunities.filter(o => o.clientId === client?.id);
  const queues = Array.from(new Set(accessible.map(c => c.queue)));
  const updateDraft = (changes: Partial<Draft>) => { if (selected) setDrafts(prev => ({ ...prev, [selected.id]: { ...draft, ...changes } })); };
  useEffect(() => { setMobileList(!conversationId); setDialog(null); setShowContext(false); }, [conversationId]);
  useEffect(() => {
    const messages = bottom.current?.parentElement;
    messages?.scrollTo({ top: messages.scrollHeight });
  }, [selected?.id, selected?.messages.length]);
  const openConversation = (id: string) => { setMobileList(false); go(`/atendimento/${id}`); };
  const send = () => {
    if (!selected || !allowed || resolved) return;
    const text = !isNote && windowClosed ? selectedTemplate?.text : draft.text.trim();
    if (!text) { notify(windowClosed && !isNote ? 'Escolha um modelo demonstrativo para continuar.' : 'Escreva uma mensagem antes de enviar.'); return; }
    const result = run(isNote ? 'conversation.note' : 'conversation.send', { id: selected.id, text, ...(!isNote && windowClosed ? { templateId: draft.templateId } : {}) }, isNote ? 'Nota interna adicionada.' : offline ? 'Mensagem pendente na demonstração: conexão indisponível.' : 'Mensagem enviada na demonstração.');
    if (result) updateDraft({ text: '', templateId: '' });
  };
  const counters = {
    mine: accessible.filter(c => c.status !== 'resolved' && c.ownerId === user.id).length,
    team: accessible.filter(c => c.status !== 'resolved' && (state.users.find(u => u.id === c.ownerId)?.teamId === user.teamId || (!c.ownerId && state.instances.find(i => i.id === c.instanceId)?.teamId === user.teamId))).length,
    all: accessible.filter(c => c.status !== 'resolved').length,
    resolved: accessible.filter(c => c.status === 'resolved').length,
  };
  const tabs = [{ id: 'mine', label: 'Minhas' }, { id: 'team', label: 'Equipe' }, { id: 'all', label: 'Todas' }, { id: 'resolved', label: 'Resolvidas' }];
  const tabFiltered = tab === 'team' ? filtered.filter(c => state.users.find(u => u.id === c.ownerId)?.teamId === user.teamId || (!c.ownerId && state.instances.find(i => i.id === c.instanceId)?.teamId === user.teamId)) : filtered;

  return <>
    <PageHeading eyebrow="Relacionamento" title="Atendimento" description="Uma conversa contínua, com todo o contexto do cliente." >
      <span className="badge"><ShieldCheck size={14} /> Ambiente de demonstração</span>
    </PageHeading>
    <div className={`inbox-layout ${selected && !mobileList ? 'has-conversation' : 'show-list'} ${showContext ? 'show-context' : ''}`}>
      <aside className="inbox-list" aria-label="Lista de conversas">
        <div className="inbox-list-heading"><div><h2>Conversas</h2><small>{accessible.filter(c => c.status !== 'resolved').length} em andamento</small></div><MessageSquare size={20} /></div>
        <label className="search-field"><Search size={16} /><input aria-label="Buscar conversas" placeholder="Nome, telefone ou mensagem" value={search} onChange={e => setSearch(e.target.value)} /></label>
        <div className="tabs inbox-tabs" role="tablist" aria-label="Filtro de responsável">{tabs.map(t => <button key={t.id} role="tab" aria-selected={tab === t.id} className={tab === t.id ? 'selected' : ''} onClick={() => setTab(t.id)}>{t.label}<span>{counters[t.id as keyof typeof counters]}</span></button>)}</div>
        <div className="inbox-filters"><select className="select" aria-label="Filtrar por fila" value={queue} onChange={e => setQueue(e.target.value)}><option value="all">Todas as filas</option>{queues.map(q => <option key={q} value={q}>{q}</option>)}</select><select className="select" aria-label="Filtrar por status" value={status} onChange={e => setStatus(e.target.value)}><option value="all">Todos os status</option>{['open', 'waiting', 'internal', 'resolved'].map(s => <option key={s} value={s}>{statusNames[s]}</option>)}</select></div>
        <div className="conversation-rows">{tabFiltered.length ? tabFiltered.map(c => {
          const person = findClient(c.clientId);
          const last = c.messages.at(-1);
          const assigned = state.users.find(u => u.id === c.ownerId);
          return <button key={c.id} className={`conversation-row ${selected?.id === c.id ? 'active' : ''}`} onClick={() => openConversation(c.id)} aria-current={selected?.id === c.id ? 'page' : undefined}>
            <Avatar name={person?.name || 'Cliente'} /><div className="conversation-content"><div className="conversation-meta"><strong>{person?.name || 'Cliente'}</strong><time>{last ? time(last.at) : '—'}</time></div><p className="conversation-preview">{last?.kind === 'note' && 'Nota interna · '}{last?.text || 'Conversa iniciada'}</p><div className="conversation-meta"><span>{assigned?.name.split(' ')[0] || 'Sem responsável'} · {c.queue}</span><Badge status={c.status} /></div></div>{c.unread > 0 && <span className="unread-count" aria-label={`${c.unread} mensagens não lidas`}>{c.unread}</span>}
          </button>;
        }) : <Empty title="Nenhuma conversa aqui" description="Experimente outra aba ou ajuste os filtros." />}</div>
      </aside>
      <section className="inbox-chat" aria-label="Conversa selecionada">
        {!selected || !client ? <Empty title={conversationId ? 'Conversa indisponível' : 'Seu próximo atendimento começa aqui'} description={conversationId ? 'Esta conversa não está disponível para o seu perfil.' : 'Selecione uma conversa para responder e acompanhar o cliente.'} /> : <>
          <header className="chat-header">
            <button className="icon-button inbox-mobile-back" aria-label="Voltar à lista" onClick={() => { setMobileList(true); go('/atendimento'); }}><ArrowLeft size={18} /></button>
            <Avatar name={client.name} /><div className="chat-client"><button className="text-button" onClick={() => go(`/clientes/${client.id}`)}><strong>{client.name}</strong></button><small>{instance?.name || 'Canal'} · {selected.queue}</small></div>
            <div className="chat-header-actions"><Badge status={selected.status} />{allowed && !selected.ownerId && <button className="button secondary small" onClick={() => run('conversation.assign', { id: selected.id, ownerId: user.id }, 'Você assumiu este atendimento.')}>Assumir</button>}{allowed && <button className="button ghost small" onClick={() => setDialog('assign')}><UserRound size={14} />{owner?.name.split(' ')[0] || 'Atribuir'}<ChevronDown size={13} /></button>}{allowed && <button className="icon-button" title={resolved ? 'Reabrir conversa' : 'Resolver conversa'} aria-label={resolved ? 'Reabrir conversa' : 'Resolver conversa'} onClick={() => run(resolved ? 'conversation.reopen' : 'conversation.resolve', { id: selected.id }, resolved ? 'Conversa reaberta.' : 'Conversa resolvida.')}><CheckCircle2 size={19} /></button>}<button className="icon-button inbox-context-toggle" aria-label={showContext ? 'Ocultar contexto do cliente' : 'Mostrar contexto do cliente'} aria-expanded={showContext} onClick={() => setShowContext(v => !v)}><Info size={19} /></button></div>
          </header>
          <div className="chat-channel-info"><span><i className={`connection-dot ${offline ? 'offline' : ''}`} />{offline ? 'Canal desconectado' : 'Canal conectado'}</span><span>{instance?.type === 'official' ? 'WhatsApp oficial' : 'Conexão alternativa'} · simulação</span></div>
          <div className="message-list" aria-live="polite" aria-relevant="additions">
            <div className="chat-day-label">Histórico do atendimento</div>
            {selected.messages.map(m => <article key={m.id} className={`message-bubble ${m.kind}`}>
              {m.kind === 'note' && <div className="message-note-label"><StickyNote size={13} /> Nota interna · visível para a equipe</div>}
              <div className="message-author">{m.kind === 'incoming' ? client.name : m.author || 'Equipe'}</div><p>{m.text}</p>
              <div className="message-meta"><time dateTime={m.at}>{messageDate(m.at)}</time>{m.kind === 'outgoing' && <span aria-label={m.status === 'read' ? 'Lida' : m.status === 'pending' ? 'Pendente' : m.status === 'failed' ? 'Falhou' : 'Enviada'}>{m.status === 'read' ? <CheckCheck size={14} /> : m.status === 'pending' ? <><Clock3 size={12} /> Pendente</> : m.status === 'failed' ? 'Falhou' : <Check size={14} />}</span>}</div>
            </article>)}
            <div ref={bottom} />
          </div>
          <div className={`composer ${isNote ? 'internal-mode' : ''}`}>
            {!allowed ? <InfoNote>Seu perfil permite consultar este histórico. O envio exige permissão de atendimento.</InfoNote> : resolved ? <InfoNote>Conversa resolvida. Reabra no cabeçalho para enviar uma mensagem ou adicionar uma nota.</InfoNote> : <>
              <div className="composer-controls"><div className="tabs" aria-label="Tipo de mensagem"><button className={!isNote ? 'selected' : ''} onClick={() => updateDraft({ mode: 'reply' })}><MessageSquare size={13} /> Resposta</button><button className={isNote ? 'selected' : ''} onClick={() => updateDraft({ mode: 'note' })}><StickyNote size={13} /> Nota interna</button></div><small>{isNote ? 'Somente a equipe vê esta nota' : 'Resposta demonstrativa ao cliente'}</small></div>
              {!isNote && offline && <InfoNote>Mensagens ficam pendentes enquanto o canal está desconectado. {can('view.whatsapp') && <button className="text-button" onClick={() => go('/whatsapp')}>Ver instância</button>}</InfoNote>}
              {!isNote && instance?.type === 'official' && <div className="chat-window-note"><Clock3 size={13} />{windowClosed ? 'Janela de 24 horas encerrada. Use um modelo aprovado de demonstração.' : `Janela de atendimento aberta até ${messageDate(new Date(expiresAt).toISOString())}.`}<small>Relógio demo: 6 out. 2026, 12h BRT</small></div>}
              {!isNote && windowClosed ? <div className="template-picker"><label htmlFor="inbox-template">Modelo aprovado · demonstração</label><select id="inbox-template" className="select" value={draft.templateId} onChange={e => updateDraft({ templateId: e.target.value })}><option value="">Selecione um modelo</option>{templates.map(t => <option value={t.id} key={t.id}>{t.label}</option>)}</select>{selectedTemplate && <p className="template-preview">{selectedTemplate.text}</p>}</div> : <textarea aria-label={isNote ? 'Escrever nota interna' : 'Escrever resposta'} rows={3} placeholder={isNote ? 'Registre um contexto para a equipe…' : 'Escreva sua resposta…'} value={draft.text} onChange={e => updateDraft({ text: e.target.value })} onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) { e.preventDefault(); send(); } }} />}
              <div className="composer-footer"><div><input ref={attachment} type="file" hidden onChange={e => { const file = e.target.files?.[0]; if (file) { updateDraft({ mode: 'note', text: `${draft.text}${draft.text ? '\n' : ''}Arquivo de referência: ${file.name} (anexo demonstrativo; conteúdo não enviado).` }); notify('Nome do arquivo adicionado ao rascunho de nota interna.'); } e.target.value = ''; }} /><button className="icon-button" aria-label="Adicionar referência de arquivo local à nota" title="Adicionar referência de arquivo à nota" onClick={() => attachment.current?.click()}><Paperclip size={18} /></button><small>Enter envia · Shift + Enter quebra linha</small></div><button className="button primary small" onClick={send} disabled={(!isNote && windowClosed) ? !selectedTemplate : !draft.text.trim()}><Send size={15} />{isNote ? 'Salvar nota' : 'Enviar'}</button></div>
            </>}
          </div>
        </>}
      </section>
      <aside className="inbox-context" aria-label="Contexto do cliente">
        {client && selected ? <>
          <div className="inbox-context-heading"><h2>Contexto do cliente</h2><button className="icon-button inbox-context-close" aria-label="Fechar contexto" onClick={() => setShowContext(false)}><X size={18} /></button></div>
          <section className="context-section"><Avatar name={client.name} size="large" /><h3>{client.name}</h3><span className="badge">{client.tag || 'Cliente'}</span><dl className="context-details"><dt>Telefone</dt><dd>{client.phone || 'Não informado'}</dd><dt>E-mail</dt><dd>{client.email || 'Não informado'}</dd><dt>Empresa</dt><dd>{client.company || 'Pessoa física'}</dd><dt>Origem</dt><dd>{client.origin || 'Não informada'}</dd><dt>Responsável</dt><dd>{owner?.name || 'Sem responsável'}</dd></dl><LinkButton onClick={() => go(`/clientes/${client.id}`)}>Abrir perfil completo</LinkButton></section>
          <section className="context-section"><div className="context-section-heading"><h3>Pedidos recentes</h3><span>{orders.length}</span>{can('orders.write') && <button className="icon-button" aria-label="Criar pedido do atendimento" onClick={() => setDialog('order')}><Plus size={16} /></button>}</div>{orders.length ? orders.map(o => <button className="context-link" key={o.id} onClick={() => go(`/vendas/${o.id}`)}><div><strong>{o.product === 'complete' ? 'Consulta completa' : 'Consulta base'}</strong><small>{o.id}</small></div><div><strong>{money(o.totalCents)}</strong><Badge status={o.status} /></div></button>) : <p className="muted">Nenhum pedido registrado.</p>}</section>
          <section className="context-section"><div className="context-section-heading"><h3>Próximas atividades</h3>{can('tasks.write') && <button className="icon-button" aria-label="Criar tarefa" onClick={() => setDialog('task')}><Plus size={16} /></button>}</div>{tasks.length ? tasks.slice(-3).map(t => <div className="context-activity" key={t.id}><CheckCircle2 size={15} /><div><strong>{t.title}</strong><small>Prazo: {date(t.dueDate)}</small><Badge status={t.status} /></div></div>) : <p className="muted">Nenhuma tarefa para este cliente.</p>}</section>
          <section className="context-section"><div className="context-section-heading"><h3>Oportunidades</h3>{can('crm.write') && <button className="icon-button" aria-label="Criar oportunidade" onClick={() => setDialog('opportunity')}><Plus size={16} /></button>}</div>{opportunities.length ? opportunities.slice(-2).map(o => <button key={o.id} className="context-link" onClick={() => go('/crm')}><strong>{o.title}</strong><Badge status={o.stage}>{stageNames[o.stage]}</Badge></button>) : <p className="muted">Nenhuma oportunidade vinculada.</p>}</section>
          <section className="context-section"><div className="context-section-heading"><h3>Chamados</h3>{can('support.write') && <button className="icon-button" aria-label="Abrir chamado" onClick={() => setDialog('ticket')}><Plus size={16} /></button>}</div>{tickets.length ? tickets.slice(-2).map(t => <button className="context-link" key={t.id} onClick={() => go('/suporte')}><strong>{t.title}</strong><Badge status={t.status} /></button>) : <p className="muted">Nenhum chamado aberto.</p>}</section>
        </> : <Empty title="Tudo no mesmo lugar" description="Dados, pedidos e atividades aparecem ao selecionar uma conversa." />}
      </aside>
    </div>
    {dialog && selected && client && <FormDialog title={dialog === 'assign' ? 'Transferir atendimento' : dialog === 'task' ? 'Criar tarefa' : dialog === 'ticket' ? 'Abrir chamado' : dialog === 'order' ? 'Criar pedido do atendimento' : 'Criar oportunidade'} close={() => setDialog(null)} submitLabel={dialog === 'assign' ? 'Transferir' : 'Criar'} fields={dialog === 'assign' ? [{ name: 'ownerId', label: 'Responsável', value: selected.ownerId || '', options: [...(isAdmin ? [{ value: '', label: 'Sem responsável' }] : []), ...assignUsers.map(u => ({ value: u.id, label: `${u.name}${u.id === user.id ? ' (você)' : ''}` }))] }] : dialog === 'task' ? [{ name: 'title', label: 'Título da tarefa', required: true }, { name: 'ownerId', label: 'Responsável', value: user.id, options: teamUsers.map(u => ({ value: u.id, label: u.name })) }, { name: 'dueDate', label: 'Prazo', type: 'date', required: true, value: '2026-10-07' }] : dialog === 'ticket' ? [{ name: 'title', label: 'Assunto', required: true }, { name: 'description', label: 'Descrição', type: 'textarea', required: true }, { name: 'orderId', label: 'Pedido relacionado', options: [{ value: '', label: 'Sem pedido relacionado' }, ...state.orders.filter(o => o.clientId === client.id).map(o => ({ value: o.id, label: `${o.id} · ${o.product === 'complete' ? 'Consulta completa' : 'Consulta base'}` }))] }] : dialog === 'order' ? [{name:'plate',label:'Placa demonstrativa',value:'DEM1A23',required:true},{name:'product',label:'Produto',options:[{value:'base',label:'Consulta Base · R$ 14,99'},{value:'complete',label:'Pacote Completo · R$ 79,90'}]}] : [{ name: 'title', label: 'Nome da oportunidade', required: true, value: `Nova consulta · ${client.name}` }, { name: 'value', label: 'Valor previsto (R$)', type: 'number', min: 0, required: true }, { name: 'ownerId', label: 'Responsável', value: user.id, options: teamUsers.map(u => ({ value: u.id, label: u.name })) }]} onSubmit={values => {
      const permission = dialog === 'assign' ? 'inbox.write' : dialog === 'task' ? 'tasks.write' : dialog === 'ticket' ? 'support.write' : dialog === 'order' ? 'orders.write' : 'crm.write';
      if (!can(permission)) { notify('Seu perfil não permite esta ação.'); return false; }
      if (!['assign','order'].includes(dialog) && !String(values.title || '').trim()) { notify('Informe um título válido.'); return false; }
      if (dialog === 'order') { const next=run('order.create',{...values,clientId:client.id,assistedBy:user.id,campaignId:null},'Pedido criado a partir do atendimento.'); if(next) go(`/vendas/${next.orders.at(-1)!.id}`); return !!next; }
      const result = dialog === 'assign' ? run('conversation.assign', { id: selected.id, ownerId: values.ownerId || null }, 'Responsável atualizado.') : dialog === 'task' ? run('task.create', { ...values, title: values.title.trim(), clientId: client.id, priority: 'normal' }, 'Tarefa criada a partir do atendimento.') : dialog === 'ticket' ? run('ticket.create', { ...values, title: values.title.trim(), clientId: client.id, orderId: values.orderId || null, priority: 'normal' }, 'Chamado aberto e vinculado ao cliente.') : run('crm.create', { title: values.title.trim(), clientId: client.id, valueCents: Math.round(Number(values.value) * 100), ownerId: values.ownerId, campaignId: null }, 'Oportunidade criada no CRM.');
      return !!result;
    }} />}
  </>;
}
