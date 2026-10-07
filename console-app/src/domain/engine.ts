/** Local, deterministic demo domain. All amounts are integer BRL cents. */
export type Role = 'owner' | 'admin' | 'manager' | 'sales' | 'support' | 'marketing' | 'finance';
export type Priority = 'high' | 'normal' | 'low';
export interface User { id: string; name: string; email: string; role: Role; teamId: string; title: string; initials: string; active: boolean; avatar?: string }
export interface Client { id: string; name: string; email: string; phone: string; company: string; tag: string; origin: string; ownerId: string; createdAt: string; newBuyer: boolean }
export interface Order { id: string; clientId: string; product: 'base' | 'complete'; totalCents: number; addonCents: number; status: 'paid' | 'pending' | 'refunded'; paidAt: string | null; campaignId: string | null; assistedBy: string | null; consultationStatus: 'delivered' | 'processing' | 'failed' | 'cancelled' | 'waiting'; plate: string; bureauCents: number; feeCents: number }
export interface Opportunity { id: string; clientId: string; title: string; valueCents: number; stage: 'new' | 'qualified' | 'proposal' | 'payment' | 'won' | 'lost'; ownerId: string; campaignId: string | null; updatedAt: string }
export interface Message { id: string; text: string; kind: 'incoming' | 'outgoing' | 'note'; at: string; author: string; status: 'sent' | 'pending' | 'failed' | 'read' }
export interface Conversation { id: string; clientId: string; instanceId: string; ownerId: string | null; queue: string; status: 'open' | 'waiting' | 'internal' | 'resolved'; unread: number; priority: boolean; messages: Message[] }
export interface Task { id: string; title: string; clientId: string | null; ownerId: string; dueDate: string; priority: Priority; status: 'todo' | 'doing' | 'blocked' | 'done' }
export interface Ticket { id: string; title: string; description: string; clientId: string | null; orderId: string | null; ownerId: string; status: 'open' | 'doing' | 'resolved'; priority: Priority; createdAt: string }
export interface Instance { id: string; name: string; number: string; type: 'official' | 'alternative'; status: 'connected' | 'disconnected'; teamId: string; lastSync: string }
export interface Campaign { id: string; name: string; channel: string; adsCents: number; status: 'active' | 'paused'; budgetCents: number; checkouts: number; lostBudget: number; lostRank: number; impressionShare: number }
export interface Invite { id: string; email: string; name: string; role: Role; teamId: string; status: 'pending' | 'accepted' | 'expired' | 'revoked'; expiresAt: string }
export interface Credit { id: string; eventId: string; userId: string; teamId: string; label: string; points: number; status: 'pending' | 'confirmed' | 'reverted'; at: string; reason?: string }
export interface Goal { id: string; label: string; current: number; target: number; unit: 'currency' | 'percent' | 'count'; scope: 'organization' | 'team' | 'individual'; ownerId: string; description: string }
export interface TrackingEvent { id: string; clientId: string | null; event: string; source: string; at: string; status: 'received' | 'failed' }
export interface Audit { id: string; at: string; userId: string; action: string; detail: string }
export interface State { schemaVersion: number; currentUserId: string; rolePermissions?: Record<string, string[]>; users: User[]; clients: Client[]; orders: Order[]; opportunities: Opportunity[]; conversations: Conversation[]; tasks: Task[]; tickets: Ticket[]; instances: Instance[]; campaigns: Campaign[]; invites: Invite[]; credits: Credit[]; goals: Goal[]; trackingEvents: TrackingEvent[]; audit: Audit[]; settings: { theme: 'dark' | 'light'; compact: boolean } }
export interface Action { type: string; payload?: any }
export interface Metrics { revenueCents: number; refundsCents: number; netCents: number; adsCents: number; bureauCents: number; feeCents: number; taxCents: number; contributionCents: number; margin: number; paidOrders: number; newBuyers: number; ticketCents: number; upsellRate: number; checkoutRate: number; cacCents: number; cpaCents: number; roas: number; pendingOrders: number; failedConsultations: number }
export interface MetricFilters { from?: string; to?: string; campaignId?: string; teamId?: string }
export const DEMO_NOW = '2026-10-06T15:00:00.000Z';
export const APPROVED_TEMPLATES = ['retomar-atendimento', 'confirmacao-pagamento', 'relatorio-disponivel'] as const;
const BASE = 1499, COMPLETE = 7990, ADDON = 1990;
const roles: Role[] = ['owner', 'admin', 'manager', 'sales', 'support', 'marketing', 'finance'];
const stages: Opportunity['stage'][] = ['new', 'qualified', 'proposal', 'payment', 'won', 'lost'];
const views = ['dashboard', 'crm', 'clients', 'sales', 'consultations', 'inbox', 'whatsapp', 'marketing', 'tracking', 'tasks', 'evolution', 'support', 'reports', 'admin'];
const writePermissions = ['finance.refund', 'admin.users', 'admin.integrations', 'marketing.send', 'crm.write', 'orders.write', 'consultations.write', 'inbox.write', 'tasks.write', 'support.write', 'goals.write', 'profile.write'];
const grants: Record<Role, string[]> = {
  owner: [...views.map(v => `view.${v}`), ...writePermissions],
  admin: [...views.map(v => `view.${v}`), ...writePermissions],
  manager: [...views.filter(v => v !== 'admin').map(v => `view.${v}`), 'marketing.send', 'crm.write', 'orders.write', 'consultations.write', 'inbox.write', 'tasks.write', 'support.write', 'goals.write', 'profile.write'],
  sales: ['view.clients', 'view.crm', 'view.sales', 'view.inbox', 'view.tasks', 'view.evolution', 'view.support', 'crm.write', 'orders.write', 'inbox.write', 'tasks.write', 'support.write', 'profile.write'],
  support: ['view.clients', 'view.sales', 'view.consultations', 'view.inbox', 'view.tasks', 'view.evolution', 'view.support', 'consultations.write', 'inbox.write', 'tasks.write', 'support.write', 'profile.write'],
  marketing: ['view.dashboard', 'view.marketing', 'view.tracking', 'view.tasks', 'view.evolution', 'marketing.send', 'tasks.write', 'profile.write'],
  finance: ['view.dashboard', 'view.sales', 'view.reports', 'finance.refund', 'profile.write'],
};
export function can(role: string, permission: string, state?: State): boolean { return (state?.rolePermissions?.[role] ?? grants[role as Role] ?? []).includes(permission); }
const firstNames = ['Ana', 'Bruno', 'Carla', 'Diego', 'Elisa', 'Felipe', 'Gabriela', 'Henrique', 'Isabela', 'João', 'Larissa', 'Marcos', 'Natália', 'Otávio', 'Patrícia', 'Rafael', 'Sofia', 'Tiago', 'Vanessa', 'William'];
const surnames = ['Almeida', 'Barbosa', 'Costa', 'Dias', 'Esteves', 'Ferreira', 'Gomes', 'Henriques', 'Lima', 'Macedo', 'Nunes', 'Oliveira', 'Pereira', 'Ramos', 'Santos', 'Teixeira', 'Vieira', 'Moura', 'Rocha', 'Souza', 'Andrade', 'Batista', 'Carvalho', 'Duarte'];
const favorites = ['Mariana Costa', 'Lucas Almeida', 'Fernanda Santos', 'Pedro Oliveira', 'Camila Ferreira', 'Rafael Lima', 'Juliana Rocha', 'André Pereira', 'Beatriz Souza', 'Gabriel Martins', 'Renata Dias', 'Thiago Ribeiro'];
const initials = (name: string) => name.split(' ').filter(Boolean).map(n => n[0]).slice(0, 2).join('').toUpperCase();
const historicalAt = (i: number) => `2026-10-0${1 + i % 6}T${String(9 + i % 6).padStart(2, '0')}:00:00.000Z`;

export function createSeed(): State {
  const users: User[] = [
    ['USR-01', 'Daniel Macedo', 'owner', 'team-sales', 'Fundador'],
    ['USR-02', 'Renata Alves', 'admin', 'team-support', 'Administradora'],
    ['USR-03', 'Marina Ribeiro', 'manager', 'team-sales', 'Gerente comercial'],
    ['USR-04', 'Lucas Mendes', 'sales', 'team-sales', 'Consultor comercial'],
    ['USR-05', 'Camila Torres', 'sales', 'team-sales', 'Consultora comercial'],
    ['USR-06', 'Bruno Rocha', 'support', 'team-support', 'Especialista de suporte'],
    ['USR-07', 'Julia Martins', 'marketing', 'team-growth', 'Analista de marketing'],
    ['USR-08', 'Felipe Santos', 'finance', 'team-growth', 'Analista financeiro'],
  ].map(([id, name, role, teamId, title]) => ({ id, name, role: role as Role, teamId, title, initials: initials(name), active: true, email: `${name.toLowerCase().split(' ')[0]}@verificaplaca.demo` }));
  const clients: Client[] = Array.from({ length: 480 }, (_, i) => {
    const buyer = i < 180, media = buyer && i < 135;
    const name = favorites[i] ?? `${firstNames[i % firstNames.length]} ${surnames[Math.floor(i / firstNames.length) % surnames.length]}`;
    return { id: `CLI-${String(i + 1).padStart(4, '0')}`, name, email: `cliente${i + 1}@exemplo.demo`, phone: `+55119${String(50000000 + i).padStart(8, '0')}`, company: i % 5 === 0 ? ['Auto Prime', 'Costa Veículos', 'Ribeiro Motors'][i % 3] : '', tag: buyer ? i < 12 ? 'Recorrente' : 'Comprador' : 'Lead', origin: media || (!buyer && i % 3 !== 0) ? 'Mídia paga' : i % 2 ? 'Indicação' : 'Orgânico', ownerId: i % 2 ? 'USR-05' : 'USR-04', createdAt: buyer ? `2026-09-${String(1 + i % 28).padStart(2, '0')}T12:00:00.000Z` : historicalAt(i), newBuyer: buyer && (media ? i >= 12 && i < 102 : i < 165) };
  });
  const campaigns: Campaign[] = [
    ['CAM-01', 'Pesquisa • consulta de placa', 'Google Ads', 65000],
    ['CAM-02', 'Pesquisa • histórico veicular', 'Google Ads', 48000],
    ['CAM-03', 'Remarketing • compradores', 'Meta Ads', 32000],
    ['CAM-04', 'Social • primeira consulta', 'Meta Ads', 22000],
    ['CAM-05', 'Marca • Verifica Placa', 'Google Ads', 13000],
  ].map(([id, name, channel, adsCents], i) => ({ id: String(id), name: String(name), channel: String(channel), adsCents: Number(adsCents), budgetCents: Number(adsCents) + 10000, status: i === 4 ? 'paused' : 'active', checkouts: 240, lostBudget: [0.12, 0.18, 0.09, 0.22, 0.04][i], lostRank: [0.15, 0.2, 0.13, 0.18, 0.06][i], impressionShare: [0.73, 0.62, 0.78, 0.6, 0.9][i] }));
  const orders: Order[] = Array.from({ length: 240 }, (_, i) => {
    const media = i < 180, local = media ? i : i - 180;
    const clientIndex = media ? local % 135 : 135 + local % 45;
    const product = (media ? local < 160 : local < 56) ? 'base' : 'complete';
    const addonCents = (media ? local < 40 : local < 14) ? ADDON : 0;
    return { id: `ORD-${1001 + i}`, clientId: clients[clientIndex].id, product, totalCents: (product === 'base' ? BASE : COMPLETE) + addonCents, addonCents, status: i === 230 || i === 231 ? 'refunded' : 'paid', paidAt: historicalAt(i), campaignId: media ? campaigns[i % 5].id : null, assistedBy: clients[clientIndex].ownerId, consultationStatus: i < 220 ? 'delivered' : i === 230 || i === 231 ? 'cancelled' : i >= 232 && i < 238 ? 'failed' : 'processing', plate: `BRA${i % 10}${String.fromCharCode(65 + i % 26)}${String(i % 100).padStart(2, '0')}`, bureauCents: 375, feeCents: 75 };
  });
  // Refunded orders remain historical paid events; consultation progress is independently tracked.
  for (let i = 0; i < 4; i++) orders.push({ id: `ORD-${1241 + i}`, clientId: clients[180 + i].id, product: i === 3 ? 'complete' : 'base', totalCents: i === 3 ? COMPLETE : BASE, addonCents: 0, status: 'pending', paidAt: null, campaignId: campaigns[i].id, assistedBy: clients[180 + i].ownerId, consultationStatus: 'waiting', plate: `DEM${i}A26`, bureauCents: 0, feeCents: 0 });
  const opportunities: Opportunity[] = Array.from({ length: 14 }, (_, i) => ({ id: `OPP-${1001 + i}`, clientId: clients[stages[i % stages.length] === 'won' ? i : 180 + i].id, title: ['Consulta para compra de veículo', 'Histórico completo', 'Retorno de orçamento', 'Consulta para revenda'][i % 4], valueCents: i % 4 === 1 ? COMPLETE : BASE, stage: stages[i % stages.length], ownerId: clients[180 + i].ownerId, campaignId: i % 3 ? campaigns[i % 5].id : null, updatedAt: historicalAt(i) }));
  const instances: Instance[] = [
    { id: 'INS-01', name: 'Comercial • Oficial', number: '+55 11 4000-1200', type: 'official', status: 'connected', teamId: 'team-sales', lastSync: DEMO_NOW },
    { id: 'INS-02', name: 'Suporte • Oficial', number: '+55 11 4000-1201', type: 'official', status: 'connected', teamId: 'team-support', lastSync: DEMO_NOW },
    { id: 'INS-03', name: 'Relacionamento • Alternativa', number: '+55 11 99900-1202', type: 'alternative', status: 'disconnected', teamId: 'team-growth', lastSync: '2026-10-06T13:10:00.000Z' },
  ];
  const questions = ['Olá! A consulta informa restrição financeira?', 'Posso consultar a placa antes de comprar?', 'Meu pagamento foi aprovado. Como acesso o relatório?', 'Gostaria de consultar o histórico completo.', 'Vocês conseguem verificar a passagem por leilão?', 'Preciso de ajuda com a consulta da minha placa.'];
  const conversations: Conversation[] = Array.from({ length: 12 }, (_, i) => {
    const client = clients[i < 6 ? i : 180 + i - 6], instanceId = instances[i % 3].id;
    return { id: `CON-${1001 + i}`, clientId: client.id, instanceId, ownerId: i % 3 === 1 ? 'USR-06' : i === 8 ? null : client.ownerId, queue: i % 3 === 1 ? 'Suporte' : 'Comercial', status: i === 10 ? 'resolved' : i === 7 ? 'internal' : i % 3 === 2 ? 'waiting' : 'open', unread: i === 10 ? 0 : 1 + i % 3, priority: i === 2 || i === 6, messages: [
      { id: `MSG-${i + 1}-1`, text: 'Olá! Tudo bem? Somos a Verifica Placa.', kind: 'outgoing', at: historicalAt(i), author: users.find(u => u.id === client.ownerId)!.name, status: 'read' },
      { id: `MSG-${i + 1}-2`, text: questions[i % questions.length], kind: 'incoming', at: '2026-10-06T14:20:00.000Z', author: client.name, status: 'read' },
      ...(i === 2 ? [{ id: 'MSG-3-3', text: 'Retornaremos assim que a conexão for restabelecida.', kind: 'outgoing' as const, at: DEMO_NOW, author: 'Lucas Mendes', status: 'pending' as const }, { id: 'MSG-3-4', text: 'Nota interna: conferir a conexão antes de responder.', kind: 'note' as const, at: DEMO_NOW, author: 'Lucas Mendes', status: 'sent' as const }] : []),
    ] };
  });
  const taskTitles = ['Retornar orçamento de histórico completo', 'Conferir consulta em processamento', 'Acompanhar pagamento pendente', 'Responder dúvidas sobre leilão', 'Revisar relatório da campanha', 'Confirmar envio do relatório'];
  const tasks: Task[] = Array.from({ length: 18 }, (_, i) => ({ id: `TSK-${1001 + i}`, title: taskTitles[i % 6], clientId: i % 6 === 4 ? null : clients[180 + i % 14].id, ownerId: i % 6 === 4 ? 'USR-07' : i % 6 === 1 ? 'USR-06' : i % 2 ? 'USR-05' : 'USR-04', dueDate: `2026-10-${String(5 + i % 5).padStart(2, '0')}`, priority: i % 4 === 0 ? 'high' : i % 5 === 0 ? 'low' : 'normal', status: ['todo', 'doing', 'blocked', 'done'][i % 4] as Task['status'] }));
  const tickets: Ticket[] = Array.from({ length: 4 }, (_, i) => ({ id: `TIC-${1001 + i}`, title: ['Consulta sem retorno do bureau', 'Dúvida sobre relatório completo', 'Comprovante de estorno', 'Placa com divergência cadastral'][i], description: ['A consulta retornou falha. Conferir dados e reprocessar.', 'Cliente solicita explicação sobre passagem por leilão.', 'Confirmar o estorno registrado no pedido.', 'Validar a placa informada antes de uma nova consulta.'][i], clientId: orders[232 + i].clientId, orderId: orders[232 + i].id, ownerId: 'USR-06', status: i === 1 ? 'doing' : i === 2 ? 'resolved' : 'open', priority: i === 0 ? 'high' : 'normal', createdAt: historicalAt(i) }));
  const credits: Credit[] = [];
  for (const order of orders.filter(o => o.paidAt).slice(0, 30)) {
    const user = users.find(u => u.id === order.assistedBy)!;
    for (const [event, label, points] of [['paid', 'Pagamento confirmado', 20], ['delivered', 'Consulta entregue', 5], ...(order.addonCents ? [['upsell', 'Histórico adicional', 8]] : [])] as [string, string, number][]) {
      credits.push({ id: `CRD-${credits.length + 1001}`, eventId: `order:${order.id}:${event}`, userId: user.id, teamId: user.teamId, label, points, status: 'pending', at: order.paidAt! });
    }
  }
  // Historical September ledger is separate from October acquisition/finance totals.
  for (let i = 0; i < 60; i++) {
    const owner = users.find(u => u.id === (i % 2 ? 'USR-05' : 'USR-04'))!;
    const [event, label, points] = ([['paid', 'Pagamento confirmado', 20], ['upsell', 'Histórico adicional', 8], ['delivered', 'Consulta entregue', 5]] as const)[i % 3];
    credits.push({ id: nextId(credits, 'CRD'), eventId: `history:2026-09:${i + 1}:${event}`, userId: owner.id, teamId: owner.teamId, label: `${label} • histórico de setembro`, points, status: i % 4 === 0 ? 'pending' : 'confirmed', at: `2026-09-${String(1 + i % 28).padStart(2, '0')}T12:00:00.000Z` });
  }
  return { schemaVersion: 1, currentUserId: 'USR-01', users, clients, orders, opportunities, conversations, tasks, tickets, instances, campaigns,
    invites: [{ id: 'INV-1001', name: 'Amanda Castro', email: 'amanda@exemplo.demo', role: 'sales', teamId: 'team-sales', status: 'pending', expiresAt: '2026-10-13T14:30:00.000Z' }], credits,
    goals: [
      { id: 'GOL-1001', label: 'Receita de outubro', current: 620006, target: 1000000, unit: 'currency', scope: 'organization', ownerId: 'USR-01', description: 'Receita líquida após estornos, em centavos.' },
      { id: 'GOL-1002', label: 'Conversão de checkout', current: 20, target: 25, unit: 'percent', scope: 'organization', ownerId: 'USR-01', description: 'Pagamentos históricos sobre checkouts recebidos.' },
      { id: 'GOL-1003', label: 'Vendas da equipe', current: 240, target: 300, unit: 'count', scope: 'team', ownerId: 'team-sales', description: 'Pedidos com pagamento confirmado.' },
      { id: 'GOL-1004', label: 'Pontos de evolução', current: credits.filter(c => c.userId === 'USR-04' && c.status === 'confirmed').reduce((s, c) => s + c.points, 0), target: 1500, unit: 'count', scope: 'individual', ownerId: 'USR-04', description: 'Somente créditos confirmados compõem a pontuação.' },
    ],
    trackingEvents: Array.from({ length: 24 }, (_, i) => ({ id: `EVT-${1001 + i}`, clientId: clients[180 + i].id, event: ['page_view', 'begin_checkout', 'purchase', 'lead'][i % 4], source: ['Google Ads', 'Meta Ads', 'Orgânico'][i % 3], at: historicalAt(i), status: i === 19 ? 'failed' : 'received' })),
    audit: [{ id: 'AUD-1001', at: '2026-10-06T13:10:00.000Z', userId: 'USR-02', action: 'instance.disconnected', detail: 'Relacionamento • Alternativa perdeu a conexão.' }, { id: 'AUD-1002', at: '2026-10-06T13:20:00.000Z', userId: 'USR-08', action: 'order.refund', detail: 'Dois pedidos base estornados: R$ 29,98.' }], settings: { theme: 'dark', compact: false } };
}

/** Historical paidAt counts conversions, including later refunds. ROAS/CAC/CPA use attributed media only. Ratios use 0–1, not percentages. */
export function getMetrics(state: State, filters: MetricFilters = {}): Metrics {
  const inPeriod = (date: string | null) => !date || ((!filters.from || date.slice(0, 10) >= filters.from.slice(0, 10)) && (!filters.to || date.slice(0, 10) <= filters.to.slice(0, 10)));
  const teamOf = (order: Order) => state.users.find(u => u.id === (order.assistedBy ?? state.clients.find(c => c.id === order.clientId)?.ownerId))?.teamId;
  const orders = state.orders.filter(o => (!filters.campaignId || o.campaignId === filters.campaignId) && (!filters.teamId || teamOf(o) === filters.teamId) && inPeriod(o.paidAt ?? state.clients.find(c => c.id === o.clientId)?.createdAt ?? null));
  const paid = orders.filter(o => o.paidAt !== null), media = paid.filter(o => o.campaignId !== null);
  const total = (key: 'totalCents' | 'bureauCents' | 'feeCents' | 'addonCents', list = paid) => list.reduce((sum, o) => sum + o[key], 0);
  const revenueCents = total('totalCents'), refundsCents = total('totalCents', paid.filter(o => o.status === 'refunded')), netCents = revenueCents - refundsCents;
  // Campaigns carry a frozen Oct 1–6 aggregate. Allocate spend/checkouts uniformly by day for a partial-period demo view.
  const days = Array.from({ length: 6 }, (_, i) => `2026-10-0${i + 1}`).filter(inPeriod).length;
  const campaignList = state.campaigns.filter(c => !filters.campaignId || c.id === filters.campaignId);
  const allAttributed = state.orders.filter(o => o.paidAt && o.campaignId && (!filters.campaignId || o.campaignId === filters.campaignId) && inPeriod(o.paidAt));
  const teamShare = filters.teamId ? (allAttributed.length ? media.length / allAttributed.length : 0) : 1;
  const adsCents = Math.round(campaignList.reduce((s, c) => s + c.adsCents, 0) * days / 6 * teamShare);
  const checkouts = campaignList.reduce((s, c) => s + c.checkouts, 0) * days / 6 * teamShare;
  const firstPayment = new Map<string, Order>();
  for (const order of state.orders.filter(o => o.paidAt).sort((a, b) => a.paidAt!.localeCompare(b.paidAt!) || a.id.localeCompare(b.id))) {
    if (!firstPayment.has(order.clientId)) firstPayment.set(order.clientId, order);
  }
  const buyerCount = (list: Order[]) => new Set(list.filter(o => state.clients.find(c => c.id === o.clientId)?.newBuyer && firstPayment.get(o.clientId)?.id === o.id).map(o => o.clientId)).size;
  const mediaBuyers = buyerCount(media), bureauCents = total('bureauCents'), feeCents = total('feeCents'), taxCents = Math.round(netCents * 0.08);
  const contributionCents = netCents - adsCents - bureauCents - feeCents - taxCents;
  const divide = (n: number, d: number) => d ? n / d : 0;
  return { revenueCents, refundsCents, netCents, adsCents, bureauCents, feeCents, taxCents, contributionCents, margin: divide(contributionCents, netCents), paidOrders: paid.length, newBuyers: buyerCount(paid), ticketCents: Math.round(divide(revenueCents, paid.length)), upsellRate: divide(paid.filter(o => o.product === 'base' && o.addonCents > 0).length, paid.filter(o => o.product === 'base').length), checkoutRate: divide(paid.length, checkouts), cacCents: Math.round(divide(adsCents, mediaBuyers)), cpaCents: Math.round(divide(adsCents, media.length)), roas: divide(total('totalCents', media), adsCents), pendingOrders: orders.filter(o => o.status === 'pending').length, failedConsultations: orders.filter(o => o.consultationStatus === 'failed').length };
}

function fail(message: string): never { throw new Error(message); }
function required<T extends { id: string }>(list: T[], id: string, label: string): T { return list.find(item => item.id === id) ?? fail(`${label} não encontrado.`); }
function nextId(list: { id: string }[], prefix: string, floor = 1000): string { return `${prefix}-${Math.max(floor, ...list.map(item => Number(item.id.split('-').at(-1)) || 0)) + 1}`; }
function nonempty(value: unknown, label: string): string { const text = typeof value === 'string' ? value.trim() : ''; return text || fail(`Informe ${label}.`); }
function validCents(value: unknown): number { return typeof value === 'number' && Number.isFinite(value) && value >= 0 ? Math.round(value) : fail('Valor inválido. Informe centavos positivos.'); }
function validRole(value: string): Role { return roles.includes(value as Role) ? value as Role : fail('Perfil inválido.'); }
function validPriority(value: string | undefined): Priority { return value === undefined ? 'normal' : ['high', 'normal', 'low'].includes(value) ? value as Priority : fail('Prioridade inválida.'); }

export function reducer(state: State, action: Action): State {
  if (action.type === 'demo.reset') return createSeed();
  if (action.type === 'demo.switchUser') { const user = required(state.users, action.payload?.id, 'Usuário'); if (!user.active) fail('Usuário desativado.'); return { ...state, currentUserId: user.id }; }
  const user = required(state.users, state.currentUserId, 'Usuário');
  if (!user.active) fail('Usuário desativado.');
  const permission = (name: string) => { if (!can(user.role, name, state)) fail('Seu perfil não tem permissão para esta ação.'); };
  const global = user.role === 'owner' || user.role === 'admin';
  const accessibleOwner = (id: string) => global || id === user.id || (user.role === 'manager' && state.users.find(u => u.id === id)?.teamId === user.teamId);
  const ensureOwner = (id: string) => { if (!accessibleOwner(id)) fail('Este registro pertence a outro responsável ou equipe.'); };
  const ensureClient = (id: string, allowTeam = false) => {
    const client = required(state.clients, id, 'Cliente');
    const supportAssignment = user.role === 'support' && state.conversations.some(c => c.clientId === id && (c.ownerId === user.id || state.instances.find(i => i.id === c.instanceId)?.teamId === user.teamId));
    const sameTeam = allowTeam && state.users.find(u => u.id === client.ownerId)?.teamId === user.teamId;
    if (!accessibleOwner(client.ownerId) && !supportAssignment && !sameTeam) fail('Este cliente pertence a outro responsável ou equipe.');
    return client;
  };
  const ensureOrder = (order: Order) => { if (!global && user.role !== 'finance' && !accessibleOwner(order.assistedBy ?? '') ) ensureClient(order.clientId); };
  const s: State = structuredClone(state), p = action.payload ?? {};
  const at = new Date(Date.parse(DEMO_NOW) + state.audit.length * 1000).toISOString();
  let changed = true, detail = '';
  const credit = (order: Order, event: string, label: string, points: number) => {
    const eventId = `order:${order.id}:${event}`;
    if (s.credits.some(c => c.eventId === eventId)) return;
    const beneficiary = event === 'delivered' ? user.id : order.assistedBy;
    if (!beneficiary) return; // Automatic sales are revenue events, without a salesperson credit.
    const owner = required(s.users, beneficiary, 'Responsável');
    s.credits.push({ id: nextId(s.credits, 'CRD'), eventId, userId: owner.id, teamId: owner.teamId, label, points, status: 'pending', at });
  };
  switch (action.type) {
    case 'settings.update':
      if (p.theme !== undefined && !['dark', 'light'].includes(p.theme)) fail('Tema inválido.');
      if (p.compact !== undefined && typeof p.compact !== 'boolean') fail('Preferência inválida.');
      s.settings = { theme: p.theme ?? s.settings.theme, compact: p.compact ?? s.settings.compact }; break;
    case 'crm.create': {
      permission('crm.write'); ensureClient(p.clientId); const ownerId = p.ownerId ?? user.id; ensureOwner(ownerId); required(s.users, ownerId, 'Responsável');
      if (p.campaignId) required(s.campaigns, p.campaignId, 'Campanha');
      s.opportunities.push({ id: nextId(s.opportunities, 'OPP'), clientId: p.clientId, title: nonempty(p.title, 'o título'), valueCents: validCents(p.valueCents ?? BASE), stage: 'new', ownerId, campaignId: p.campaignId ?? null, updatedAt: at }); break;
    }
    case 'crm.move': {
      permission('crm.write'); const opportunity = required(s.opportunities, p.id, 'Oportunidade'); ensureOwner(opportunity.ownerId);
      if (!stages.includes(p.stage)) fail('Etapa inválida.');
      if (p.stage === 'won' && !s.orders.some(o => o.clientId === opportunity.clientId && o.status === 'paid' && o.paidAt)) fail('Para marcar como ganho, confirme o pagamento de um pedido vinculado ao cliente.');
      changed = opportunity.stage !== p.stage; opportunity.stage = p.stage; opportunity.updatedAt = at; break;
    }
    case 'client.create': {
      permission('crm.write'); const ownerId = p.ownerId ?? user.id; ensureOwner(ownerId); required(s.users, ownerId, 'Responsável');
      s.clients.push({ id: nextId(s.clients, 'CLI', 0), name: nonempty(p.name, 'o nome'), email: String(p.email ?? '').trim(), phone: String(p.phone ?? '').trim(), company: String(p.company ?? '').trim(), tag: 'Lead', origin: 'Cadastro manual', ownerId, createdAt: at, newBuyer: false }); break;
    }
    case 'order.create': {
      permission('orders.write'); ensureClient(p.clientId); const assistedBy = p.assistedBy === null ? null : p.assistedBy ?? user.id; if (assistedBy) { ensureOwner(assistedBy); required(s.users, assistedBy, 'Responsável'); }
      if (!['base', 'complete'].includes(p.product)) fail('Produto inválido.');
      if (p.campaignId) required(s.campaigns, p.campaignId, 'Campanha');
      const id = nextId(s.orders, 'ORD');
      const plate = p.plate === undefined ? `DEM${id.split('-').at(-1)!.padStart(4, '0').slice(-4)}` : nonempty(p.plate, 'uma placa válida').toUpperCase().replace(/[-\s]/g, '');
      if (!/^[A-Z]{3}(?:[0-9]{4}|[0-9][A-Z][0-9]{2})$/.test(plate)) fail('Placa inválida. Use o formato ABC1234 ou ABC1D23.');
      s.orders.push({ id, clientId: p.clientId, product: p.product, totalCents: p.product === 'base' ? BASE : COMPLETE, addonCents: 0, status: 'pending', paidAt: null, campaignId: p.campaignId ?? null, assistedBy, consultationStatus: 'waiting', plate, bureauCents: 0, feeCents: 0 }); break;
    }
    case 'order.pay': {
      permission('orders.write'); const order = required(s.orders, p.id, 'Pedido'); ensureOrder(order);
      if (order.status === 'paid') { changed = false; break; }
      if (order.status === 'refunded') fail('Pedido estornado não pode ser pago novamente.');
      const firstPurchase = !s.orders.some(o => o.clientId === order.clientId && o.paidAt !== null);
      const client = required(s.clients, order.clientId, 'Cliente'); if (firstPurchase) client.newBuyer = true; client.tag = 'Comprador';
      order.status = 'paid'; order.paidAt = at; order.consultationStatus = 'processing'; order.bureauCents = 375; order.feeCents = 75;
      s.opportunities.filter(o => o.clientId === order.clientId && !['won', 'lost'].includes(o.stage) && accessibleOwner(o.ownerId)).forEach(o => { o.stage = 'won'; o.updatedAt = at; });
      credit(order, 'paid', 'Pagamento confirmado', 20); break;
    }
    case 'order.refund': {
      permission('finance.refund'); const order = required(s.orders, p.id, 'Pedido');
      if (order.status === 'refunded') { changed = false; break; }
      if (order.status !== 'paid') fail('Somente um pedido pago pode ser estornado.');
      order.status = 'refunded'; order.consultationStatus = 'cancelled';
      s.credits.filter(c => c.eventId.startsWith(`order:${order.id}:`)).forEach(c => { c.status = 'reverted'; c.reason = 'Pedido estornado'; }); break;
    }
    case 'order.upsell': {
      permission('orders.write'); const order = required(s.orders, p.id, 'Pedido'); ensureOrder(order);
      if (order.status !== 'paid') fail('O adicional exige um pedido pago.');
      if (order.product !== 'base') fail('O adicional está disponível para consultas base.');
      if (order.addonCents > 0) { changed = false; break; }
      order.addonCents = ADDON; order.totalCents += ADDON; credit(order, 'upsell', 'Histórico adicional', 8); break;
    }
    case 'consultation.retry':
    case 'consultation.deliver': {
      permission('view.consultations'); permission('consultations.write'); const order = required(s.orders, p.id, 'Pedido'); if (user.role !== 'support') ensureOrder(order);
      if (order.status !== 'paid') fail('A consulta exige um pedido pago.');
      if (order.consultationStatus === 'delivered') { changed = false; break; }
      if (action.type === 'consultation.retry') { changed = order.consultationStatus !== 'processing'; order.consultationStatus = 'processing'; }
      else { order.consultationStatus = 'delivered'; credit(order, 'delivered', 'Consulta entregue', 5); } break;
    }
    case 'conversation.create': {
      permission('inbox.write'); ensureClient(p.clientId, true);
      const instance = required(s.instances, p.instanceId, 'Instância');
      if (!global && instance.teamId !== user.teamId) fail('Esta instância pertence a outra equipe.');
      if (s.conversations.some(c => c.clientId === p.clientId && c.instanceId === instance.id && c.status === 'open')) { changed = false; break; }
      s.conversations.push({ id: nextId(s.conversations, 'CON'), clientId: p.clientId, instanceId: instance.id, ownerId: user.id, queue: p.queue === undefined ? (user.role === 'support' ? 'Suporte' : 'Comercial') : nonempty(p.queue, 'a fila'), status: 'open', unread: 0, priority: false, messages: [] }); break;
    }
    case 'conversation.assign':
    case 'conversation.send':
    case 'conversation.note':
    case 'conversation.resolve':
    case 'conversation.reopen': {
      permission('inbox.write'); const conversation = required(s.conversations, p.id, 'Conversa');
      const instance = required(s.instances, conversation.instanceId, 'Instância');
      const inQueue = !conversation.ownerId && instance.teamId === user.teamId;
      const inboxOwnerAllowed = (id: string) => global || s.users.find(u => u.id === id)?.teamId === user.teamId;
      if (!global && !inQueue && !inboxOwnerAllowed(conversation.ownerId ?? '') && !(user.role === 'manager' && instance.teamId === user.teamId)) fail('Esta conversa pertence a outro responsável ou equipe.');
      if (action.type === 'conversation.assign') {
        if (p.ownerId !== null) { const owner = required(s.users, p.ownerId, 'Responsável'); if (!owner.active || !can(owner.role, 'inbox.write')) fail('Responsável indisponível para atendimento.'); if (!inboxOwnerAllowed(owner.id)) fail('Este responsável pertence a outra equipe.'); }
        changed = conversation.ownerId !== p.ownerId; conversation.ownerId = p.ownerId;
      } else if (action.type === 'conversation.resolve' || action.type === 'conversation.reopen') {
        const status = action.type === 'conversation.resolve' ? 'resolved' : 'open'; changed = conversation.status !== status; conversation.status = status; if (status === 'resolved') conversation.unread = 0;
      } else {
        const note = action.type === 'conversation.note';
        if (!note && conversation.status === 'resolved') fail('Reabra a conversa antes de enviar uma mensagem.');
        if (!note && instance.type === 'official') {
          if (p.templateId && !(APPROVED_TEMPLATES as readonly string[]).includes(p.templateId)) fail('Modelo de mensagem não aprovado.');
          const lastIncoming = conversation.messages.filter(m => m.kind === 'incoming').map(m => Date.parse(m.at)).filter(Number.isFinite).sort((a, b) => b - a)[0];
          const windowOpen = lastIncoming !== undefined && Date.parse(DEMO_NOW) - lastIncoming <= 24 * 3600000;
          if (!windowOpen && !p.templateId) fail('A janela de 24 horas está encerrada. Use um modelo aprovado.');
        }
        conversation.messages.push({ id: nextId(s.conversations.flatMap(c => c.messages), 'MSG'), text: nonempty(p.text, 'a mensagem'), kind: note ? 'note' : 'outgoing', at, author: user.name, status: note || instance.status === 'connected' ? 'sent' : 'pending' });
        if (!note) { conversation.unread = 0; conversation.status = instance.status === 'connected' ? 'open' : 'waiting'; }
      } break;
    }
    case 'task.create': {
      permission('tasks.write'); const ownerId = p.ownerId ?? user.id; ensureOwner(ownerId); required(s.users, ownerId, 'Responsável');
      if (p.clientId) ensureClient(p.clientId);
      const dueDate = p.dueDate ?? '2026-10-07'; if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate) || Number.isNaN(Date.parse(dueDate))) fail('Data inválida.');
      s.tasks.push({ id: nextId(s.tasks, 'TSK'), title: nonempty(p.title, 'o título'), clientId: p.clientId ?? null, ownerId, dueDate, priority: validPriority(p.priority), status: 'todo' }); break;
    }
    case 'task.update': {
      permission('tasks.write'); const task = required(s.tasks, p.id, 'Tarefa'); ensureOwner(task.ownerId); if (!['todo', 'doing', 'blocked', 'done'].includes(p.status)) fail('Estado da tarefa inválido.'); changed = task.status !== p.status; task.status = p.status; break;
    }
    case 'ticket.create': {
      permission('support.write');
      const order = p.orderId ? required(s.orders, p.orderId, 'Pedido') : undefined;
      if (order && p.clientId && order.clientId !== p.clientId) fail('Pedido e cliente não correspondem.');
      const consultationQueue = user.role === 'support' && order?.status === 'paid' && can(user.role, 'view.consultations', state) && can(user.role, 'consultations.write', state);
      if (p.clientId && !consultationQueue) ensureClient(p.clientId);
      if (order && !consultationQueue) ensureOrder(order);
      s.tickets.push({ id: nextId(s.tickets, 'TIC'), title: nonempty(p.title, 'o título'), description: nonempty(p.description, 'a descrição'), clientId: p.clientId ?? order?.clientId ?? null, orderId: p.orderId ?? null, ownerId: user.id, status: 'open', priority: validPriority(p.priority), createdAt: at }); break;
    }
    case 'ticket.resolve': { permission('support.write'); const ticket = required(s.tickets, p.id, 'Chamado'); ensureOwner(ticket.ownerId); changed = ticket.status !== 'resolved'; ticket.status = 'resolved'; break; }
    case 'instance.toggle': {
      permission('admin.integrations'); const instance = required(s.instances, p.id, 'Instância'); instance.status = instance.status === 'connected' ? 'disconnected' : 'connected'; instance.lastSync = at;
      if (instance.status === 'connected') s.conversations.filter(c => c.instanceId === instance.id).forEach(c => { c.messages.filter(m => m.kind === 'outgoing' && m.status === 'pending').forEach(m => { m.status = 'sent'; }); if (c.status === 'waiting') c.status = 'open'; }); break;
    }
    case 'invite.create': {
      permission('admin.users'); const email = nonempty(p.email, 'o e-mail').toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail('E-mail inválido.');
      if (s.users.some(u => u.email.toLowerCase() === email) || s.invites.some(i => i.email.toLowerCase() === email && i.status === 'pending')) fail('Já existe um usuário ou convite pendente para este e-mail.');
      const teamId = nonempty(p.teamId, 'a equipe'); if (!s.users.some(u => u.teamId === teamId)) fail('Equipe inválida.');
      s.invites.push({ id: nextId(s.invites, 'INV'), name: nonempty(p.name, 'o nome'), email, role: validRole(p.role), teamId, status: 'pending', expiresAt: new Date(Date.parse(at) + 7 * 86400000).toISOString() }); break;
    }
    case 'invite.accept':
    case 'invite.revoke': {
      permission('admin.users'); const invite = required(s.invites, p.id, 'Convite');
      if (action.type === 'invite.revoke') { if (invite.status === 'revoked') { changed = false; break; } if (invite.status === 'accepted') fail('Um convite aceito não pode ser revogado; desative o usuário.'); invite.status = 'revoked'; }
      else { if (invite.status === 'accepted') { changed = false; break; } if (invite.status !== 'pending' || invite.expiresAt <= at) fail('Convite indisponível ou expirado.');
        if (s.users.some(u => u.email.toLowerCase() === invite.email.toLowerCase())) fail('Já existe um usuário para este e-mail.');
        invite.status = 'accepted'; s.users.push({ id: nextId(s.users, 'USR', 0), name: invite.name, email: invite.email, role: invite.role, teamId: invite.teamId, title: 'Novo integrante', initials: initials(invite.name), active: true }); }
      break;
    }
    case 'permissions.update': {
      permission('admin.users'); const role = validRole(p.role);
      if (role === 'owner') fail('As permissões do proprietário são protegidas.');
      const known = [...views.map(v => `view.${v}`), ...writePermissions];
      if (!known.includes(p.permission) || typeof p.enabled !== 'boolean') fail('Permissão inválida.');
      if (p.enabled && !can(user.role, p.permission, state)) fail('Você não pode conceder uma permissão que não possui.');
      const current = [...(s.rolePermissions?.[role] ?? grants[role])];
      s.rolePermissions ??= {};
      s.rolePermissions[role] = p.enabled ? [...new Set([...current, p.permission])] : current.filter(x => x !== p.permission);
      break;
    }
    case 'user.update': {
      permission('admin.users'); const target = required(s.users, p.id, 'Usuário');
      if (target.role === 'owner' || p.role === 'owner') fail('A propriedade da organização não é alterada neste fluxo.');
      if (target.id === user.id) fail('Seu próprio acesso não pode ser alterado neste fluxo.');
      const role = validRole(p.role); if (!s.users.some(u => u.teamId === p.teamId)) fail('Equipe inválida.');
      const changingAccess = role !== target.role || p.teamId !== target.teamId;
      const unfinished = s.tasks.some(t => t.ownerId === target.id && t.status !== 'done') || s.opportunities.some(o => o.ownerId === target.id && !['won', 'lost'].includes(o.stage)) || s.conversations.some(c => c.ownerId === target.id && c.status !== 'resolved') || s.tickets.some(t => t.ownerId === target.id && t.status !== 'resolved');
      if (changingAccess && unfinished) fail('Resolva ou redistribua as pendências deste usuário antes de alterar sua equipe ou perfil.');
      target.role = role; target.teamId = p.teamId; target.title = nonempty(p.title, 'a função'); break;
    }
    case 'user.toggle': {
      permission('admin.users'); const target = required(s.users, p.id, 'Usuário'); if (target.id === user.id) fail('Não é possível desativar o próprio usuário.'); if (target.role === 'owner' && target.active && s.users.filter(u => u.active && u.role === 'owner').length === 1) fail('Mantenha ao menos um proprietário ativo.'); target.active = !target.active;
      if (!target.active) {
        const replacementFor = (write: string, view: string) => {
          const eligible = s.users.filter(u => u.active && u.id !== target.id && can(u.role, write, s) && can(u.role, view, s));
          return eligible.find(u => u.teamId === target.teamId && u.role === 'manager') || eligible.find(u => u.teamId === target.teamId) || eligible.find(u => u.id === user.id) || eligible.find(u => u.role === 'owner' || u.role === 'admin') || fail('Não há responsável ativo e autorizado para redistribuir as pendências.');
        };
        s.tasks.filter(t => t.ownerId === target.id && t.status !== 'done').forEach(t => t.ownerId = replacementFor('tasks.write', 'view.tasks').id);
        s.opportunities.filter(o => o.ownerId === target.id && !['won', 'lost'].includes(o.stage)).forEach(o => o.ownerId = replacementFor('crm.write', 'view.crm').id);
        s.conversations.filter(c => c.ownerId === target.id && c.status !== 'resolved').forEach(c => c.ownerId = replacementFor('inbox.write', 'view.inbox').id);
        s.tickets.filter(t => t.ownerId === target.id && t.status !== 'resolved').forEach(t => t.ownerId = replacementFor('support.write', 'view.support').id);
      }
      break;
    }
    case 'profile.update': {
      permission('profile.write'); const target = required(s.users, user.id, 'Usuário'); if (p.name !== undefined) { target.name = nonempty(p.name, 'o nome'); target.initials = initials(target.name); }
      if (p.email !== undefined) { const email = nonempty(p.email, 'o e-mail').toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail('E-mail inválido.'); if (s.users.some(u => u.id !== target.id && u.email.toLowerCase() === email)) fail('Este e-mail já está em uso.'); target.email = email; }
      if (p.title !== undefined) target.title = nonempty(p.title, 'o cargo');
      if (p.avatar !== undefined) { if (p.avatar !== '' && !/^data:image\/(png|jpeg|webp);base64,/.test(p.avatar)) fail('Imagem de perfil inválida.'); if (p.avatar.length > 1500000) fail('Imagem de perfil muito grande.'); target.avatar = p.avatar; }
      break;
    }
    case 'goal.update': {
      permission('goals.write'); const goal = required(s.goals, p.id, 'Meta');
      if (!global && !(goal.scope === 'team' && goal.ownerId === user.teamId) && !(goal.scope === 'individual' && accessibleOwner(goal.ownerId))) fail('Esta meta pertence a outro responsável ou equipe.');
      goal.target = validCents(p.target); if (goal.target <= 0) fail('A meta deve ser maior que zero.'); break;
    }
    case 'campaign.toggle': { permission('marketing.send'); const campaign = required(s.campaigns, p.id, 'Campanha'); campaign.status = campaign.status === 'active' ? 'paused' : 'active'; break; }
    case 'credit.confirm': {
      permission('goals.write'); const item = required(s.credits, p.id, 'Crédito'); ensureOwner(item.userId);
      if (item.status === 'confirmed') { changed = false; break; }
      if (item.status === 'reverted') fail('Crédito revertido não pode ser confirmado.');
      const cutoff = Date.parse(DEMO_NOW) - 7 * 86400000;
      const historical = /^history:2026-09:([1-9]|[1-5][0-9]|60):(paid|upsell|delivered)$/.test(item.eventId) && item.at.startsWith('2026-09-');
      const [, orderId, event] = item.eventId.split(':'); const order = s.orders.find(o => o.id === orderId);
      const eligibleOrder = order && order.status === 'paid' && order.paidAt && Date.parse(order.paidAt) <= cutoff && (event !== 'delivered' || order.consultationStatus === 'delivered') && (event !== 'upsell' || order.addonCents > 0) && ['paid', 'delivered', 'upsell'].includes(event);
      if (!Number.isFinite(Date.parse(item.at)) || Date.parse(item.at) > cutoff || (!historical && !eligibleOrder)) fail('Crédito ainda não elegível: aguarde 7 dias e a validação do histórico.');
      item.status = 'confirmed'; break;
    }
    default: fail('Ação desconhecida.');
  }
  if (!changed) return state;
  const metrics = getMetrics(s);
  s.goals.forEach(goal => {
    if (goal.id === 'GOL-1001') goal.current = metrics.netCents;
    else if (goal.id === 'GOL-1002') goal.current = metrics.checkoutRate * 100;
    else if (goal.id === 'GOL-1003') goal.current = getMetrics(s, { teamId: goal.ownerId }).paidOrders;
    else if (goal.id === 'GOL-1004') goal.current = s.credits.filter(c => c.userId === goal.ownerId && c.status === 'confirmed').reduce((sum, c) => sum + c.points, 0);
  });
  detail = p.id ? String(p.id) : p.title ? String(p.title) : p.name ? String(p.name) : action.type;
  s.audit.push({ id: nextId(s.audit, 'AUD'), at, userId: user.id, action: action.type, detail });
  return s;
}
