import test from 'node:test';
import assert from 'node:assert/strict';
import { APPROVED_TEMPLATES, DEMO_NOW, can, createSeed, getMetrics, reducer, type State } from './engine';

const act = (s: State, type: string, payload?: unknown) => reducer(s, { type, payload });
const asUser = (s: State, id: string) => act(s, 'demo.switchUser', { id });

test('seed is deterministic and coherent across finance, acquisition and operations', () => {
  const s = createSeed(), m = getMetrics(s);
  assert.deepEqual(s, createSeed());
  assert.equal(s.users.length, 8);
  assert.equal(new Set(s.users.map(u => u.teamId)).size, 3);
  assert.equal(s.orders.filter(o => o.paidAt).length, 240);
  assert.equal(new Set(s.orders.filter(o => o.paidAt).map(o => o.clientId)).size, 180);
  assert.equal(s.clients.length, 480);
  assert.ok(s.clients.slice(0, 12).every(c => c.tag === 'Recorrente' && !c.newBuyer));
  assert.equal(s.opportunities.length, 14);
  assert.equal(s.conversations.length, 12);
  assert.equal(s.tasks.length, 18);
  assert.equal(s.tickets.length, 4);
  assert.equal(s.instances.filter(i => i.type === 'official').length, 2);
  assert.deepEqual(['delivered', 'processing', 'failed', 'cancelled'].map(status => s.orders.filter(o => o.paidAt && o.consultationStatus === status).length), [220, 12, 6, 2]);
  assert.equal(m.revenueCents, 623004);
  assert.equal(m.refundsCents, 2998);
  assert.equal(m.netCents, 620006);
  assert.equal(m.paidOrders, 240);
  assert.equal(m.adsCents, 180000);
  assert.equal(m.bureauCents, 90000);
  assert.equal(m.feeCents, 18000);
  assert.equal(m.taxCents, 49600);
  assert.equal(m.contributionCents, 282406);
  assert.equal(m.newBuyers, 120);
  assert.equal(m.upsellRate, 54 / 216);
  assert.deepEqual(s.orders.filter(o => o.consultationStatus === 'cancelled').map(o => o.id).sort(), s.orders.filter(o => o.status === 'refunded').map(o => o.id).sort());
  assert.ok(s.credits.filter(c => c.status === 'confirmed').every(c => c.eventId.startsWith('history:2026-09:') && Date.parse(c.at) <= Date.parse(DEMO_NOW) - 7 * 86400000));
  assert.ok(s.credits.filter(c => c.eventId.startsWith('order:')).every(c => c.status === 'pending'));
  assert.equal(m.checkoutRate, 0.2);
  assert.equal(m.cacCents, 2000);
  assert.equal(m.cpaCents, 1000);
  assert.ok(Math.abs(m.roas - 2.6624444444444446) < 1e-12);
  assert.equal(m.pendingOrders, 4);
  assert.equal(m.failedConsultations, 6);
  assert.deepEqual(getMetrics(s, { from: '2026-10-01', to: '2026-10-06' }), m);
  assert.equal(s.orders.filter(o => o.campaignId && o.paidAt).reduce((sum, o) => sum + o.totalCents, 0), 479240);
});

test('campaign/team/date filters derive metrics without NaN or duplicated spend', () => {
  const s = createSeed(), m = getMetrics(s);
  const campaigns = s.campaigns.map(c => getMetrics(s, { campaignId: c.id }));
  assert.equal(campaigns.reduce((sum, c) => sum + c.adsCents, 0), m.adsCents);
  assert.equal(campaigns.reduce((sum, c) => sum + c.paidOrders, 0), 180);
  assert.equal(campaigns.reduce((sum, c) => sum + c.revenueCents, 0), 479240);
  assert.equal(getMetrics(s, { teamId: 'team-sales' }).revenueCents, m.revenueCents);
  assert.equal(getMetrics(s, { from: '2026-10-01', to: '2026-10-01' }).adsCents, 30000);
  assert.equal(Array.from({ length: 6 }, (_, i) => getMetrics(s, { from: `2026-10-0${i + 1}`, to: `2026-10-0${i + 1}` }).newBuyers).reduce((sum, n) => sum + n, 0), 120);
  assert.equal(getMetrics(s, { from: '2027-01-01', to: '2027-01-31' }).revenueCents, 0);
  assert.ok(Object.values(getMetrics(s, { campaignId: 'missing' })).every(Number.isFinite));
});

test('pay is atomic, advances related opportunity and issues one pending credit', () => {
  const original = createSeed();
  const paid = act(original, 'order.pay', { id: 'ORD-1241' });
  assert.equal(original.orders.find(o => o.id === 'ORD-1241')!.status, 'pending');
  const order = paid.orders.find(o => o.id === 'ORD-1241')!;
  assert.equal(order.status, 'paid');
  assert.equal(order.consultationStatus, 'processing');
  assert.equal(order.bureauCents, 375);
  assert.equal(paid.opportunities.find(o => o.clientId === order.clientId)!.stage, 'won');
  assert.equal(paid.clients.find(c => c.id === order.clientId)!.newBuyer, true);
  assert.deepEqual(paid.credits.filter(c => c.eventId === 'order:ORD-1241:paid').map(c => [c.points, c.status]), [[20, 'pending']]);
  assert.equal(act(paid, 'order.pay', { id: order.id }), paid);
  assert.equal(getMetrics(paid).paidOrders, 241);
  assert.equal(getMetrics(paid).revenueCents, 624503);
});

test('upsell/delivery credit once and refund reverses every related event', () => {
  let s = act(createSeed(), 'order.pay', { id: 'ORD-1241' });
  s = act(s, 'order.upsell', { id: 'ORD-1241' });
  assert.equal(act(s, 'order.upsell', { id: 'ORD-1241' }), s);
  assert.equal(s.orders.find(o => o.id === 'ORD-1241')!.totalCents, 3489);
  s = act(s, 'consultation.deliver', { id: 'ORD-1241' });
  assert.equal(act(s, 'consultation.deliver', { id: 'ORD-1241' }), s);
  assert.deepEqual(s.credits.filter(c => c.eventId.startsWith('order:ORD-1241:')).map(c => c.points), [20, 8, 5]);
  const before = getMetrics(s);
  s = act(s, 'order.refund', { id: 'ORD-1241' });
  assert.equal(act(s, 'order.refund', { id: 'ORD-1241' }), s);
  assert.equal(getMetrics(s).revenueCents, before.revenueCents);
  assert.equal(getMetrics(s).refundsCents, before.refundsCents + 3489);
  assert.equal(getMetrics(s).netCents, before.netCents - 3489);
  assert.ok(s.credits.filter(c => c.eventId.startsWith('order:ORD-1241:')).every(c => c.status === 'reverted'));
  assert.throws(() => act(s, 'order.pay', { id: 'ORD-1241' }), /estornado/);
});

test('permissions deny commands as well as navigation; owner and manager scope remains enforced', () => {
  const seed = createSeed();
  assert.equal(can('sales', 'finance.refund'), false);
  assert.equal(can('finance', 'finance.refund'), true);
  assert.equal(can('finance', 'orders.write'), false);
  assert.equal(can('manager', 'admin.integrations'), false);
  assert.equal(can('marketing', 'view.inbox'), false);
  assert.equal(can('owner', 'does.not.exist'), false);
  assert.throws(() => act(asUser(seed, 'USR-04'), 'order.refund', { id: 'ORD-1001' }), /permissão/);
  assert.throws(() => act(asUser(seed, 'USR-08'), 'order.pay', { id: 'ORD-1241' }), /permissão/);
  assert.throws(() => act(asUser(seed, 'USR-04'), 'task.update', { id: 'TSK-1002', status: 'done' }), /outro responsável/);
  assert.throws(() => act(asUser(seed, 'USR-04'), 'crm.move', { id: 'OPP-1002', stage: 'won' }), /outro responsável/);
  assert.throws(() => act(asUser(seed, 'USR-03'), 'task.update', { id: 'TSK-1002', status: 'done' }), /outro responsável/);
  assert.equal(act(asUser(seed, 'USR-03'), 'task.update', { id: 'TSK-1004', status: 'todo' }).tasks.find(t => t.id === 'TSK-1004')!.status, 'todo');
  assert.equal(act(asUser(seed, 'USR-08'), 'order.refund', { id: 'ORD-1001' }).orders[0].status, 'refunded');
});

test('reconnection reconciles queued outgoing messages once and keeps notes internal', () => {
  let s = createSeed();
  s = act(s, 'conversation.send', { id: 'CON-1003', text: 'Seu relatório está disponível.' });
  s = act(s, 'conversation.note', { id: 'CON-1003', text: 'Conferir pedido antes do retorno.' });
  const before = s.conversations.find(c => c.id === 'CON-1003')!.messages;
  assert.equal(before.at(-2)!.status, 'pending');
  assert.equal(before.at(-1)!.kind, 'note');
  assert.equal(before.at(-1)!.status, 'sent');
  s = act(s, 'instance.toggle', { id: 'INS-03' });
  assert.equal(s.conversations.find(c => c.id === 'CON-1003')!.messages.length, before.length);
  assert.ok(s.conversations.find(c => c.id === 'CON-1003')!.messages.every(m => m.status !== 'pending'));
  s = act(act(s, 'instance.toggle', { id: 'INS-03' }), 'instance.toggle', { id: 'INS-03' });
  const after = s.conversations.find(c => c.id === 'CON-1003')!.messages;
  assert.equal(after.length, before.length);
  assert.equal(after.at(-1)!.kind, 'note');
});

test('official WhatsApp 24h window requires an approved template; internal notes remain available', () => {
  const s = createSeed();
  const conversation = s.conversations[0];
  conversation.messages.filter(m => m.kind === 'incoming').forEach(m => { m.at = '2026-10-04T15:00:00.000Z'; });
  assert.throws(() => act(s, 'conversation.send', { id: conversation.id, text: 'Olá' }), /24 horas/);
  assert.throws(() => act(s, 'conversation.send', { id: conversation.id, text: 'Olá', templateId: 'unapproved' }), /não aprovado/);
  assert.equal(act(s, 'conversation.send', { id: conversation.id, text: 'Olá', templateId: APPROVED_TEMPLATES[0] }).conversations[0].messages.at(-1)!.kind, 'outgoing');
  assert.equal(act(s, 'conversation.note', { id: conversation.id, text: 'Aguardar resposta' }).conversations[0].messages.at(-1)!.kind, 'note');
});

test('same-team inbox can be reassigned while another team remains blocked', () => {
  const s = asUser(createSeed(), 'USR-04');
  const otherSales = s.conversations.find(c => c.ownerId === 'USR-05')!;
  assert.equal(act(s, 'conversation.assign', { id: otherSales.id, ownerId: 'USR-04' }).conversations.find(c => c.id === otherSales.id)!.ownerId, 'USR-04');
  assert.throws(() => act(s, 'conversation.assign', { id: otherSales.id, ownerId: 'USR-06' }), /outra equipe/);
});

test('unassigned team queue can be claimed by its member, never by another team', () => {
  const s = createSeed();
  s.conversations[0].ownerId = null;
  assert.equal(act(asUser(s, 'USR-04'), 'conversation.assign', { id: s.conversations[0].id, ownerId: 'USR-04' }).conversations[0].ownerId, 'USR-04');
  assert.throws(() => act(asUser(s, 'USR-06'), 'conversation.assign', { id: s.conversations[0].id, ownerId: 'USR-06' }), /outro responsável/);
});

test('only valid historical pending credits can be confirmed; October credits wait 7 days', () => {
  let s = createSeed();
  const credit = s.credits.find(c => c.status === 'pending' && c.eventId.startsWith('history:2026-09:'))!;
  s = act(s, 'credit.confirm', { id: credit.id });
  assert.equal(s.credits.find(c => c.id === credit.id)!.status, 'confirmed');
  assert.equal(act(s, 'credit.confirm', { id: credit.id }), s);
  const october = s.credits.find(c => c.eventId.startsWith('order:'))!;
  assert.throws(() => act(s, 'credit.confirm', { id: october.id }), /7 dias/);
  const orderId = october.eventId.split(':')[1];
  s = act(s, 'order.refund', { id: orderId });
  assert.equal(s.credits.find(c => c.id === october.id)!.status, 'reverted');
  assert.throws(() => act(s, 'credit.confirm', { id: october.id }), /revertido/);
  s = act(s, 'order.pay', { id: 'ORD-1241' });
  const fresh = s.credits.find(c => c.eventId === 'order:ORD-1241:paid')!;
  assert.throws(() => act(s, 'credit.confirm', { id: fresh.id }), /histórico/);
});

test('invitation accepts once, expires after 7 days and protected owner cannot be deactivated', () => {
  let s = createSeed();
  s = act(s, 'invite.create', { name: 'Amanda Moura', email: 'amanda.moura@exemplo.demo', role: 'sales', teamId: 'team-sales' });
  const invite = s.invites.at(-1)!;
  const audit = s.audit.at(-1)!;
  assert.equal(Date.parse(invite.expiresAt) - Date.parse(audit.at), 7 * 86400000);
  s = act(s, 'invite.accept', { id: invite.id });
  assert.equal(s.users.length, 9);
  assert.equal(act(s, 'invite.accept', { id: invite.id }), s);
  assert.throws(() => act(s, 'user.toggle', { id: 'USR-01' }), /próprio usuário/);
  assert.throws(() => act(asUser(s, 'USR-02'), 'user.toggle', { id: 'USR-01' }), /proprietário ativo/);
});


test('conversation.create starts an owned conversation and is idempotent for open client/instance pairs', () => {
  let s = asUser(createSeed(), 'USR-04');
  s = act(s, 'client.create', { name: 'Paula Azevedo', ownerId: 'USR-04' });
  const clientId = s.clients.at(-1)!.id;
  const next = act(s, 'conversation.create', { clientId, instanceId: 'INS-01', queue: 'Novos contatos' });
  const conversation = next.conversations.at(-1)!;
  assert.deepEqual({ clientId: conversation.clientId, instanceId: conversation.instanceId, ownerId: conversation.ownerId, queue: conversation.queue, status: conversation.status, unread: conversation.unread, messages: conversation.messages }, { clientId, instanceId: 'INS-01', ownerId: 'USR-04', queue: 'Novos contatos', status: 'open', unread: 0, messages: [] });
  assert.equal(act(next, 'conversation.create', { clientId, instanceId: 'INS-01' }), next);
  assert.equal(next.conversations.length, s.conversations.length + 1);
  assert.throws(() => act(s, 'conversation.create', { clientId, instanceId: 'INS-02' }), /outra equipe/);
  assert.throws(() => act(asUser(s, 'USR-08'), 'conversation.create', { clientId, instanceId: 'INS-01' }), /permissão/);
});

test('conversation.create supports same-team sales clients and assigned support clients without weakening CRM scope', () => {
  const s = createSeed();
  const sales = asUser(s, 'USR-04');
  const colleagueClient = s.clients.find(c => c.ownerId === 'USR-05')!;
  assert.equal(act(sales, 'conversation.create', { clientId: colleagueClient.id, instanceId: 'INS-01' }).conversations.at(-1)!.ownerId, 'USR-04');
  assert.throws(() => act(sales, 'crm.create', { clientId: colleagueClient.id, title: 'Contato', valueCents: 1499 }), /outro responsável/);
  const support = asUser(s, 'USR-06');
  const supportClient = s.conversations.find(c => c.ownerId === 'USR-06')!.clientId;
  const existing = s.conversations.find(c => c.clientId === supportClient && c.instanceId === 'INS-02')!;
  existing.status = 'resolved';
  const next = act(asUser(s, 'USR-06'), 'conversation.create', { clientId: supportClient, instanceId: 'INS-02' });
  assert.equal(next.conversations.at(-1)!.ownerId, 'USR-06');
  assert.equal(next.conversations.at(-1)!.queue, 'Suporte');
  const unauthorizedClient = s.clients.find(c => !s.conversations.some(con => con.clientId === c.id && con.ownerId === 'USR-06'))!;
  assert.throws(() => act(support, 'conversation.create', { clientId: unauthorizedClient.id, instanceId: 'INS-02' }), /outro responsável/);
});


test('consultation writes require both grants and support can operate the organization queue without finance privileges', () => {
  const s = createSeed();
  for (const role of ['owner', 'admin', 'manager', 'support']) assert.equal(can(role, 'consultations.write'), true);
  for (const role of ['sales', 'marketing', 'finance']) assert.equal(can(role, 'consultations.write'), false);
  const failed = s.orders.find(o => o.status === 'paid' && o.consultationStatus === 'failed' && !s.conversations.some(c => c.clientId === o.clientId && c.ownerId === 'USR-06'))!;
  let support = asUser(s, 'USR-06');
  support = act(support, 'consultation.retry', { id: failed.id });
  assert.equal(support.orders.find(o => o.id === failed.id)!.consultationStatus, 'processing');
  support = act(support, 'consultation.deliver', { id: failed.id });
  const credit = support.credits.find(c => c.eventId === `order:${failed.id}:delivered`)!;
  assert.equal(credit.userId, 'USR-06');
  assert.equal(credit.teamId, 'team-support');
  assert.equal(credit.points, 5);
  assert.equal(credit.status, 'pending');
  assert.equal(act(support, 'consultation.deliver', { id: failed.id }), support);
  assert.throws(() => act(support, 'order.pay', { id: 'ORD-1241' }), /permissão/);
  assert.throws(() => act(support, 'order.refund', { id: failed.id }), /permissão/);
  const readOnly = act(s, 'permissions.update', { role: 'support', permission: 'consultations.write', enabled: false });
  assert.equal(can('support', 'view.consultations', readOnly), true);
  assert.throws(() => act(asUser(readOnly, 'USR-06'), 'consultation.retry', { id: failed.id }), /permissão/);
  const noView = act(s, 'permissions.update', { role: 'support', permission: 'view.consultations', enabled: false });
  assert.throws(() => act(asUser(noView, 'USR-06'), 'consultation.deliver', { id: failed.id }), /permissão/);
});

test('sale and upsell credits retain the assisted salesperson while automatic sales receive none', () => {
  let s = createSeed();
  s = act(s, 'order.pay', { id: 'ORD-1241' });
  s = act(s, 'order.upsell', { id: 'ORD-1241' });
  assert.ok(s.credits.filter(c => c.eventId.startsWith('order:ORD-1241:')).every(c => c.userId === 'USR-04'));
  s = act(s, 'order.create', { clientId: s.clients[0].id, product: 'base', assistedBy: null });
  const automaticId = s.orders.at(-1)!.id;
  assert.equal(s.orders.at(-1)!.assistedBy, null);
  s = act(s, 'order.pay', { id: automaticId });
  s = act(s, 'order.upsell', { id: automaticId });
  assert.equal(s.credits.filter(c => c.eventId.startsWith(`order:${automaticId}:`)).length, 0);
  s = act(asUser(s, 'USR-06'), 'consultation.deliver', { id: automaticId });
  assert.deepEqual(s.credits.filter(c => c.eventId.startsWith(`order:${automaticId}:`)).map(c => [c.userId, c.points]), [['USR-06', 5]]);
});

test('permissions.update applies immediately to reducer actions and protects owner permissions', () => {
  const seed = createSeed();
  const updated = act(seed, 'permissions.update', { role: 'sales', permission: 'orders.write', enabled: false });
  assert.equal(can('sales', 'orders.write', seed), true);
  assert.equal(can('sales', 'orders.write', updated), false);
  assert.throws(() => act(asUser(updated, 'USR-04'), 'order.pay', { id: 'ORD-1241' }), /permissão/);
  assert.throws(() => act(seed, 'permissions.update', { role: 'owner', permission: 'admin.users', enabled: false }), /proprietário/);
  assert.throws(() => act(asUser(seed, 'USR-04'), 'permissions.update', { role: 'sales', permission: 'finance.refund', enabled: true }), /permissão/);
  const restored = act(updated, 'permissions.update', { role: 'sales', permission: 'orders.write', enabled: true });
  assert.equal(act(asUser(restored, 'USR-04'), 'order.pay', { id: 'ORD-1241' }).orders.find(o => o.id === 'ORD-1241')!.status, 'paid');
});

test('user.update protects ownership and own access, preserving historical attribution on team and role changes', () => {
  const s = createSeed();
  assert.throws(() => act(s, 'user.update', { id: 'USR-01', role: 'admin', teamId: 'team-support', title: 'Admin' }), /propriedade/);
  assert.throws(() => act(s, 'user.update', { id: 'USR-04', role: 'owner', teamId: 'team-sales', title: 'Owner' }), /propriedade/);
  assert.throws(() => act(asUser(s, 'USR-02'), 'user.update', { id: 'USR-02', role: 'support', teamId: 'team-support', title: 'Suporte' }), /próprio acesso/);
  assert.throws(() => act(s, 'user.update', { id: 'USR-04', role: 'support', teamId: 'team-support', title: 'Especialista de atendimento' }), /redistribua as pendências/);
  const sameAccess = act(s, 'user.update', { id: 'USR-04', role: 'sales', teamId: 'team-sales', title: 'Consultor sênior' });
  assert.equal(sameAccess.users.find(u => u.id === 'USR-04')!.title, 'Consultor sênior');
  // Suspension already redistributes unfinished work; historical completed assignments remain intact.
  const redistributed = act(s, 'user.toggle', { id: 'USR-04' });
  const next = act(redistributed, 'user.update', { id: 'USR-04', role: 'support', teamId: 'team-support', title: 'Especialista de atendimento' });
  assert.equal(next.users.find(u => u.id === 'USR-04')!.role, 'support');
  assert.equal(next.users.find(u => u.id === 'USR-04')!.teamId, 'team-support');
  assert.deepEqual(next.orders, s.orders);
  assert.deepEqual(next.credits, s.credits);
});

test('suspension redistributes unfinished work to authorized active users and preserves historical orders and credits', () => {
  const s = createSeed();
  s.tasks[0].status = 'done';
  s.tickets[0].ownerId = 'USR-04';
  s.tickets[1].ownerId = 'USR-04'; s.tickets[1].status = 'resolved';
  const next = act(s, 'user.toggle', { id: 'USR-04' });
  assert.equal(next.users.find(u => u.id === 'USR-04')!.active, false);
  assert.deepEqual(next.orders, s.orders);
  assert.deepEqual(next.credits, s.credits);
  assert.deepEqual(next.clients, s.clients);
  const groups = [
    [next.tasks.filter(t => t.status !== 'done'), 'tasks.write'],
    [next.opportunities.filter(o => !['won', 'lost'].includes(o.stage)), 'crm.write'],
    [next.conversations.filter(c => c.status !== 'resolved'), 'inbox.write'],
    [next.tickets.filter(t => t.status !== 'resolved'), 'support.write'],
  ] as const;
  for (const [items, permission] of groups) {
    assert.ok(items.every(item => item.ownerId !== 'USR-04'));
    for (const item of items.filter(item => item.ownerId === 'USR-03')) {
      const owner = next.users.find(u => u.id === item.ownerId)!;
      assert.equal(owner.active, true); assert.equal(can(owner.role, permission, next), true);
    }
  }
  assert.equal(next.tasks[0].ownerId, 'USR-04');
  assert.equal(next.tickets[1].ownerId, 'USR-04');
  assert.throws(() => asUser(next, 'USR-04'), /desativado/);
  const reactivated = act(next, 'user.toggle', { id: 'USR-04' });
  assert.equal(reactivated.users.find(u => u.id === 'USR-04')!.active, true);
  assert.deepEqual(reactivated.orders, s.orders);
  assert.deepEqual(reactivated.credits, s.credits);
  assert.deepEqual(reactivated.tasks, next.tasks);
});

test('suspension avoids routing marketing tasks to a finance-only teammate', () => {
  const s = createSeed();
  const next = act(s, 'user.toggle', { id: 'USR-07' });
  for (const task of next.tasks.filter(t => s.tasks.find(previous => previous.id === t.id)?.ownerId === 'USR-07' && t.status !== 'done')) {
    const owner = next.users.find(u => u.id === task.ownerId)!;
    assert.equal(owner.active, true);
    assert.equal(can(owner.role, 'tasks.write', next), true);
    assert.notEqual(owner.id, 'USR-08');
  }
});


test('won CRM stages require an active confirmed payment and seed won opportunities have one', () => {
  let s = createSeed();
  assert.ok(s.opportunities.filter(o => o.stage === 'won').every(opportunity => s.orders.some(o => o.clientId === opportunity.clientId && o.status === 'paid' && o.paidAt)));
  const unpaid = s.opportunities.find(o => o.stage === 'new' && !s.orders.some(order => order.clientId === o.clientId && order.status === 'paid'))!;
  assert.throws(() => act(s, 'crm.move', { id: unpaid.id, stage: 'won' }), /confirme o pagamento/);
  assert.equal(s.opportunities.find(o => o.id === unpaid.id)!.stage, 'new');
  s = act(s, 'crm.create', { clientId: 'CLI-0001', title: 'Venda assistida confirmada', valueCents: 1499, ownerId: 'USR-04' });
  const paidOpportunityId = s.opportunities.at(-1)!.id;
  s = act(s, 'crm.move', { id: paidOpportunityId, stage: 'won' });
  assert.equal(s.opportunities.at(-1)!.stage, 'won');
  const clientOrders = s.orders.filter(o => o.clientId === 'CLI-0001' && o.status === 'paid');
  for (const order of clientOrders) s = act(s, 'order.refund', { id: order.id });
  s = act(s, 'crm.move', { id: paidOpportunityId, stage: 'proposal' });
  assert.throws(() => act(s, 'crm.move', { id: paidOpportunityId, stage: 'won' }), /confirme o pagamento/);
});

test('order.create validates and normalizes old and Mercosur plates with an ID-based demo fallback', () => {
  const s = createSeed();
  const oldPlate = act(s, 'order.create', { clientId: 'CLI-0001', product: 'base', plate: 'abc-1234' }).orders.at(-1)!;
  assert.equal(oldPlate.plate, 'ABC1234');
  assert.equal(oldPlate.totalCents, 1499);
  const mercosur = act(s, 'order.create', { clientId: 'CLI-0001', product: 'complete', plate: ' abc1d23 ' }).orders.at(-1)!;
  assert.equal(mercosur.plate, 'ABC1D23');
  assert.equal(mercosur.totalCents, 7990);
  const fallback = act(s, 'order.create', { clientId: 'CLI-0001', product: 'base' }).orders.at(-1)!;
  assert.equal(fallback.plate, `DEM${fallback.id.split('-').at(-1)}`);
  for (const plate of ['AB1234', 'ABCD123', 'ABC12D3', '1234567', '']) assert.throws(() => act(s, 'order.create', { clientId: 'CLI-0001', product: 'base', plate }), /placa|Placa/);
});


test('support can open an internal ticket for a paid consultation in the organization queue with a matching client', () => {
  let s = createSeed();
  const order = s.orders.find(o => o.status === 'paid' && o.consultationStatus === 'failed' && !s.conversations.some(c => c.clientId === o.clientId && c.ownerId === 'USR-06'))!;
  const payload = { title: 'Bureau sem retorno', description: 'Revisar a falha desta consulta paga.', clientId: order.clientId, orderId: order.id, priority: 'high' };
  const next = act(asUser(s, 'USR-06'), 'ticket.create', payload);
  assert.equal(next.tickets.at(-1)!.clientId, order.clientId);
  assert.equal(next.tickets.at(-1)!.orderId, order.id);
  assert.equal(next.tickets.at(-1)!.ownerId, 'USR-06');
  assert.equal(next.tickets.at(-1)!.status, 'open');
  assert.throws(() => act(asUser(s, 'USR-06'), 'ticket.create', { ...payload, clientId: 'CLI-0001' }), /não correspondem/);
  assert.equal(act(asUser(s, 'USR-06'), 'ticket.create', { ...payload, clientId: undefined }).tickets.at(-1)!.clientId, order.clientId);
  assert.throws(() => act(asUser(s, 'USR-06'), 'crm.create', { clientId: order.clientId, title: 'Venda', valueCents: 1499 }), /permissão/);
  s = act(s, 'permissions.update', { role: 'support', permission: 'consultations.write', enabled: false });
  assert.throws(() => act(asUser(s, 'USR-06'), 'ticket.create', payload), /outro responsável/);
});
