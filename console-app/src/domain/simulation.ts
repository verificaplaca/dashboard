/** Model v1. BRL throughout; round money at daily recognition, orders remain expected values. */
export type RevenueRow = { date: string; revenue: number; paid_orders: number };
export type BureauRow = { date: string; custo_bureau: number };
export type AdsRow = { date: string; campaign_id: string; cost_micros: number };
export type TaxRow = { month: string; tax_pct: number | null };
export type Snapshot = { start: string; end: string; loadedAt: string; revenue: RevenueRow[]; bureau: BureauRow[]; ads: AdsRow[]; targets: TaxRow[]; warnings: string[] };
export type Baseline = { days: number; revenue: number; orders: number; ads: number; bureau: number; tax: number; profit: number; ticket: number; bureauUnit: number; adsDay: number; ordersDay: number; taxPct: number; issues: string[] };
export type Cycle = { id: string; days: number; budget: number; seasonality: number; learningPct: number };
export type Scenario = { name: string; start: string; baselineBudget: number; elasticity: number; responsivePct: number; ticket: number; bureauUnit: number; conversionPct: number; taxOverride: number | null; extraFeePct: number; fixedDay: number; cycles: Cycle[] };
export type Projection = { days: number; orders: number; revenue: number; ads: number; bureau: number; tax: number; panelProfit: number; fees: number; fixed: number; profit: number; margin: number; cpa: number | null; roas: number | null };
export type CycleResult = { id: string; start: string; end: string; control: Projection; scenario: Projection };
export const round = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;
export function dates(start: string, end: string): string[] {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) throw new Error('Informe datas válidas.');
  const a = Date.parse(start + 'T00:00:00Z'), b = Date.parse(end + 'T00:00:00Z');
  if (!Number.isFinite(a) || !Number.isFinite(b) || new Date(a).toISOString().slice(0, 10) !== start || new Date(b).toISOString().slice(0, 10) !== end || b < a || (b - a) / 86400000 > 1095) throw new Error('O período deve ter entre 1 e 1.096 dias.');
  return Array.from({ length: (b - a) / 86400000 + 1 }, (_, i) => new Date(a + i * 86400000).toISOString().slice(0, 10));
}
export function addDays(day: string, n: number) { return new Date(Date.parse(day + 'T00:00:00Z') + n * 86400000).toISOString().slice(0, 10); }
export function saoPauloToday() { const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date()); return ['year', 'month', 'day'].map(k => parts.find(p => p.type === k)!.value).join('-'); }
export function taxAt(targets: TaxRow[], day: string): { value: number; fallback: boolean } {
  const row = targets.filter(t => t.month.slice(0, 7) <= day.slice(0, 7) && t.tax_pct !== null).sort((a, b) => b.month.localeCompare(a.month))[0];
  return row ? { value: Number(row.tax_pct), fallback: false } : { value: 8, fallback: true };
}
export function buildBaseline(s: Snapshot): Baseline {
  const calendar = dates(s.start, s.end), inRange = (r: { date: string }) => r.date >= s.start && r.date <= s.end;
  const rev = s.revenue.filter(inRange), bur = s.bureau.filter(inRange), ads = s.ads.filter(inRange), issues: string[] = [];
  for (const [label, rows, key] of [['receita', rev, (r: { date: string }) => r.date], ['bureau', bur, (r: { date: string }) => r.date], ['Ads', ads, (r: { date: string; campaign_id?: string }) => `${r.date}:${r.campaign_id}`]] as const) {
    if (new Set(rows.map(key)).size !== rows.length) issues.push(`Registros duplicados na fonte de ${label}.`);
  }
  if (calendar.some(d => !rev.some(r => r.date === d))) issues.push('A fonte de receita não cobre todos os dias selecionados.');
  if (rev.some(r => (r.revenue > 0 || r.paid_orders > 0) && !bur.some(b => b.date === r.date))) issues.push('Há dias de vendas sem custo de bureau informado.');
  if (!ads.length) issues.push('Não há investimento Ads informado neste período.');
  else if (rev.some(r => (r.revenue > 0 || r.paid_orders > 0) && !ads.some(a => a.date === r.date))) issues.push('Há dias de vendas sem registros de Ads; ausência não foi tratada como gasto zero.');
  const sum = <T,>(rows: T[], f: (r: T) => number) => rows.reduce((a, r) => a + Number(f(r)), 0);
  const revenue = round(sum(rev, r => r.revenue)), orders = sum(rev, r => r.paid_orders), bureau = round(sum(bur, r => r.custo_bureau)), adCost = round(sum(ads, r => r.cost_micros / 1e6));
  const tax = round(sum(rev, r => r.revenue * taxAt(s.targets, r.date).value / 100));
  if (!(revenue > 0 && orders > 0 && adCost > 0)) issues.push('A base precisa ter receita, pedidos e investimento positivos.');
  return { days: calendar.length, revenue, orders, bureau, ads: adCost, tax, profit: round(revenue - bureau - adCost - tax), ticket: revenue / orders, bureauUnit: bureau / orders, adsDay: adCost / calendar.length, ordersDay: orders / calendar.length, taxPct: revenue > 0 ? tax / revenue * 100 : 0, issues };
}
export function defaultScenario(b: Baseline): Scenario {
  return { name: 'Minha hipótese', start: '2026-10-26', baselineBudget: round(b.adsDay), elasticity: .8, responsivePct: 100, ticket: b.ticket, bureauUnit: b.bureauUnit, conversionPct: 0, taxOverride: null, extraFeePct: 0, fixedDay: 0, cycles: [2790, 3100, 3450, 3750, 3750, 4100, 4100].map((budget, i) => ({ id: `C${i + 1}`, days: 14, budget, seasonality: 1, learningPct: 0 })) };
}
export function validateScenario(s: Scenario) {
  const errors: string[] = [];
  const between = (n: number, lo: number, hi: number) => Number.isFinite(n) && n >= lo && n <= hi;
  if (!s.name.trim()) errors.push('Dê um nome ao cenário.');
  try { dates(s.start, s.start); } catch { errors.push('Data inicial inválida.'); }
  if (!between(s.baselineBudget, .01, 1e7) || !between(s.ticket, .01, 1e6) || !between(s.bureauUnit, 0, 1e6)) errors.push('Investimento de controle e ticket devem ser positivos; bureau não pode ser negativo.');
  if (!between(s.elasticity, 0, 2) || !between(s.responsivePct, 0, 100) || !between(s.conversionPct, -99, 500)) errors.push('Elasticidade: 0 a 2; parcela sensível à mídia: 0 a 100%; conversão: −99% a 500%.');
  if ((s.taxOverride !== null && !between(s.taxOverride, 0, 100)) || !between(s.extraFeePct, 0, 100) || !between(s.fixedDay, 0, 1e7)) errors.push('Confira imposto, taxas e custos fixos.');
  if (!s.cycles.length || s.cycles.length > 24 || s.cycles.some(c => !Number.isInteger(c.days) || !between(c.days, 1, 366) || !between(c.budget, 0, 1e7) || !between(c.seasonality, .01, 3) || !between(c.learningPct, 0, 100)) || s.cycles.reduce((a, c) => a + c.days, 0) > 1096) errors.push('Use até 24 ciclos, dias inteiros, verba ≥ 0, sazonalidade entre 0,01 e 3 e ajuste de 0 a 100%. Máximo de 1.096 dias.');
  return errors;
}
function empty(): Projection { return { days: 0, orders: 0, revenue: 0, ads: 0, bureau: 0, tax: 0, panelProfit: 0, fees: 0, fixed: 0, profit: 0, margin: 0, cpa: null, roas: null }; }
export function total(rows: Projection[]): Projection {
  const out = empty();
  for (const r of rows) for (const k of ['days', 'orders', 'revenue', 'ads', 'bureau', 'tax', 'fees', 'fixed'] as const) out[k] += r[k];
  for (const k of ['revenue', 'ads', 'bureau', 'tax', 'fees', 'fixed'] as const) out[k] = round(out[k]);
  out.panelProfit = round(out.revenue - out.ads - out.bureau - out.tax); out.profit = round(out.panelProfit - out.fees - out.fixed);
  out.margin = out.revenue ? out.profit / out.revenue * 100 : 0; out.cpa = out.orders ? out.ads / out.orders : null; out.roas = out.ads + out.bureau ? out.revenue / (out.ads + out.bureau) : null;
  return out;
}
export function project(b: Baseline, targets: TaxRow[], s: Scenario): { cycles: CycleResult[]; control: Projection; scenario: Projection; delta: number; extraAds: number; extraOrders: number; marginalCpa: number | null; breakEvenCpa: number; extraRevenueNeeded: number | null } {
  const errors = validateScenario(s); if (errors.length || b.issues.length) throw new Error([...errors, ...b.issues].join(' '));
  let offset = 0;
  const controlRatio = s.baselineBudget / b.adsDay;
  const response = (ratio: number) => (1 - s.responsivePct / 100) + s.responsivePct / 100 * Math.pow(ratio, s.elasticity);
  const cycles = s.cycles.map(c => {
    const start = addDays(s.start, offset), end = addDays(start, c.days - 1); offset += c.days;
    const make = (isControl: boolean) => total(dates(start, end).map(day => {
      const p = empty(); p.days = 1;
      p.orders = b.ordersDay * response(isControl ? controlRatio : c.budget / b.adsDay) * c.seasonality * (isControl ? 1 : (1 - c.learningPct / 100) * (1 + s.conversionPct / 100));
      p.revenue = round(p.orders * (isControl ? b.ticket : s.ticket)); p.bureau = round(p.orders * (isControl ? b.bureauUnit : s.bureauUnit));
      p.ads = isControl ? s.baselineBudget : c.budget; p.tax = round(p.revenue * (s.taxOverride ?? taxAt(targets, day).value) / 100);
      p.fees = round(p.revenue * s.extraFeePct / 100); p.fixed = s.fixedDay; return p;
    }));
    return { id: c.id, start, end, control: make(true), scenario: make(false) };
  });
  const control = total(cycles.map(c => c.control)), scenario = total(cycles.map(c => c.scenario));
  const delta = round(scenario.profit - control.profit), extraAds = round(scenario.ads - control.ads), extraOrders = scenario.orders - control.orders;
  const taxRate = scenario.revenue ? scenario.tax / scenario.revenue : (s.taxOverride ?? taxAt(targets, s.start).value) / 100;
  const unit = s.ticket * (1 - taxRate - s.extraFeePct / 100) - s.bureauUnit;
  const contributionRate = unit / s.ticket;
  return { cycles, control, scenario, delta, extraAds, extraOrders, marginalCpa: extraAds > 0 && extraOrders > 0 ? extraAds / extraOrders : null, breakEvenCpa: unit, extraRevenueNeeded: contributionRate > 0 ? Math.max(0, -delta / contributionRate) : null };
}
