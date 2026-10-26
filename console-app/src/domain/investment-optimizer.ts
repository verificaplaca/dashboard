import { addDays, buildBaseline, dates, project, round, taxAt, type Baseline, type Projection, type Scenario, type Snapshot, type TaxRow } from './simulation';
export type OptimizationSettings = { start: string; days: number; controlBudget: number; elasticity: number; responsivePct: number; taxOverride: number | null; extraFeePct: number; fixedDay: number; minBudget: number; maxBudget: number; learningDays: number; learningPct: number };
export type BudgetPoint = { budget: number; profit: number };
export type Optimum = { budget: number; projection: Projection; profit: number; delta: number; status: 'interior' | 'keep' | 'upper-bound' | 'lower-bound'; scenario: Scenario };
export type OptimizationWindow = 7 | 14 | 30 | 'post-change' | 'recent';
export type Intervention = { changeDate: string; recentDays: number };
export type WindowOptimization = { window: OptimizationWindow; label: string; start: string; end: string; base: Baseline; steady?: Optimum; learning?: Optimum; changedPeak?: Optimum; control?: BudgetPoint; curve: { budget: number; steady: number; learning: number }[] };
export function optimizationSettings(s: Scenario): OptimizationSettings { return { start: s.start, days: s.cycles.reduce((a, c) => a + c.days, 0), controlBudget: s.baselineBudget, elasticity: s.elasticity, responsivePct: s.responsivePct, taxOverride: s.taxOverride, extraFeePct: s.extraFeePct, fixedDay: s.fixedDay, minBudget: 0, maxBudget: 10000, learningDays: 7, learningPct: 6 }; }
export function validateOptimization(p: OptimizationSettings): string[] {
  const errors: string[] = [];
  const valid = (n: number, min: number, max: number) => Number.isFinite(n) && n >= min && n <= max;
  try { dates(p.start, p.start); } catch { errors.push('Data inicial inválida.'); }
  if (!Number.isInteger(p.days) || !valid(p.days, 1, 1096)) errors.push('O horizonte deve ter entre 1 e 1.096 dias inteiros.');
  if (!valid(p.minBudget, 0, 1e7) || !valid(p.maxBudget, 0, 1e7) || p.maxBudget <= p.minBudget) errors.push('A faixa precisa de limite mínimo ≥ 0 e máximo maior que o mínimo, até R$ 10 milhões/dia.');
  if (!valid(p.controlBudget, .01, 1e7) || !valid(p.elasticity, 0, 2) || !valid(p.responsivePct, 0, 100)) errors.push('Confira verba de controle, elasticidade e parcela sensível à mídia.');
  if (!Number.isInteger(p.learningDays) || !valid(p.learningDays, 0, 366) || !valid(p.learningPct, 0, 100)) errors.push('Aprendizado: de 0 a 366 dias inteiros e perda de 0% a 100%.');
  if ((p.taxOverride !== null && !valid(p.taxOverride, 0, 100)) || !valid(p.extraFeePct, 0, 100) || !valid(p.fixedDay, 0, 1e7)) errors.push('Confira imposto, taxas extras e custos fixos.');
  return errors;
}
export function constantBudgetScenario(base: Baseline, p: OptimizationSettings, budget: number, learning: boolean, forceChange = false): Scenario {
  const learningDays = Math.min(p.days, p.learningDays);
  const learns = learning && learningDays > 0 && p.learningPct > 0 && (forceChange || Math.abs(budget - p.controlBudget) > .005);
  const cycles = learns ? [{ id: 'C1', days: learningDays, budget, seasonality: 1, learningPct: p.learningPct }, ...(p.days > learningDays ? [{ id: 'C2', days: p.days - learningDays, budget, seasonality: 1, learningPct: 0 }] : [])] : [{ id: 'C1', days: p.days, budget, seasonality: 1, learningPct: 0 }];
  return { name: `Ótimo projetado · R$ ${budget.toLocaleString('pt-BR')}/dia${learning ? ' · com aprendizado' : ' · estabilizado'}`, start: p.start, baselineBudget: p.controlBudget, elasticity: p.elasticity, responsivePct: p.responsivePct, ticket: base.ticket, bureauUnit: base.bureauUnit, conversionPct: 0, taxOverride: p.taxOverride, extraFeePct: p.extraFeePct, fixedDay: p.fixedDay, cycles };
}
/** Global comparison of bounded constant budgets, plus the isolated option of keeping the control. All monetary values use the existing engine. */
export function optimizeWindow(base: Baseline, targets: TaxRow[], p: OptimizationSettings) {
  const errors = [...validateOptimization(p), ...base.issues]; if (errors.length) throw new Error(errors.join(' '));
  const keepScenario = constantBudgetScenario(base, p, p.controlBudget, false);
  const controlProjection = project(base, targets, keepScenario).control;
  const cache = new Map<string, Projection>();
  const evaluate = (budget: number, learning: boolean, forceChange = false): Optimum => {
    const scenario = constantBudgetScenario(base, p, budget, learning, forceChange);
    const key = `${budget}:${scenario.cycles[0].learningPct}`;
    let projection = cache.get(key);
    if (!projection) { projection = project(base, targets, scenario).scenario; cache.set(key, projection); }
    const same = Math.abs(budget - p.controlBudget) <= .005 && !forceChange;
    return { budget, projection, profit: projection.profit, delta: round(projection.profit - controlProjection.profit), status: same ? 'keep' : budget >= p.maxBudget - .005 ? 'upper-bound' : budget <= p.minBudget + .005 ? 'lower-bound' : 'interior', scenario };
  };
  // The smooth objective is A0 + A1 × (budget/baseAdsDay)^elasticity − days × budget.
  // Its only interior maximum for positive A1 has 0 < elasticity < 1. Other cases are bounded by endpoints.
  const horizon = dates(p.start, addDays(p.start, p.days - 1));
  const stationary = (learning: boolean) => {
    const coefficient = horizon.reduce((sum, day, i) => sum + base.ordersDay * (base.ticket * (1 - (p.taxOverride ?? taxAt(targets, day).value) / 100 - p.extraFeePct / 100) - base.bureauUnit) * (learning && i < p.learningDays ? 1 - p.learningPct / 100 : 1), 0) * p.responsivePct / 100;
    if (!(p.elasticity > 0 && p.elasticity < 1 && coefficient > 0)) return null;
    const value = base.adsDay * Math.pow(p.elasticity * coefficient / (p.days * base.adsDay), 1 / (1 - p.elasticity));
    return Number.isFinite(value) ? value : null;
  };
  const pick = (a: Optimum, b: Optimum) => b.profit > a.profit + .005 || (Math.abs(b.profit - a.profit) < .005 && (b.status === 'keep' || (a.status !== 'keep' && b.budget < a.budget))) ? b : a;
  const find = (learning: boolean, includeKeep: boolean) => {
    const candidates = new Set(Array.from({ length: 81 }, (_, i) => round(p.minBudget + (p.maxBudget - p.minBudget) * i / 80)));
    const point = stationary(learning);
    if (point !== null && point >= p.minBudget && point <= p.maxBudget) {
      candidates.add(round(point));
      // Check nearby whole-real budgets because daily money recognition can create cent-level plateaus.
      for (let offset = -50; offset <= 50; offset++) { const n = Math.round(point) + offset; if (n >= p.minBudget && n <= p.maxBudget) candidates.add(n); }
    }
    let best = [...candidates].map(n => evaluate(n, learning, !includeKeep)).reduce(pick);
    if (includeKeep && p.controlBudget >= p.minBudget && p.controlBudget <= p.maxBudget) best = pick(best, evaluate(p.controlBudget, false));
    return best;
  };
  const steady = find(false, true), changedPeak = find(true, false), learning = find(true, true);
  const budgets = new Set(Array.from({ length: 81 }, (_, i) => round(p.minBudget + (p.maxBudget - p.minBudget) * i / 80)));
  for (const b of [steady.budget, changedPeak.budget, learning.budget, p.controlBudget]) if (b >= p.minBudget && b <= p.maxBudget) budgets.add(b);
  const curve = [...budgets].sort((a, b) => a - b).map(budget => ({ budget, steady: evaluate(budget, false, true).profit, learning: evaluate(budget, true, true).profit }));
  return { steady, learning, changedPeak, control: { budget: p.controlBudget, profit: controlProjection.profit }, curve };
}
export function historicalBase(snapshot: Snapshot, start: string, end = snapshot.end): Baseline {
  if (start > end) return { days: 0, revenue: 0, orders: 0, ads: 0, bureau: 0, tax: 0, profit: 0, ticket: 0, bureauUnit: 0, adsDay: 0, ordersDay: 0, taxPct: 0, issues: ['Ainda não há dia completo após as alterações nesta leitura.'] };
  const base = buildBaseline({ ...snapshot, start, end });
  if (start < snapshot.start || end > snapshot.end) base.issues.push('A leitura disponível não cobre toda esta janela.');
  return base;
}
export function interventionBases(snapshot: Snapshot, intervention: Intervention) {
  dates(intervention.changeDate, intervention.changeDate);
  if (!Number.isInteger(intervention.recentDays) || intervention.recentDays < 1 || intervention.recentDays > 30) throw new Error('Use entre 1 e 30 dias recentes.');
  const postStart = addDays(intervention.changeDate, 1), recentStart = [postStart, addDays(snapshot.end, 1 - intervention.recentDays)].sort().at(-1)!;
  return {
    before: historicalBase(snapshot, addDays(intervention.changeDate, -7), addDays(intervention.changeDate, -1)),
    transition: historicalBase(snapshot, intervention.changeDate, intervention.changeDate),
    post: historicalBase(snapshot, postStart), recent: historicalBase(snapshot, recentStart), postStart, recentStart,
  };
}
export function compareOptimizationWindows(snapshot: Snapshot, p: OptimizationSettings, intervention?: Intervention): WindowOptimization[] {
  const contexts: { window: OptimizationWindow; label: string; start: string }[] = ([7, 14, 30] as const).map(window => ({ window, label: `${window} dias`, start: addDays(snapshot.end, 1 - window) }));
  if (intervention) {
    const b = interventionBases(snapshot, intervention);
    contexts.unshift({ window: 'post-change', label: 'Após alterações', start: b.postStart }, { window: 'recent', label: 'Recuperação recente', start: b.recentStart });
  }
  return contexts.map(({ window, label, start }) => {
    const base = historicalBase(snapshot, start);
    const context = { window, label, start, end: snapshot.end, base, curve: [] };
    return base.issues.length ? context : { ...context, ...optimizeWindow(base, snapshot.targets, p) };
  });
}
/** Reconcile every additional real of revenue/cost to the same engine and control. */
export function investmentGrowth(base: Baseline, targets: TaxRow[], p: OptimizationSettings, budget: number, learning: boolean) {
  if (!Number.isFinite(budget) || budget < 0 || budget > 1e7) throw new Error('Investimento deve estar entre zero e R$ 10 milhões/dia.');
  const result = project(base, targets, constantBudgetScenario(base, p, budget, learning));
  const steady = project(base, targets, constantBudgetScenario(base, p, budget, false));
  const delta = { revenue: round(result.scenario.revenue - result.control.revenue), ads: result.extraAds, bureau: round(result.scenario.bureau - result.control.bureau), tax: round(result.scenario.tax - result.control.tax), fees: round(result.scenario.fees - result.control.fees), fixed: round(result.scenario.fixed - result.control.fixed), profit: result.delta };
  return { budget, ...result, delta, learningCost: round(steady.scenario.profit - result.scenario.profit), incrementalReturn: result.extraAds > 0 ? result.delta / result.extraAds : null, profitGrowthPct: result.control.profit > 0 ? result.delta / result.control.profit * 100 : null };
}
