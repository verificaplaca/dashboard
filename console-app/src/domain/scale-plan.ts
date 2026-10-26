import { addDays, dates, defaultScenario, round, taxAt, type Baseline, type Scenario, type Snapshot } from './simulation';

export const SCALE_LEVELS = [2790, 3100, 3400, 3700, 4100, 4500, 5000] as const;
export const SCALE_CALENDAR = ['Agora–31/10/2026', '01–14/11/2026', '15–28/11/2026', '29/11–12/12/2026', '13–26/12/2026', '27/12/2026–09/01/2027', '10–31/01/2027'];
export function scaleScenario(base: Baseline): Scenario { return { ...defaultScenario(base), name: 'Plano gradual até R$ 5.000/dia · base da etapa atual', baselineBudget: SCALE_LEVELS[0], cycles: SCALE_LEVELS.map((budget, i) => ({ id: `C${i + 1}`, days: [6, 14, 14, 14, 14, 14, 22][i], budget, seasonality: 1, learningPct: 0 })) }; }
export const SCALE_RULES = { cacTarget: 9, cacAlert: 9.5, profitGain: .05, profitDrop: .05, minimumDays: 14, extensionDays: 7 };
export type ActualWindow = { start: string; end: string; days: number; revenue: number; orders: number; ads: number; bureau: number; tax: number; profit: number; profitDay: number; adsDay: number; cac: number | null; issues: string[] };
export type StageReference = { level: number; window: ActualWindow; frozenAt: string };
export type ScaleEntry = { id: string; level: number; changedOn: string; recordedAt: string; reference?: StageReference };
export type ScaleState = { version: 1; entries: ScaleEntry[]; lagDays: number; conversionDays: number };
export type ScaleCheck = { label: string; pass: boolean; detail: string };
export type ScaleDecision = { action: 'increase' | 'hold' | 'reduce'; reasons: string[]; checks: ScaleCheck[]; current: ActualWindow | null; cumulative: ActualWindow | null; recent: ActualWindow | null; previous: StageReference | null; elapsed: number; required: number; eligibleEnd: string; nextLevel: number | null; reduceLevel: number | null; recovered: number | null; valid: boolean };
export function initialScaleState(): ScaleState { return { version: 1, entries: [{ id: 'initial-2026-10-16', level: 0, changedOn: '2026-10-16', recordedAt: '2026-10-23T00:00:00-03:00' }], lagDays: 0, conversionDays: 1 }; }
const isDay = (s: unknown): s is string => { try { dates(String(s), String(s)); return typeof s === 'string'; } catch { return false; } };
const isStamp = (s: unknown): s is string => typeof s === 'string' && Number.isFinite(Date.parse(s));
export function requiredScaleDays(conversionDays: number) { return Math.max(14, Math.ceil(2 * conversionDays / 7) * 7); }
export function validateScaleState(value: unknown): string[] {
  const errors: string[] = [], s = value as ScaleState;
  if (!s || s.version !== 1 || !Array.isArray(s.entries) || !s.entries.length || s.entries.length > 100) return ['Histórico de etapas inválido.'];
  if (!Number.isInteger(s.lagDays) || s.lagDays < 0 || s.lagDays > 30 || !Number.isInteger(s.conversionDays) || s.conversionDays < 1 || s.conversionDays > 60) errors.push('Use reserva de 0 a 30 dias e ciclo de 1 a 60 dias inteiros.');
  s.entries.forEach((e, i) => {
    if (!e || typeof e.id !== 'string' || !Number.isInteger(e.level) || e.level < 0 || e.level >= SCALE_LEVELS.length || !isDay(e.changedOn) || !isStamp(e.recordedAt)) { errors.push('Registro de etapa inválido.'); return; }
    if (i > 0 && (!s.entries[i - 1] || e.changedOn <= s.entries[i - 1].changedOn)) errors.push('As etapas precisam de datas crescentes, sem duas alterações no mesmo dia.');
    if (e.reference) {
      const r = e.reference, w = r.window;
      if (!Number.isInteger(r.level) || r.level < 0 || r.level > e.level || !isStamp(r.frozenAt) || !w || !isDay(w.start) || !isDay(w.end) || w.end >= e.changedOn || w.end < w.start || Date.parse(w.end) - Date.parse(w.start) > 1095 * 86400000 || !Array.isArray(w.issues) || w.issues.length) { errors.push('Referência anterior inválida.'); return; }
      const numeric = [w.days, w.revenue, w.orders, w.ads, w.bureau, w.tax, w.profit, w.profitDay, w.adsDay, w.cac];
      if (numeric.some(n => typeof n !== 'number' || !Number.isFinite(n)) || w.days < 14 || w.days % 7 !== 0 || w.days !== dates(w.start, w.end).length || w.orders <= 0 || w.ads <= 0 || w.revenue < 0 || w.bureau < 0 || w.tax < 0 || w.profit <= 0 || Math.abs(w.profit - round(w.revenue - w.ads - w.bureau - w.tax)) > .011 || Math.abs(w.profitDay - w.profit / w.days) > .00001 || Math.abs(w.adsDay - w.ads / w.days) > .00001 || Math.abs(w.cac! - w.ads / w.orders) > .00001) errors.push('Valores da referência anterior inconsistentes.');
    }
  });
  if (new Set(s.entries.map(e => e?.id)).size !== s.entries.length) errors.push('Há registros de etapa duplicados.');
  return [...new Set(errors)];
}
/** Actual financial results. Explicit zero-sales days remain in the profit denominator. */
export function actualScaleWindow(s: Snapshot, start: string, end: string): ActualWindow | null {
  if (start > end) return null;
  const calendar = dates(start, end), inside = (r: { date: string }) => r.date >= start && r.date <= end;
  const revenue = s.revenue.filter(inside), bureau = s.bureau.filter(inside), ads = s.ads.filter(inside), issues: string[] = [];
  if (start < s.start || end > s.end) issues.push('Leitura fora da cobertura disponível.');
  for (const [name, rows, key] of [['receita', revenue, (r: { date: string }) => r.date], ['bureau', bureau, (r: { date: string }) => r.date], ['Ads', ads, (r: { date: string; campaign_id?: string }) => `${r.date}:${r.campaign_id}`]] as const) {
    if (new Set(rows.map(key)).size !== rows.length) issues.push(`Duplicação na fonte de ${name}.`);
    if (calendar.some(d => !rows.some(r => r.date === d))) issues.push(`Faltam dias de ${name}; ausência não equivale a zero.`);
  }
  if (revenue.some(r => !Number.isFinite(r.revenue) || r.revenue < 0 || !Number.isInteger(r.paid_orders) || r.paid_orders < 0) || bureau.some(r => !Number.isFinite(r.custo_bureau) || r.custo_bureau < 0) || ads.some(r => !Number.isFinite(r.cost_micros) || r.cost_micros < 0)) issues.push('Valores financeiros inválidos nas fontes.');
  if (calendar.some(d => taxAt(s.targets, d).fallback)) issues.push('Imposto sem cadastro: fallback de 8% requer conferência.');
  const sum = <T,>(rows: T[], fn: (r: T) => number) => rows.reduce((total, r) => total + fn(r), 0);
  const r = round(sum(revenue, r => r.revenue)), a = round(sum(ads, r => r.cost_micros / 1e6)), b = round(sum(bureau, r => r.custo_bureau)), o = sum(revenue, r => r.paid_orders), t = round(sum(revenue, r => r.revenue * taxAt(s.targets, r.date).value / 100)), p = round(r - a - b - t);
  return { start, end, days: calendar.length, revenue: r, orders: o, ads: a, bureau: b, tax: t, profit: p, profitDay: p / calendar.length, adsDay: a / calendar.length, cac: o > 0 ? a / o : null, issues };
}
export function evaluateScalePlan(s: Snapshot | null, state: ScaleState, today: string, requireFresh = true): ScaleDecision {
  const entry = state.entries?.at(-1), required = requiredScaleDays(state.conversionDays), eligibleEnd = addDays(today, -1 - state.lagDays);
  const out: ScaleDecision = { action: 'hold', reasons: [], checks: [], current: null, cumulative: null, recent: null, previous: entry?.reference || null, elapsed: 0, required, eligibleEnd, nextLevel: entry && entry.level < 6 ? entry.level + 1 : null, reduceLevel: null, recovered: null, valid: false };
  const invalid = validateScaleState(state);
  if (invalid.length || !entry || entry.changedOn > today) { out.reasons = [...invalid, ...(entry && entry.changedOn > today ? ['A etapa registrada ainda não foi aplicada: data futura.'] : [])]; return out; }
  if (!s) { out.reasons = ['Aguardando leitura atualizada das fontes reais.']; return out; }
  if (!isStamp(s.loadedAt) || s.end !== addDays(today, -1) || requireFresh && new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(s.loadedAt)) !== today) { out.reasons = ['Leitura desatualizada. Atualize os dados antes de decidir.']; return out; }
  const start = addDays(entry.changedOn, 1);
  out.elapsed = start <= eligibleEnd ? dates(start, eligibleEnd).length : 0;
  out.current = actualScaleWindow(s, [start, addDays(eligibleEnd, 1 - required)].sort().at(-1)!, eligibleEnd);
  out.cumulative = actualScaleWindow(s, entry.changedOn, eligibleEnd);
  out.recent = actualScaleWindow(s, [start, addDays(eligibleEnd, -1)].sort().at(-1)!, eligibleEnd);
  const last3 = (out.cumulative?.days || 0) >= 3 ? actualScaleWindow(s, addDays(eligibleEnd, -2), eligibleEnd) : null;
  const last7 = out.elapsed >= 7 ? actualScaleWindow(s, addDays(eligibleEnd, -6), eligibleEnd) : null;
  const quality = [...(out.current?.issues || []), ...(out.cumulative?.issues || []), ...s.warnings];
  if (quality.length || !out.current) { out.reasons = quality.length ? [...new Set(quality)] : ['Ainda não há dia completo após a alteração, respeitando a reserva de atualização.']; return out; }
  out.valid = true;
  const w = out.current, ref = entry.reference, prior = ref?.window, higher = !!ref && entry.level > ref.level;
  const lower = state.entries.flatMap(e => e.reference ? [e.reference.level] : []).filter(level => level < entry.level);
  out.reduceLevel = lower.length ? Math.max(...lower) : null;
  out.recovered = higher && prior && out.cumulative ? round(out.cumulative.profit - prior.profitDay * out.cumulative.days) : null;
  const minimum = out.elapsed >= required;
  const comparable = !prior || prior.days === required;
  const profitPass = prior ? w.profitDay >= prior.profitDay * (higher ? 1.05 : .95) - .000001 : w.profitDay > 0;
  const spendPass = !higher || w.adsDay > prior!.adsDay + .005;
  const budgetPass = w.adsDay >= SCALE_LEVELS[entry.level] * .8 && w.adsDay <= SCALE_LEVELS[entry.level] * 1.2;
  out.checks = [
    { label: 'Janela completa e maturação', pass: minimum, detail: `${out.elapsed}/${required} dias completos; mínimo de dois ciclos de conversão.` },
    { label: 'CAC ponderado até R$ 9,00', pass: w.cac !== null && w.cac <= 9, detail: w.cac === null ? 'Sem pedidos pagos para calcular o CAC.' : `CAC observado: R$ ${w.cac.toFixed(2).replace('.', ',')}. Soma de Ads ÷ soma de pedidos pagos.` },
    { label: prior ? higher ? 'Lucro/dia pelo menos 5% maior' : 'Revalidar patamar de retorno' : 'Base inicial com lucro positivo', pass: profitPass && comparable && (entry.level === 0 || !!prior), detail: !prior && entry.level > 0 ? 'Sem referência validada anterior. Registre a sequência real de etapas.' : prior ? `Referência de ${prior.days} dias congelada em ${ref!.frozenAt.slice(0, 10)}; janelas de mesma duração.` : 'Primeiro aumento é um teste; ainda não há patamar anterior para exigir +5%.' },
    { label: 'Gasto efetivo compatível com a etapa', pass: spendPass && budgetPass, detail: higher ? 'Ads/dia precisa crescer; média dentro de ±20% da verba registrada.' : 'Média de Ads dentro de ±20% da verba registrada; confira mudanças não registradas.' },
    { label: 'Transição recuperada no acumulado', pass: out.recovered === null ? (out.cumulative?.profit || 0) > 0 : out.recovered >= -.005, detail: 'O dia da alteração e as perdas iniciais entram no resultado acumulado.' }
  ];
  const red: string[] = [];
  if (last3 && !last3.issues.length && last3.profit < 0) red.push('Prejuízo líquido acumulado nos últimos três dias completos. Conter a escala e investigar.');
  if (higher && comparable && prior && minimum && w.profitDay < prior.profitDay * .95 - .000001) red.push('Lucro/dia mais de 5% abaixo do patamar anterior validado.');
  if (higher && comparable && prior && out.elapsed >= required + 7 && w.profitDay <= prior.profitDay + .000001 && w.adsDay > prior.adsDay + .005) red.push('Após a extensão de sete dias, o gasto cresceu sem ganho de lucro/dia.');
  if (higher && prior && last7 && !last7.issues.length && (last7.cac === null || last7.cac > 9.5) && last7.profitDay < prior.profitDay) red.push('Nos últimos sete dias, CAC acima de R$ 9,50 e lucro/dia inferior à referência.');
  if (red.length) { out.action = 'reduce'; out.reasons = red; return out; }
  out.reasons = out.checks.filter(c => !c.pass).map(c => c.label + ': ' + c.detail);
  if (!out.reasons.length) {
    if (entry.level === 6) out.reasons = ['Teto de R$ 5.000/dia atingido. Manter e acompanhar o lucro líquido.'];
    else { out.action = 'increase'; out.reasons = ['Critérios atendidos. Testar o próximo patamar e reiniciar a validação após registrar a alteração.']; }
  } else if (minimum) out.reasons.push(`Sem novo aumento: reavaliar após sete dias completos adicionais. CAC alto isolado não manda reduzir.`);
  return out;
}
/** Record an action already performed in the ad account; never mutates the account. */
export function recordScaleStage(state: ScaleState, s: Snapshot, level: number, changedOn: string, today: string): ScaleState {
  if (!isDay(changedOn) || changedOn > today || changedOn <= state.entries.at(-1)!.changedOn || !Number.isInteger(level) || level < 0 || level > 6) throw new Error('Escolha um patamar e uma data já ocorrida, posterior ao último registro.');
  const previous = state.entries.at(-1)!;
  const evaluation = evaluateScalePlan({ ...s, end: addDays(changedOn, -1) }, state, changedOn, false);
  let reference: StageReference | undefined;
  if (level === previous.level + 1 && evaluation.action === 'increase' && evaluation.current) reference = { level: previous.level, window: structuredClone(evaluation.current), frozenAt: new Date().toISOString() };
  else if (level <= previous.level) reference = [...state.entries].reverse().flatMap(e => e.reference ? [e.reference] : []).find(r => r.level === level && r.window.days === requiredScaleDays(state.conversionDays));
  const next: ScaleState = { ...state, entries: [...state.entries, { id: crypto.randomUUID(), level, changedOn, recordedAt: new Date().toISOString(), ...(reference ? { reference } : {}) }] };
  const errors = validateScaleState(next); if (errors.length) throw new Error(errors.join(' '));
  return next;
}
