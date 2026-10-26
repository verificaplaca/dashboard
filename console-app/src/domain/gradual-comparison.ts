import { addDays, dates, project, round, taxAt, type Cycle, type Snapshot } from './simulation';
import { historicalBase } from './investment-optimizer';
import { scaleScenario } from './scale-plan';
export type GradualLearning = { days: number; lossPct: number };
export const DEFAULT_GRADUAL_LEARNING: GradualLearning = { days: 7, lossPct: 6 };
export function compareGradualPlan(snapshot: Snapshot, learning: GradualLearning = DEFAULT_GRADUAL_LEARNING) {
  if (!Number.isInteger(learning.days) || learning.days < 0 || learning.days > 30 || !Number.isFinite(learning.lossPct) || learning.lossPct < 0 || learning.lossPct > 100) throw new Error('Aprendizado: use 0 a 30 dias inteiros e perda entre 0% e 100%.');
  const start = addDays(snapshot.end, -1), end = snapshot.end;
  if (start <= '2026-10-16') throw new Error('A comparação precisa de dois dias completos após as alterações de 16/10.');
  const base = historicalBase(snapshot, start, end);
  if (base.issues.length || snapshot.warnings.length) throw new Error([...base.issues, ...snapshot.warnings].join(' '));
  if (dates(start, end).some(d => taxAt(snapshot.targets, d).fallback)) throw new Error('Imposto sem cadastro na recuperação recente. Confira as metas antes de projetar.');
  const nominal = scaleScenario(base);
  // Each change gets a single transition. Keeping the initial budget never gets another penalty.
  const transitionCycles = nominal.cycles.flatMap((c, i): Cycle[] => {
    const changed = i > 0 && c.budget !== nominal.cycles[i - 1].budget;
    if (!changed || !learning.days || !learning.lossPct) return [{ ...c }];
    const days = Math.min(c.days, learning.days);
    return [{ ...c, id: `${c.id}a`, days, learningPct: learning.lossPct }, ...(c.days > days ? [{ ...c, id: `${c.id}b`, days: c.days - days }] : [])];
  });
  const affectedDays = transitionCycles.filter(c => c.learningPct > 0).reduce((sum, c) => sum + c.days, 0);
  const directCycle: Cycle = { id: 'direct', days: 98, budget: 5000, seasonality: 1, learningPct: 0 };
  const directAffectedDays = learning.days && learning.lossPct ? Math.min(98, learning.days) : 0;
  const directCycles: Cycle[] = directAffectedDays ? [
    { ...directCycle, id: 'direct-learning', days: directAffectedDays, learningPct: learning.lossPct },
    { ...directCycle, days: 98 - directAffectedDays },
  ] : [directCycle];
  const rows = ([1, .8] as const).flatMap(elasticity => {
    const steady = project(base, snapshot.targets, { ...nominal, elasticity });
    const directSteady = project(base, snapshot.targets, { ...nominal, elasticity, cycles: [directCycle] });
    return [false, true].map(withLearning => {
      const label = `${elasticity === 1 ? 'CAC constante' : 'Elasticidade 0,8'} · ${withLearning ? 'com' : 'sem'} aprendizado`;
      const scenario = { ...nominal, name: label, elasticity, cycles: withLearning ? transitionCycles : nominal.cycles };
      const result = withLearning ? project(base, snapshot.targets, scenario) : steady;
      const directScenario = { ...nominal, name: `Direto a R$ 5.000 · ${label}`, elasticity, cycles: withLearning ? directCycles : [directCycle] };
      const directResult = withLearning ? project(base, snapshot.targets, directScenario) : directSteady;
      const directAdvantage = round(directResult.scenario.profit - result.scenario.profit);
      const winner = directAdvantage > 0 ? 'direct' as const : directAdvantage < 0 ? 'gradual' as const : 'tie' as const;
      return { id: `${elasticity}-${withLearning}`, label, elasticity, withLearning, scenario, result, learningCost: round(steady.scenario.profit - result.scenario.profit), affectedDays: withLearning ? affectedDays : 0,
        directScenario, directResult, directLearningCost: round(directSteady.scenario.profit - directResult.scenario.profit), directAffectedDays: withLearning ? directAffectedDays : 0, directAdvantage, winner,
        beatsControl: Math.max(result.scenario.profit, directResult.scenario.profit) > result.control.profit };
    });
  });
  return { source: { start, end, loadedAt: snapshot.loadedAt }, base, observedCac: base.ads / base.orders, learning: { ...learning }, start: nominal.start, end: addDays(nominal.start, 97), days: 98, controlBudget: 2790, changes: 6, affectedDays, directAffectedDays, plannedMedia: rows[0].result.scenario.ads, controlMedia: rows[0].result.control.ads, extraMedia: rows[0].result.extraAds, directBudget: 5000, directMedia: rows[0].directResult.scenario.ads, extraDirectMedia: round(rows[0].directResult.scenario.ads - rows[0].result.scenario.ads), rows };
}
export type GradualComparison = ReturnType<typeof compareGradualPlan>;
