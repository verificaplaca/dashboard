import { addDays, project, round, type Baseline, type Scenario, type TaxRow } from './simulation';

export type ProfitPoint = { x: number; y: number };
export type ProfitSeries = { id: string; label: string; color: string; dashed?: boolean; points: ProfitPoint[]; cycleProfits: number[]; total: number };
export type ProfitCharts = { metric: 'profit'; cycleLabels: string[]; cycleDates: string[]; series: ProfitSeries[]; budgetSeries: ProfitSeries[]; current: ProfitPoint; control: number };
const elasticityLabel = (n: number) => n.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
/** All chart ordinates are net profit after every informed cost. Never switch to revenue/ROAS. */
export function buildProfitCharts(base: Baseline, targets: TaxRow[], scenario: Scenario): ProfitCharts {
  const current = project(base, targets, scenario);
  const variants = [
    { id: 'conservative', label: 'Conservador · ε 0,6', color: '#d99536', elasticity: .6 },
    { id: 'central', label: 'Central · ε 0,8', color: '#9d78df', elasticity: .8 },
    { id: 'linear', label: 'Linear · ε 1,0', color: '#2aa786', elasticity: 1 },
  ].filter(v => Math.abs(v.elasticity - scenario.elasticity) > 1e-9);
  variants.splice(1, 0, { id: 'current', label: `Seu cenário · ε ${elasticityLabel(scenario.elasticity)}`, color: '#4c83ff', elasticity: scenario.elasticity });
  const points = (profits: number[]) => { let accumulated = 0; return [{ x: 0, y: 0 }, ...profits.map((profit, i) => ({ x: i + 1, y: accumulated = round(accumulated + profit) }))]; };
  const controlProfits = current.cycles.map(c => c.control.profit);
  const series: ProfitSeries[] = [{ id: 'control', label: 'Manter verba', color: '#93a1b5', dashed: true, cycleProfits: controlProfits, total: current.control.profit, points: points(controlProfits) }];
  const budgetSeries: ProfitSeries[] = [];
  for (const variant of variants) {
    const s = { ...scenario, elasticity: variant.elasticity };
    const result = variant.id === 'current' ? current : project(base, targets, s);
    const profits = result.cycles.map(c => c.scenario.profit);
    series.push({ ...variant, total: result.scenario.profit, cycleProfits: profits, points: points(profits) });
    if (current.scenario.ads > 0) {
      const curve = Array.from({ length: 11 }, (_, i) => {
        const scale = .5 + i / 10;
        const result = project(base, targets, { ...s, cycles: scale === 1 ? s.cycles : s.cycles.map(c => ({ ...c, budget: round(c.budget * scale) })) });
        return { x: result.scenario.ads, y: result.scenario.profit };
      });
      budgetSeries.push({ ...variant, total: result.scenario.profit, cycleProfits: [], points: curve });
    }
  }
  if (budgetSeries.length) {
    const xs = budgetSeries[0].points.map(p => p.x);
    budgetSeries.unshift({ id: 'control', label: 'Manter verba · referência atual', color: '#93a1b5', dashed: true, total: current.control.profit, cycleProfits: [], points: [{ x: xs[0], y: current.control.profit }, { x: xs.at(-1)!, y: current.control.profit }] });
  }
  return { metric: 'profit', cycleLabels: ['Início', ...scenario.cycles.map((_, i) => `C${i + 1}`)], cycleDates: [addDays(scenario.start, -1), ...current.cycles.map(c => c.end)], series, budgetSeries, current: { x: current.scenario.ads, y: current.scenario.profit }, control: current.control.profit };
}
