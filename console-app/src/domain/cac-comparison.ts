import { investmentGrowth, validateOptimization, type OptimizationSettings } from './investment-optimizer';
import type { Baseline, TaxRow } from './simulation';
export const CAC_BUDGETS = [3400, 4100, 5000] as const;
export type CacProjection = ReturnType<typeof investmentGrowth>;
export type CacComparisonRow = { budget: number; constant: CacProjection; constantLearning: CacProjection; elastic: CacProjection; elasticLearning: CacProjection };
export function compareCacScenarios(base: Baseline, targets: TaxRow[], settings: OptimizationSettings): CacComparisonRow[] {
  // Both hypotheses use the same paid-order base and a fully media-responsive population.
  // Elasticity 1 then preserves the observed Ads/paid-order ratio before any transition loss.
  const p = { ...settings, responsivePct: 100, elasticity: .8 };
  const errors = [...base.issues, ...validateOptimization(p)];
  if (errors.length) throw new Error(errors.join(' '));
  return CAC_BUDGETS.map(budget => ({ budget,
    constant: investmentGrowth(base, targets, { ...p, elasticity: 1 }, budget, false),
    constantLearning: investmentGrowth(base, targets, { ...p, elasticity: 1 }, budget, true),
    elastic: investmentGrowth(base, targets, p, budget, false),
    elasticLearning: investmentGrowth(base, targets, p, budget, true),
  }));
}
