/** Presentation only. All amounts and targets are supplied by the preserved overview. */
export type MetricTone = 'positive' | 'negative' | 'neutral' | 'attention';
export function comparisonStatus(delta: number | null, lower = false, costOnly = false) {
  if (delta === null || !Number.isFinite(delta)) return { tone: 'neutral' as MetricTone, text: 'Base indisponível' };
  if (Math.abs(delta) < .005) return { tone: 'neutral' as MetricTone, text: 'Sem mudança' };
  if (costOnly) return { tone: 'neutral' as MetricTone, text: delta > 0 ? 'Custo maior' : 'Custo menor' };
  const better = lower ? delta < 0 : delta > 0;
  return { tone: (better ? 'positive' : 'negative') as MetricTone, text: better ? 'Melhora' : 'Piora' };
}
export function targetStatus(value: number | null, target: number | null, lower = false) {
  if (value === null || !Number.isFinite(value)) return { tone: 'neutral' as MetricTone, text: 'Sem dados' };
  if (target === null || !Number.isFinite(target) || target < 0) return { tone: 'neutral' as MetricTone, text: 'Meta não definida' };
  const met = lower ? value <= target : value >= target;
  return { tone: (met ? 'positive' : 'negative') as MetricTone, text: met ? '✓ Dentro da meta' : '! Fora da meta' };
}
export function resultStatus(value: number | null) {
  return value === null || !Number.isFinite(value) ? { tone: 'neutral' as MetricTone, text: 'Sem dados' }
    : value < 0 ? { tone: 'negative' as MetricTone, text: '− Resultado negativo' }
    : value > 0 ? { tone: 'positive' as MetricTone, text: '+ Resultado positivo' }
    : { tone: 'neutral' as MetricTone, text: 'Resultado zero' };
}
export function goalProgress(value: number, target: number) {
  if (!Number.isFinite(value) || !Number.isFinite(target) || target <= 0) return null;
  return { percent: value / target * 100, track: Math.max(0, Math.min(100, value / target * 100)), remaining: target - value };
}
