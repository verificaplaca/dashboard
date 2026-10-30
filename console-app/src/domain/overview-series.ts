export function netMarginPercent(revenue: number, netProfit: number): number | null {
  return Number.isFinite(revenue) && Number.isFinite(netProfit) && revenue > 0 ? netProfit / revenue * 100 : null;
}
export function costValue(day: { cost?: number; costBureau?: number; costTotal?: number }, mode: string): number | null {
  const value = mode === 'ads' ? day.cost : mode === 'bureau' ? day.costBureau : day.costTotal;
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}
