export const targetFields = [
  ['target_cpa', 'Meta CAC (R$)', '0.01'], ['target_upsell_pct', 'Meta upsell (%)', '0.1'],
  ['monthly_budget', 'Orçamento Ads (R$)', '100'], ['revenue_target', 'Meta receita (R$)', '100'],
  ['profit_target', 'Meta lucro bruto (R$)', '100'], ['net_profit_target', 'Meta lucro líquido (R$)', '100'], ['tax_pct', 'Imposto (%)', '0.01'],
] as const;
export type TargetField = typeof targetFields[number][0];
export type Target = { month: string; updated_at?: string } & Record<TargetField, number | null>;
export function targetPayload(month: string, values: Record<string, string>, now = new Date()): Target {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Escolha um mês válido.');
  const row = { month: `${month}-01`, updated_at: now.toISOString() } as Target;
  for (const [key, label] of targetFields) {
    const text = values[key]?.trim() || '';
    const value = text === '' ? null : Number(text);
    if (value !== null && (!Number.isFinite(value) || value < 0 || (['tax_pct', 'target_upsell_pct'].includes(key) && value > 100))) throw new Error(`Confira ${label}: informe um valor válido.`);
    row[key] = value;
  }
  return row;
}
