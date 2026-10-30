import type { MetricTone } from './metric-presentation';

type Totals = Partial<Record<'rev' | 'cost' | 'costTotal' | 'conv' | 'pft' | 'netPftOpt' | 'cac' | 'margin' | 'upsell', number>>;
export interface InsightSnapshot {
  totals: Totals;
  previous?: Totals | null;
  targets: { cpa?: number; upsell?: number; netProfit?: number };
  months: number;
  real: boolean;
  dayCount?: number;
  upsellOrders?: number | null;
}
export interface MetricDetail {
  title: string;
  value: string;
  formula: string;
  rows: { label: string; value: string }[];
  comparison: string;
  target: string;
  tone: MetricTone;
  context: string;
}
const finite = (n: unknown): n is number => typeof n === 'number' && Number.isFinite(n);
const number = (n: unknown, digits = 2) => finite(n) ? n.toLocaleString('pt-BR', { maximumFractionDigits: digits }) : '—';
const brl = (n: unknown) => finite(n) ? n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : '—';
const percent = (n: unknown) => finite(n) ? number(n, 1) + '%' : '—';
const hasDays = (s: InsightSnapshot) => s.dayCount === undefined || s.dayCount > 0;
const available = (key: string, t: Totals) => key === 'cac'
  ? finite(t.conv) && t.conv > 0 && finite(t.cost) && t.cost > 0 && finite(t.cac)
  : key === 'margin' ? finite(t.rev) && t.rev > 0 && finite(t.margin)
  : finite(t.conv) && t.conv > 0 && finite(t.upsell);

function comparison(key: string, current: number, previous: Totals | null | undefined): string {
  if (!previous || !available(key, previous)) return 'Comparação indisponível: sem base válida no período anterior.';
  const prior = previous[key as keyof Totals]!;
  const delta = current - prior;
  if (Math.abs(delta) < (key === 'margin' ? .00005 : .005)) return 'Estável em relação ao período anterior.';
  const better = key === 'cac' ? delta < 0 : delta > 0;
  const distance = key === 'cac' ? brl(Math.abs(delta)) : number(Math.abs(delta) * (key === 'margin' ? 100 : 1), 2) + ' p.p.';
  return `${delta > 0 ? 'Aumento' : 'Redução'} de ${distance} vs. período anterior · ${better ? 'melhora' : 'piora'}.`;
}

/** Narrative from existing aggregates only; no inferred causes or forecasts. */
export function buildPeriodSummary(s: InsightSnapshot): string {
  if (!hasDays(s) || !finite(s.totals.netPftOpt)) return 'Sem dados no período selecionado. Escolha outra janela para analisar o resultado.';
  const result = s.totals.netPftOpt;
  const parts = [result > 0 ? 'Resultado positivo' : result < 0 ? 'Resultado negativo' : 'Resultado zero'];
  if (s.months === 1 && finite(s.targets.netProfit) && s.targets.netProfit > 0) parts.push(result >= s.targets.netProfit ? 'meta mensal atingida' : 'abaixo da meta mensal');
  else if (s.months > 1) parts.push('período abrange mais de um mês');
  const cacCurrent = available('cac', s.totals), cacPrevious = s.previous && available('cac', s.previous);
  const cac = cacCurrent && cacPrevious && Math.abs(s.totals.cac! - s.previous!.cac!) >= .005
    ? ` CAC ${s.totals.cac! < s.previous!.cac! ? 'melhorou' : 'piorou'} em relação ao período anterior.` : '';
  return `${s.real ? '' : 'Dados de exemplo · '}${parts.join(', ')}.${cac}`;
}

export function buildMetricDetail(key: 'cac' | 'margin' | 'upsell', s: InsightSnapshot): MetricDetail {
  const t = s.totals;
  const valid = hasDays(s) && available(key, t);
  const value = t[key];
  const shared = `${s.real ? 'Dados reais' : 'Dados de exemplo'} · seleção atual. Comparação com a janela anterior de mesma duração.`;
  let detail: MetricDetail;
  if (key === 'cac') {
    detail = { title: 'Custo de aquisição', value: valid ? brl(value) : '—', formula: 'Google Ads ÷ pedidos pagos', rows: [{ label: 'Google Ads', value: brl(t.cost) }, { label: 'Pedidos pagos', value: number(t.conv, 0) }], comparison: '', target: '', tone: 'neutral', context: shared + ' Não inclui bureau. O CAC atribuído pelo Google Ads pode ser diferente.' };
  } else if (key === 'margin') {
    detail = { title: 'Margem de contribuição', value: valid ? percent(value! * 100) : '—', formula: 'Lucro bruto ÷ receita bruta × 100', rows: [{ label: 'Receita bruta', value: brl(t.rev) }, { label: 'COGS · Ads + bureau', value: brl(t.costTotal) }, { label: 'Lucro bruto', value: brl(t.pft) }], comparison: '', target: 'Referência saudável: 22%. A barra de eficiência usa referência de 30%; não é uma meta mensal configurada.', tone: valid ? value! >= .22 ? 'positive' : 'attention' : 'neutral', context: shared + ' Margem antes de imposto e despesas complementares da DRE.' };
  } else {
    detail = { title: 'Taxa de upsell', value: valid ? percent(value) : '—', formula: s.real ? 'Pedidos com adicional ÷ pedidos pagos × 100' : 'Média das taxas diárias de exemplo', rows: [{ label: 'Pedidos pagos', value: number(t.conv, 0) }, { label: 'Pedidos com adicional', value: s.real ? number(s.upsellOrders, 0) : 'Não disponível na demonstração' }], comparison: '', target: '', tone: 'neutral', context: shared + (s.real ? ' Taxa ponderada pelo volume de pedidos.' : ' A demonstração conserva o cálculo original por média diária; não estima a quantidade de adicionais.') };
  }
  detail.comparison = valid ? comparison(key, value!, s.previous) : 'Sem base válida para calcular este indicador no período.';
  if (key !== 'margin') {
    const goal = key === 'cac' ? s.targets.cpa : s.targets.upsell;
    if (s.months !== 1) detail.target = 'Metas variam por mês. Escolha um período dentro de um mês para comparar.';
    else if (!finite(goal) || goal <= 0) detail.target = 'Meta não definida para comparação.';
    else if (!valid) detail.target = `${key === 'cac' ? 'Teto' : 'Meta'}: ${key === 'cac' ? brl(goal) : percent(goal)} · sem dados para comparação.`;
    else {
      const met = key === 'cac' ? value! <= goal : value! >= goal;
      const delta = Math.abs(value! - goal);
      detail.tone = met ? 'positive' : 'negative';
      const distance = key === 'cac' ? brl(delta) : number(delta, 2) + ' p.p.';
      detail.target = delta < .005 ? 'Exatamente na meta.' : `${distance} ${value! > goal ? 'acima' : 'abaixo'} ${key === 'cac' ? 'do teto' : 'da meta'} · ${met ? 'dentro da meta' : 'fora da meta'}.`;
    }
  }
  return detail;
}
