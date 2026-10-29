import { readOnlyOverviewFetch } from './overview';
import config from '../production-config.json';
const OVERVIEW_URL = config.url, OVERVIEW_ANON_KEY = config.key;
import { dates, saoPauloToday, taxAt, type Snapshot, type RevenueRow, type BureauRow, type AdsRow, type TaxRow } from './simulation';

export async function rows(table: string, select: string, filter: string, signal?: AbortSignal): Promise<Record<string, unknown>[]> {
  const all: Record<string, unknown>[] = [];
  for (let offset = 0; offset <= 100000; offset += 1000) {
    const r = await readOnlyOverviewFetch(`${OVERVIEW_URL}/rest/v1/${table}?select=${select}&${filter}`, { headers: { apikey: OVERVIEW_ANON_KEY, Authorization: `Bearer ${OVERVIEW_ANON_KEY}`, Range: `${offset}-${offset + 999}` } }, signal);
    if (!r.ok) throw new Error(`Fonte ${table} indisponível (HTTP ${r.status}).`);
    const page: unknown = await r.json();
    if (!Array.isArray(page) || page.some(v => !v || typeof v !== 'object')) throw new Error(`Formato inválido em ${table}.`);
    all.push(...page);
    if (page.length < 1000) return all;
  }
  throw new Error('A fonte excedeu o limite de paginação. Reduza o período.');
}
function numeric(r: Record<string, unknown>, key: string): number {
  if (r[key] === null || r[key] === undefined || r[key] === '' || !Number.isFinite(Number(r[key])) || Number(r[key]) < 0) throw new Error(`Valor ausente ou inválido: ${key}.`);
  return Number(r[key]);
}
function day(r: Record<string, unknown>) { const d = String(r.date).slice(0, 10); dates(d, d); return d; }
export async function loadSimulationData(start: string, end: string, signal?: AbortSignal): Promise<Snapshot> {
  dates(start, end);
  if (end >= saoPauloToday()) throw new Error('Use dias encerrados: o dia atual ainda tem dados parciais.');
  const filter = `date=gte.${start}&date=lte.${end}&order=date.asc`;
  const [revenue, bureau, ads, taxResult] = await Promise.all([
    rows('revenue_daily', 'date,revenue,paid_orders', filter, signal),
    rows('bureau_daily', 'date,custo_bureau', filter, signal),
    rows('google_ads_campaign_daily', 'date,campaign_id,cost_micros', filter + ',campaign_id.asc', signal),
    rows('monthly_targets', 'month,tax_pct', 'order=month.asc', signal).then(value => ({ value, failed: false })).catch(error => { if (signal?.aborted) throw error; return { value: [], failed: true }; })
  ]);
  const targets: TaxRow[] = taxResult.value.map(r => {
    const month = String(r.month).slice(0, 10); dates(month, month);
    const value = r.tax_pct === null ? null : numeric(r, 'tax_pct');
    if (value !== null && value > 100) throw new Error('Imposto inválido na fonte.');
    return { month, tax_pct: value };
  });
  const warnings: string[] = [];
  if (taxResult.failed) warnings.push('Metas indisponíveis: imposto de 8% aplicado como fallback explícito do painel.');
  else if (dates(start, end).some(d => taxAt(targets, d).fallback)) warnings.push('Há datas sem imposto cadastrado: fallback de 8% conforme a Visão Geral.');
  return { start, end, loadedAt: new Date().toISOString(), warnings, targets,
    revenue: revenue.map(r => ({ date: day(r), revenue: numeric(r, 'revenue'), paid_orders: numeric(r, 'paid_orders') } as RevenueRow)),
    bureau: bureau.map(r => ({ date: day(r), custo_bureau: numeric(r, 'custo_bureau') } as BureauRow)),
    ads: ads.map(r => { if (r.campaign_id == null) throw new Error('Campanha sem identificador.'); return { date: day(r), campaign_id: String(r.campaign_id), cost_micros: numeric(r, 'cost_micros') } as AdsRow; })
  };
}
