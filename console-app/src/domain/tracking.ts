export type ChannelStatus = 'success' | 'failed' | 'pending' | 'skipped' | string;
export interface TrackingDispatch {
  checkout_id: string; order_nsu: string | null; event_id: string | null;
  value: number | null; currency: string | null;
  ga4_status: ChannelStatus; ads_status: ChannelStatus; meta_status: ChannelStatus;
  ga4_error?: string | null; ads_error?: string | null; meta_error?: string | null;
  attempts: number; created_at: string; updated_at?: string; paid_at: string | null;
  gclid?: string | null; gbraid?: string | null; wbraid?: string | null;
  fbp?: string | null; fbc?: string | null; ga_client_id?: string | null;
  event_source_url?: string | null; has_email?: boolean; has_phone?: boolean;
}
export const trackingChannels = [{ key: 'ga4', label: 'GA4' }, { key: 'ads', label: 'Google Ads' }, { key: 'meta', label: 'Meta CAPI' }] as const;
export const trackingStatuses = { success: 'Enviado', failed: 'Falhou', pending: 'Pendente', skipped: 'Não enviado' };
export const channelStatus = (row: TrackingDispatch, key: typeof trackingChannels[number]['key']) => row[`${key}_status`];
export const googleClickId = (row: TrackingDispatch) => row.gclid || row.gbraid || row.wbraid || null;
export const orderIsFull = (row: TrackingDispatch) => trackingChannels.every(({ key }) => ['success', 'skipped'].includes(channelStatus(row, key)));
export const orderHasFailure = (row: TrackingDispatch) => trackingChannels.some(({ key }) => ['failed', 'pending'].includes(channelStatus(row, key)));
export function matchIdsDetail(row: TrackingDispatch) {
  const ids: [string, unknown][] = [['Email', row.has_email], ['Telefone', row.has_phone], ['fbp', row.fbp], ['fbc', row.fbc], ['GA client ID', row.ga_client_id], ['Google click ID', googleClickId(row)]];
  return { present: ids.filter(([, value]) => !!value).map(([label]) => label), missing: ids.filter(([, value]) => !value).map(([label]) => label) };
}
export function matchQuality(row: TrackingDispatch) {
  const count = matchIdsDetail(row).present.length;
  return count >= 4 ? 'Forte' : count >= 2 ? 'Médio' : 'Fraco';
}
export function eventIdCheck(row: TrackingDispatch) {
  if (!row.event_id || row.event_id === 'unknown') return { status: 'pending', label: 'Sem event ID', expected: row.order_nsu ? `evt_${row.order_nsu}` : null };
  if (!row.order_nsu) return { status: 'pending', label: 'Sem pedido', expected: null };
  const expected = `evt_${row.order_nsu}`;
  return { status: row.event_id === expected ? 'success' : 'failed', label: row.event_id === expected ? 'ID alinhado' : 'ID divergente', expected };
}
export function dispatchLatency(row: TrackingDispatch) {
  if (!row.paid_at || !row.created_at) return null;
  const ms = Date.parse(row.created_at) - Date.parse(row.paid_at);
  if (!Number.isFinite(ms)) return null;
  const abs = Math.abs(ms);
  const label = abs < 1000 ? '<1s' : abs < 60000 ? `${Math.round(abs / 1000)}s` : abs < 3600000 ? `${Math.round(abs / 60000)}min` : `${(abs / 3600000).toFixed(1).replace('.', ',')}h`;
  return { milliseconds: ms, label: ms < 0 ? `−${label}` : label, status: ms < 0 ? 'pending' : abs >= 86400000 ? 'failed' : abs >= 3600000 ? 'pending' : 'success' };
}
export function trackingSummary(rows: TrackingDispatch[], now: number) {
  const week = rows.filter(row => Date.parse(row.created_at) >= now - 7 * 86400000);
  const total = week.length, full = week.filter(orderIsFull).length, failing = week.filter(orderHasFailure).length;
  return { total, full, failing, quality: total ? full / total * 100 : null, attempts: total ? week.reduce((sum, row) => sum + (Number(row.attempts) || 0), 0) / total : null,
    channels: trackingChannels.map(channel => ({ ...channel, success: week.filter(row => channelStatus(row, channel.key) === 'success').length, failed: week.filter(row => channelStatus(row, channel.key) === 'failed').length, pending: week.filter(row => channelStatus(row, channel.key) === 'pending').length, skipped: week.filter(row => channelStatus(row, channel.key) === 'skipped').length })),
    attribution: [{ label: 'Google click ID', detail: 'gclid / gbraid / wbraid', count: week.filter(googleClickId).length }, { label: 'Meta fbc', detail: 'Identificador de clique', count: week.filter(row => row.fbc).length }, { label: 'Meta fbp', detail: 'Identificador do navegador', count: week.filter(row => row.fbp).length }, { label: 'GA4 client_id', detail: 'Identificador do cliente', count: week.filter(row => row.ga_client_id).length }],
  };
}
export interface TrackingFilters { status: string; channel: string; attribution: string; search: string; min: string; max: string; }
export const emptyTrackingFilters: TrackingFilters = { status: '', channel: '', attribution: '', search: '', min: '', max: '' };
export function filterTracking(rows: TrackingDispatch[], filter: TrackingFilters) {
  const search = filter.search.trim().toLowerCase();
  return rows.filter(row => {
    if (filter.status === 'full' && !orderIsFull(row) || filter.status === 'partial' && orderIsFull(row)) return false;
    if (['failed', 'pending'].includes(filter.status) && !trackingChannels.some(({ key }) => channelStatus(row, key) === filter.status)) return false;
    if (filter.channel && !['failed', 'pending'].includes(row[`${filter.channel}_status` as keyof TrackingDispatch] as string)) return false;
    if (filter.attribution === 'com_g' && !googleClickId(row) || filter.attribution === 'sem_g' && googleClickId(row) || filter.attribution === 'sem_nada' && matchIdsDetail(row).present.length || filter.attribution === 'match_fraco' && matchQuality(row) !== 'Fraco') return false;
    if (filter.min !== '' && Number(row.value) < Number(filter.min) || filter.max !== '' && Number(row.value) > Number(filter.max)) return false;
    const value = row.value == null ? '' : Number(row.value).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    return !search || [row.order_nsu, row.event_id, row.checkout_id, row.value, value, value.replace(/\./g, ''), value.replace(/[^\d,.-]/g, ''), value.replace(/[^\d,.-]/g, '').replace(/\./g, '')].filter(v => v != null).join(' ').toLowerCase().includes(search);
  });
}
const localDate = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
export function trackingWindows(now = new Date()) {
  const from = new Date(now); from.setDate(from.getDate() - 90);
  const cac = new Date(now); cac.setDate(cac.getDate() - 29);
  return { from: from.toISOString(), until: now.toISOString(), cacFrom: localDate(cac), cacTo: localDate(now) };
}
export function dailyDispatches(rows: TrackingDispatch[], days: number, now: Date) {
  const cutoff = new Date(now); cutoff.setDate(cutoff.getDate() - days + 1); cutoff.setHours(0, 0, 0, 0);
  const byDay = new Map<string, { ga4: number; ads: number; meta: number; fail: number }>();
  for (const row of rows) {
    if (new Date(row.created_at) < cutoff) continue;
    const key = new Date(row.created_at).toISOString().slice(0, 10), value = byDay.get(key) || { ga4: 0, ads: 0, meta: 0, fail: 0 };
    trackingChannels.forEach(({ key }) => { if (channelStatus(row, key) === 'success') value[key]++; });
    if (trackingChannels.some(({ key }) => channelStatus(row, key) === 'failed')) value.fail++;
    byDay.set(key, value);
  }
  return Array.from({ length: days }, (_, index) => { const date = new Date(now); date.setDate(date.getDate() - days + 1 + index); const key = date.toISOString().slice(0, 10); return { date: key, ...byDay.get(key) || { ga4: 0, ads: 0, meta: 0, fail: 0 } }; });
}
export interface TrackingRevenue { date: string; revenue: number; paid_orders: number; }
export interface TrackingAds { date: string; cost_micros: number | string; conversions: number | string; }
export function trackingCac(revenue: TrackingRevenue[], ads: TrackingAds[]) {
  const rev = new Map(revenue.map(row => [row.date.slice(0, 10), row])), cost = new Map<string, number>(), conv = new Map<string, number>();
  for (const row of ads) { const date = row.date.slice(0, 10); cost.set(date, (cost.get(date) || 0) + (parseInt(String(row.cost_micros)) || 0) / 1e6); conv.set(date, (conv.get(date) || 0) + (parseFloat(String(row.conversions)) || 0)); }
  const daily = [...new Set([...rev.keys(), ...cost.keys()])].sort().map(date => {
    const spend = cost.get(date) || 0, paid = parseInt(String(rev.get(date)?.paid_orders)) || 0, reported = conv.get(date) || 0;
    const realCac = paid > 0 && spend > 0 ? spend / paid : null, reportedCac = reported > 0 && spend > 0 ? spend / reported : null;
    return { date, spend, paid, reported, realCac, reportedCac, gap: realCac !== null && reportedCac !== null && realCac > 0 ? (realCac - reportedCac) / realCac * 100 : null };
  });
  const spend = daily.reduce((sum, row) => sum + row.spend, 0), paid = daily.reduce((sum, row) => sum + row.paid, 0), reported = daily.reduce((sum, row) => sum + row.reported, 0);
  const realCac = paid ? spend / paid : null, reportedCac = reported ? spend / reported : null;
  return { daily: daily.reverse(), spend, paid, reported, realCac, reportedCac, gap: realCac && reportedCac ? (realCac - reportedCac) / realCac * 100 : null };
}

/** Read the complete 90-day window, not just the visible page or API row cap. */
export async function readTrackingOrders(read: (query: URLSearchParams, signal: AbortSignal) => Promise<TrackingDispatch[]>, window: ReturnType<typeof trackingWindows>, signal: AbortSignal) {
  return readTrackingPages(read, new URLSearchParams({ select: '*', and: `(created_at.gte.${window.from},created_at.lte.${window.until})`, order: 'created_at.desc,checkout_id.asc' }), signal);
}

export async function readTrackingPages<T>(read: (query: URLSearchParams, signal: AbortSignal) => Promise<T[]>, base: URLSearchParams, signal: AbortSignal) {
  const controller = new AbortController(), combined = AbortSignal.any([signal, controller.signal, AbortSignal.timeout(90_000)]);
  const size = 500;
  const page = (offset: number) => { const query = new URLSearchParams(base); query.set('limit', String(size)); query.set('offset', String(offset)); return read(query, combined); };
  try {
    let rows = await page(0); combined.throwIfAborted();
    if (rows.length < size) return rows;
    for (let offset = size; ; offset += size * 4) {
      const pages = await Promise.all([0, 1, 2, 3].map(index => page(offset + index * size))); combined.throwIfAborted();
      for (const batch of pages) { rows = rows.concat(batch); if (batch.length < size) return rows; }
    }
  } finally { controller.abort(); }
}
