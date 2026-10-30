/** Presentation-only chart comparisons. Missing calendar days stay missing. */
export interface OverviewDay {
  dateStr: string;
  revenue?: number;
  costTotal?: number;
  profit?: number;
  netProfit?: number;
}
export interface ChartSnapshot {
  current: OverviewDay[];
  allDays: OverviewDay[];
  compare: boolean;
  reducedMotion: boolean;
  selectedDate?: string;
}
const dayMs = 86400000;
const timestamp = (date: string) => Date.parse(date + 'T00:00:00Z');
const iso = (time: number) => new Date(time).toISOString().slice(0, 10);
const finite = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value);
const brl = (value: number) => value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const supported = new Set(['revenueChart', 'profitChart', 'netProfitChart', 'cacChart', 'upsellChart']);
const fields: Record<string, keyof OverviewDay> = { revenueChart: 'revenue', profitChart: 'profit', netProfitChart: 'netProfit' };

export function comparisonSeries(current: OverviewDay[], allDays: OverviewDay[], field: keyof OverviewDay, length = current.length) {
  if (!current.length) return { data: Array.from({ length }, () => null as number | null), dates: [] as string[] };
  const start = timestamp(current[0].dateStr), end = timestamp(current[current.length - 1].dateStr);
  if (!Number.isFinite(start) || !Number.isFinite(end) || end < start) return { data: Array.from({ length }, () => null as number | null), dates: [] as string[] };
  const duration = Math.round((end - start) / dayMs) + 1;
  const byDate = new Map(allDays.map(day => [day.dateStr, day]));
  const dates = current.map(day => iso(timestamp(day.dateStr) - duration * dayMs));
  const data: (number | null)[] = dates.map(date => { const value = byDate.get(date)?.[field]; return finite(value) ? value : null; });
  while (data.length < length) data.push(null);
  return { data, dates };
}

export function targetDistance(actual: number | null, target: number | null, lowerIsBetter = false, percent = false) {
  if (actual === null || target === null || !Number.isFinite(actual) || !Number.isFinite(target)) return null;
  const difference = actual - target;
  const formatted = percent ? Math.abs(difference).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' p.p.' : brl(Math.abs(difference));
  return { tone: difference === 0 || (lowerIsBetter ? difference < 0 : difference > 0) ? 'positive' : 'negative', text: difference === 0 ? 'Na meta' : `${formatted} ${difference > 0 ? 'acima' : 'abaixo'} ${lowerIsBetter ? 'do teto' : 'da meta'}` };
}

/** Call after the existing chart styling. Does not change original datasets or callbacks. */
export function enhanceOverviewChart(id: string, config: any, snapshot: ChartSnapshot) {
  if (!supported.has(id)) return;
  const options = config.options ||= {};
  options.animation = snapshot.reducedMotion ? false : { ...(typeof options.animation === 'object' ? options.animation : {}), duration: 240 };
  options.transitions = { ...options.transitions, active: { animation: { duration: snapshot.reducedMotion ? 0 : 140 } } };
  const field = fields[id];
  if (field && snapshot.compare) {
    const comparison = comparisonSeries(snapshot.current, snapshot.allDays, field, config.data.labels?.length);
    // An unavailable prior window does not become a synthetic zero line.
    if (comparison.data.some(finite)) {
      config.data.datasets.push({ label: 'Período anterior', data: comparison.data, type: 'line', borderColor: '#929baa', backgroundColor: 'transparent', borderDash: [6, 5], borderWidth: 1.5, pointRadius: 0, pointHoverRadius: 4, tension: .2, spanGaps: false, fill: false, order: 0 });
      const tooltip = (options.plugins ||= {}).tooltip ||= {};
      const callbacks = tooltip.callbacks ||= {};
      const originalAfterBody = callbacks.afterBody;
      callbacks.afterBody = function(items: any[]) {
        const original = originalAfterBody?.call(this, items);
        const lines = original ? Array.isArray(original) ? [...original] : [String(original)] : [];
        const index = items[0]?.dataIndex;
        if (index === undefined || index >= snapshot.current.length) return lines;
        const previous = comparison.data[index], current = snapshot.current[index]?.[field];
        if (!finite(previous) || !finite(current)) return [...lines, 'Comparação: sem dados no dia anterior equivalente'];
        const difference = current - previous;
        const percent = previous === 0 ? '' : ` (${difference > 0 ? '+' : ''}${(difference / Math.abs(previous) * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%)`;
        return [...lines, `Dia equivalente: ${comparison.dates[index].split('-').reverse().join('/')}`, `Variação: ${difference > 0 ? '+' : ''}${brl(difference)}${percent}`];
      };
    }
  }
  // Draw explanatory labels without adding fake observations or extending scales.
  const selectedIndex = snapshot.current.findIndex(day => day.dateStr === snapshot.selectedDate);
  config.plugins ||= [];
  config.plugins.push({
    id: 'overview-reading-guides',
    afterDatasetsDraw(chart: any) {
      const { ctx, chartArea: area, scales } = chart;
      if (!area) return;
      const activeIndex = chart.getActiveElements?.()[0]?.index;
      const index = activeIndex ?? chart.$overviewSelectedIndex ?? selectedIndex;
      if (index >= 0 && scales.x) {
        const x = scales.x.getPixelForValue(index);
        if (x >= area.left && x <= area.right) {
          ctx.save(); ctx.strokeStyle = '#78849488'; ctx.lineWidth = 1; ctx.setLineDash([3, 4]); ctx.beginPath(); ctx.moveTo(x, area.top); ctx.lineTo(x, area.bottom); ctx.stroke(); ctx.restore();
        }
      }
      if (!['cacChart', 'upsellChart'].includes(id)) return;
      const series = chart.data.datasets.find((dataset: any) => /meta/i.test(dataset.label || ''));
      const value = series?.data?.find(finite);
      const y = finite(value) ? scales.y?.getPixelForValue(value) : NaN;
      if (!Number.isFinite(y) || y < area.top || y > area.bottom) return;
      const label = id === 'cacChart' ? `Teto ${brl(value)}` : `Meta ${value.toLocaleString('pt-BR')}%`;
      ctx.save(); ctx.font = '11px Inter, system-ui, sans-serif';
      const width = ctx.measureText(label).width + 12;
      const dark = chart.canvas.getRootNode()?.host?.dataset?.theme === 'dark';
      ctx.fillStyle = dark ? '#152132ed' : '#ffffffed'; ctx.fillRect(area.right - width, Math.max(area.top, y - 18), width, 18);
      ctx.fillStyle = dark ? '#d0d7e3' : '#505b6d'; ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(label, area.right - 5, Math.max(area.top + 9, y - 9)); ctx.restore();
    },
  });
}

/** Persist a selected day while its contextual detail is open. Pass -1 on close. */
export function selectOverviewChartDay(chart: any, index: number) {
  chart.$overviewSelectedIndex = index;
  chart.draw();
}
