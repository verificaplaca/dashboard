const palette = { positive: '#a8dc69', negative: '#fb8792', attention: '#edc75f', ink: '#242b39', neutral: '#b6bdc9' };
const ids = new Set(['profitChart', 'netProfitChart', 'cacChart', 'upsellChart', 'convChart', 'produtoTesteChart', 'produtoAtualChart', 'lostIsShareChart', 'lostIsEvolutionChart']);

/** Change presentation only. Observations, labels and callback formatters stay original. */
export function styleAnalyticalChart(id: string, config: any) {
  if (!ids.has(id)) return;
  const options = config.options;
  if (options.plugins?.datalabels) options.plugins.datalabels.display = (id === 'profitChart' || id === 'netProfitChart') ? (ctx: any) => ctx.datasetIndex === 0 : false;
  if (options.plugins?.legend?.labels) Object.assign(options.plugins.legend.labels, { usePointStyle: true, pointStyle: 'circle', boxWidth: 6, boxHeight: 6, padding: 14, font: { family: 'Inter, system-ui, sans-serif', size: 10, weight: '400' } });
  if ((id === 'profitChart' || id === 'netProfitChart') && options.plugins?.legend) options.plugins.legend.display = false;
  if (id === 'convChart' || id === 'lostIsShareChart') {
    options.cutout = '76%';
    const colors = ['#a8dc69', '#a4cbd0', '#b9b0cf', '#edc75f', '#f0ad98', '#9da8bf', '#7b9f84', '#c5cbb4', '#d6b7c5'];
    config.data.datasets.forEach((series: any) => { series.backgroundColor = series.data.map((_: unknown, index: number) => colors[index % colors.length]); series.borderWidth = 4; series.borderRadius = 4; });
    return;
  }
  options.layout = { ...options.layout, padding: { top: 12, right: 8, bottom: 4, left: 0 } };
  // Give daily results room above peaks and below losses, including comparisons.
  // Keep the bar chart's zero baseline and let Chart.js choose readable ticks.
  if ((id === 'profitChart' || id === 'netProfitChart') && options.scales?.y) {
    options.scales.y.grace = '20%';
    options.scales.y.beginAtZero = true;
    options.scales.y.bounds = 'data';
  }
  for (const [name, scale] of Object.entries<any>(options.scales || {})) {
    if (scale.ticks) scale.ticks.maxTicksLimit = name === 'x' ? 6 : 5;
  }
  for (const series of config.data.datasets) {
    if (series.type === 'line' || config.type === 'line') {
      if (id === 'lostIsEvolutionChart') series.borderColor = /budget|orçamento/i.test(series.label) ? '#fb8792' : '#b3a0db';
      series.borderWidth = 1.6; series.pointRadius = 0; series.pointHoverRadius = 4;
      if (series.label === 'Média' || series.label === 'Meta') series.borderColor = '#9ea6b4';
      continue;
    }
    series.borderRadius = 5; series.barPercentage = .72; series.categoryPercentage = .82;
    if (id === 'profitChart' || id === 'netProfitChart') series.backgroundColor = series.data.map((value: number | null) => value === null || !Number.isFinite(value) ? 'transparent' : value < 0 ? palette.negative : palette.positive);
    if (id === 'cacChart' && Array.isArray(series.backgroundColor)) series.backgroundColor = series.backgroundColor.map((color: string) => color.includes('22,163,74') ? palette.positive : color.includes('217,119,6') ? palette.attention : color.includes('220,38,38') ? palette.negative : color);
    if (id === 'upsellChart' && Array.isArray(series.backgroundColor)) series.backgroundColor = series.backgroundColor.map((color: string) => color.includes('124,58,237,.35') ? palette.negative : color.includes('124,58,237,.7') ? palette.positive : color);
  }
  // A presentation ceiling must not hide an observation above the legacy 40% scale.
  if (id === 'upsellChart' && Number.isFinite(options.scales?.y?.max)) {
    const max = Math.max(0, ...config.data.datasets.flatMap((series: any) => series.data.filter((value: unknown) => typeof value === 'number' && Number.isFinite(value))));
    if (max > options.scales.y.max) options.scales.y.max = Math.ceil(max * 1.1 / 10) * 10;
  }
}
