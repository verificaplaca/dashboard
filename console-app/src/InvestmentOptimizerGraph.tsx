import { balancedChartScales } from './domain/chart-scales';
import { useEffect, useRef, useState } from 'react';
import type { Chart } from 'chart.js';
import type { Optimum, OptimizationSettings, WindowOptimization } from './domain/investment-optimizer';
const money = (v: number) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const compact = (v: number) => v.toLocaleString('pt-BR', { notation: 'compact', maximumFractionDigits: 1 });
export default function OptimizationGraph({ row, settings, learning, theme, gain }: { row: WindowOptimization; settings: OptimizationSettings; learning: boolean; theme: string; gain: boolean }) {
  const canvas = useRef<HTMLCanvasElement>(null), [failed, setFailed] = useState(false);
  useEffect(() => {
    let disposed = false, chart: Chart<'line', { x: number; y: number }[]> | undefined;
    setFailed(false);
    void import('chart.js/auto').then(({ default: ChartJS }) => {
      if (disposed || !canvas.current || !row.steady || !row.learning || !row.control) return;
      const css = getComputedStyle(document.documentElement), text = css.getPropertyValue('--muted').trim(), border = css.getPropertyValue('--border').trim();
      const offset = gain ? row.control.profit : 0;
      const point = (o: Optimum) => [{ x: o.budget, y: o.profit - offset }];
      const datasets = [
        { label: 'Sem aprendizado', data: row.curve.map(p => ({ x: p.budget, y: p.steady - offset })), borderColor: '#4c83ff', backgroundColor: '#4c83ff', borderWidth: 2.5, pointRadius: 0, pointHoverRadius: 5, borderDash: [] },
        { label: 'Após mudança com aprendizado', data: row.curve.map(p => ({ x: p.budget, y: p.learning - offset })), borderColor: '#d99536', backgroundColor: '#d99536', borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, borderDash: [6, 4] },
        { label: 'Referência: manter verba', data: [{ x: settings.minBudget, y: row.control.profit - offset }, { x: settings.maxBudget, y: row.control.profit - offset }], borderColor: '#93a1b5', backgroundColor: '#93a1b5', borderWidth: 1, pointRadius: 0, pointHoverRadius: 0, borderDash: [3, 5] },
        { label: 'Pico sem aprendizado', data: point(row.steady), borderColor: '#4c83ff', backgroundColor: '#4c83ff', borderWidth: 0, pointRadius: learning ? 5 : 9, pointHoverRadius: 10, borderDash: [] },
        { label: row.learning.status === 'keep' ? 'Melhor decisão com aprendizado: manter' : 'Pico com aprendizado', data: point(row.learning), borderColor: '#d99536', backgroundColor: '#d99536', borderWidth: 0, pointRadius: learning ? 9 : 5, pointHoverRadius: 10, borderDash: [] },
      ];
      chart = new ChartJS(canvas.current, { plugins: [balancedChartScales], type: 'line', data: { datasets }, options: {
        responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'nearest', intersect: false },
        plugins: { legend: { display: false }, tooltip: { callbacks: { title: items => `${money(Number(items[0]?.parsed.x))}/dia`, label: item => `${item.dataset.label}: ${money(Number(item.parsed.y))} ${gain ? 'de lucro adicional' : 'de lucro total'}` } } },
        scales: { x: { type: 'linear', min: settings.minBudget, max: settings.maxBudget, title: { display: true, text: 'Investimento Google Ads por dia (R$)', color: text }, grid: { display: false }, ticks: { color: text, maxTicksLimit: 6, callback: n => `R$ ${compact(Number(n))}` }, border: { color: border } }, y: { beginAtZero: true, title: { display: true, text: `${gain ? 'Ganho de lucro líquido' : 'Lucro líquido total'} em ${settings.days} dias (R$)`, color: text }, ticks: { color: text, maxTicksLimit: 6, callback: n => `R$ ${compact(Number(n))}` }, grid: { color: context => context.tick.value === 0 ? text : border, lineWidth: context => context.tick.value === 0 ? 1.5 : .6 }, border: { display: false } } },
      } });
    }).catch(() => { if (!disposed) setFailed(true); });
    return () => { disposed = true; chart?.destroy(); };
  }, [row, settings, learning, theme, gain]);
  return <div className="sim-profit-canvas sim-opt-canvas">{failed && <p role="status">O gráfico não carregou. Use os valores da tabela.</p>}<canvas ref={canvas} role="img" aria-label={`Investimento diário versus lucro líquido, ${row.label}, com os pontos de maior lucro destacados`}>Tabela de valores disponível abaixo.</canvas></div>;
}
