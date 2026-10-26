import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import type { Chart } from 'chart.js';
import { LockKeyhole, TrendingUp } from 'lucide-react';
import { Card } from './components';
import { useUI } from './ui';
import { buildProfitCharts, type ProfitSeries } from './domain/simulation-charts';
import type { Baseline, Scenario, TaxRow } from './domain/simulation';
const money = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const compact = (n: number) => `${n < 0 ? '−' : ''}R$ ${Math.abs(n) >= 1000 ? `${(Math.abs(n) / 1000).toLocaleString('pt-BR', { maximumFractionDigits: 1 })} mil` : Math.abs(n).toLocaleString('pt-BR', { maximumFractionDigits: 0 })}`;
const date = (s: string) => s.split('-').reverse().join('/');

function ProfitChart({ series, labels, dates, mode, title, current, theme }: { series: ProfitSeries[]; labels: string[]; dates?: string[]; mode: 'accumulated' | 'cycles' | 'budget'; title: string; current?: { x: number; y: number }; theme: string }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const [error, setError] = useState(false);
  useEffect(() => {
    let disposed = false, chart: Chart<'line', { x: number; y: number }[]> | undefined;
    setError(false);
    void import('chart.js/auto').then(({ default: ChartJS }) => {
      if (disposed || !canvas.current) return;
      const css = getComputedStyle(document.documentElement), text = css.getPropertyValue('--muted').trim(), border = css.getPropertyValue('--border').trim();
      const datasets = series.map(s => ({
        label: s.label, data: mode === 'cycles' ? s.cycleProfits.map((y, i) => ({ x: i + 1, y })) : s.points,
        borderColor: s.color, backgroundColor: s.color, borderWidth: s.id === 'current' ? 3 : 2, borderDash: s.dashed ? [6, 5] : [],
        pointRadius: mode === 'budget' ? 2 : s.id === 'current' ? 4 : 3, pointHoverRadius: 6, pointHitRadius: 12, tension: 0, fill: false,
      }));
      if (mode === 'budget' && current) datasets.push({ label: 'Seu plano atual', data: [current], borderColor: '#4c83ff', backgroundColor: '#4c83ff', borderWidth: 0, borderDash: [], pointRadius: 7, pointHoverRadius: 9, pointHitRadius: 15, tension: 0, fill: false });
      chart = new ChartJS(canvas.current, {
        type: 'line', data: { datasets }, options: {
          responsive: true, maintainAspectRatio: false, animation: false,
          interaction: { mode: mode === 'budget' ? 'nearest' : 'index', intersect: false },
          plugins: { legend: { display: false }, tooltip: { padding: 12, callbacks: {
            title: items => { const x = items[0]?.parsed.x; return mode === 'budget' ? `Investimento: ${money(Number(x))}` : `${labels[Number(x)] || ''}${dates?.[Number(x)] ? ` · ${date(dates[Number(x)])}` : ''}`; },
            label: item => `${item.dataset.label}: ${money(Number(item.parsed.y))}`,
          } } },
          scales: {
            x: { type: 'linear', afterBuildTicks: axis => { if (mode !== 'budget') axis.ticks = labels.map((_, i) => ({ value: i })).filter(t => mode !== 'cycles' || t.value > 0); }, min: mode === 'budget' ? undefined : mode === 'cycles' ? 1 : 0, max: mode === 'budget' ? undefined : labels.length - 1,
              title: { display: true, text: mode === 'budget' ? 'Investimento Google Ads no período' : 'Ciclos da projeção', color: text, font: { size: 11 } },
              grid: { display: false }, border: { color: border }, ticks: { color: text, autoSkip: mode === 'budget' || labels.length > 12, maxRotation: 0, maxTicksLimit: mode === 'budget' || labels.length > 12 ? 7 : labels.length, precision: 0, stepSize: mode === 'budget' ? undefined : 1, callback: value => mode === 'budget' ? compact(Number(value)) : labels[Number(value)] || '', font: { size: 10 } },
            },
            y: { beginAtZero: true, title: { display: true, text: 'Lucro líquido (R$)', color: text, font: { size: 11 } },
              grid: { color: context => context.tick.value === 0 ? text : border, lineWidth: context => context.tick.value === 0 ? 1.5 : .6 }, border: { display: false }, ticks: { color: text, maxTicksLimit: 6, callback: value => compact(Number(value)), font: { size: 10 } },
            },
          },
        },
      });
    }).catch(() => { if (!disposed) setError(true); });
    return () => { disposed = true; chart?.destroy(); };
  }, [series, labels, dates, mode, current, theme]);
  return <div className="sim-profit-canvas">{error && <p role="status">O gráfico não carregou. Os valores continuam disponíveis na tabela abaixo.</p>}<canvas ref={canvas} role="img" aria-label={title}>Lucro líquido dos cenários. Consulte a tabela de valores logo após o gráfico.</canvas></div>;
}
export default function SimulationCharts({ base, targets, scenario, optimizer, showSummary = true }: { base: Baseline; targets: TaxRow[]; scenario: Scenario; optimizer?: ReactNode; showSummary?: boolean }) {
  const { theme } = useUI();
  const [mode, setMode] = useState<'accumulated' | 'cycles'>('accumulated');
  const model = useMemo(() => buildProfitCharts(base, targets, scenario), [base, targets, scenario]);
  const actual = model.series.find(s => s.id === 'current')!;
  const delta = actual.total - model.control;
  return <>
    {showSummary && <div className="sim-profit-focus">
      <div className="sim-profit-lock"><LockKeyhole size={17} /><div><strong>Lucro líquido</strong><span>Objetivo fixo · cenário aberto</span></div></div>
      <div><span>Manter · cenário aberto</span><strong>{money(model.control)}</strong></div>
      <div><span>Seu cenário</span><strong>{money(actual.total)}</strong></div>
      <div><span>Diferença no período</span><strong className={delta >= 0 ? 'sim-good' : 'sim-bad'}>{money(delta)}</strong></div>
    </div>
    }
    {optimizer}
    <section className="sim-profit-section" aria-labelledby="sim-profit-title">
    <div className="sim-profit-heading"><div><h2 id="sim-profit-title"><TrendingUp size={19} />Gráficos de cenários</h2><p>Lucro líquido sempre no eixo financeiro. As curvas se atualizam quando você altera as premissas.</p></div><span className="sim-profit-badge"><LockKeyhole size={12} />Lucro líquido fixo</span></div>
    <Card className="sim-profit-card" title={mode === 'accumulated' ? 'Quanto lucro se acumula ao longo do plano?' : 'Quanto lucro cada ciclo entrega?'} subtitle="Compare manter a verba com as diferentes respostas à escala." action={<div className="sim-chart-modes" role="group" aria-label="Período do gráfico de lucro"><button aria-pressed={mode === 'accumulated'} onClick={() => setMode('accumulated')}>Acumulado</button><button aria-pressed={mode === 'cycles'} onClick={() => setMode('cycles')}>Por ciclo</button></div>}>
      <div className="sim-profit-legend" aria-label="Cenários comparados">{model.series.map(s => <div key={s.id}><span style={{ background: s.color }} /><div><strong>{s.label}</strong><small>{money(s.total)} no período</small></div></div>)}</div>
      <ProfitChart series={model.series} labels={model.cycleLabels} dates={model.cycleDates} mode={mode} title={`Lucro líquido ${mode === 'accumulated' ? 'acumulado' : 'por ciclo'}: controle e cenários`} theme={theme} />
      <details className="sim-profit-values"><summary>Ver valores do gráfico de lucro por ciclo</summary><div className="table-scroll"><table><caption className="sr-only">Lucro líquido {mode === 'accumulated' ? 'acumulado' : 'por ciclo'}, em reais</caption><thead><tr><th>Ciclo / fim do período</th>{model.series.map(s => <th key={s.id} className="num">{s.label}</th>)}</tr></thead><tbody>{model.cycleLabels.slice(1).map((label, i) => <tr key={label}><td>{label}<small className="sim-block">{date(model.cycleDates[i + 1])}</small></td>{model.series.map(s => <td key={s.id} className="num">{money(mode === 'accumulated' ? s.points[i + 1].y : s.cycleProfits[i])}</td>)}</tr>)}</tbody></table></div></details>
    </Card>
    <Card className="sim-profit-card" title="Como o investimento muda o lucro líquido?" subtitle="Verbas entre 50% e 150% do seu plano, mantendo a distribuição entre ciclos e as demais premissas.">
      {model.budgetSeries.length ? <>
        <p className="sim-profit-marker"><i />O ponto azul destaca seu plano atual: <strong>{money(model.current.x)} de investimento → {money(model.current.y)} de lucro líquido.</strong></p>
        <ProfitChart series={model.budgetSeries} labels={model.cycleLabels} mode="budget" title="Investimento Google Ads versus lucro líquido dos cenários" current={model.current} theme={theme} />
        <details className="sim-profit-values"><summary>Ver valores do gráfico de investimento e lucro</summary><div className="table-scroll"><table><caption className="sr-only">Investimento e lucro líquido projetado, em reais</caption><thead><tr><th className="num">Investimento Ads</th>{model.budgetSeries.filter(s => s.id !== 'control').map(s => <th className="num" key={s.id}>{s.label}</th>)}</tr></thead><tbody>{model.budgetSeries.find(s => s.id === 'current')!.points.map((point, i) => <tr key={i}><td className="num">{money(point.x)}</td>{model.budgetSeries.filter(s => s.id !== 'control').map(s => <td className="num" key={s.id}>{money(s.points[i].y)}</td>)}</tr>)}</tbody></table></div></details>
      </> : <p className="sim-help">Defina um investimento positivo em pelo menos um ciclo para explorar a curva de verba.</p>}
    </Card>
    <p className="sim-profit-method">Lucro líquido exibido = lucro do painel − taxas extras − custos fixos informados. O controle mantém as premissas atuais; as outras linhas variam a elasticidade das vendas. As hipóteses não representam probabilidades. ε indica a elasticidade.</p>
  </section></>;
}
