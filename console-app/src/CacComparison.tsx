import { useEffect, useRef, useState } from 'react';
import type { Chart } from 'chart.js';
import { LockKeyhole } from 'lucide-react';
import type { CacComparisonRow, CacProjection } from './domain/cac-comparison';
import type { OptimizationSettings, WindowOptimization } from './domain/investment-optimizer';
import { addDays } from './domain/simulation';
const money = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const signed = (n: number) => `${n > 0 ? '+' : n < 0 ? '−' : ''}${money(Math.abs(n))}`;
const date = (s: string) => s.split('-').reverse().join('/');
const series = [
  { key: 'constant', label: 'CAC constante · sem aprendizado', color: '#18b892' },
  { key: 'constantLearning', label: 'CAC após transição · com aprendizado', color: '#83d8c0' },
  { key: 'elastic', label: 'Elasticidade 0,8 · sem aprendizado', color: '#4c83ff' },
  { key: 'elasticLearning', label: 'Elasticidade 0,8 · com aprendizado', color: '#d99536' },
] as const;
function ComparisonGraph({ rows, gain, days, theme }: { rows: CacComparisonRow[]; gain: boolean; days: number; theme: string }) {
  const canvas = useRef<HTMLCanvasElement>(null), [failed, setFailed] = useState(false);
  useEffect(() => {
    let disposed = false, chart: Chart<'bar'> | undefined;
    setFailed(false);
    void import('chart.js/auto').then(({ default: ChartJS }) => {
      if (disposed || !canvas.current) return;
      const css = getComputedStyle(document.documentElement), text = css.getPropertyValue('--muted').trim(), border = css.getPropertyValue('--border').trim();
      chart = new ChartJS(canvas.current, { type: 'bar', data: { labels: rows.map(r => `${money(r.budget)}/dia`), datasets: series.map(s => ({ label: s.label, data: rows.map(r => gain ? r[s.key].delta.profit : r[s.key].scenario.profit), backgroundColor: s.color, borderColor: s.color, borderRadius: 3, maxBarThickness: 28 })) }, options: {
        responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: item => { const r = rows[item.dataIndex][series[item.datasetIndex].key]; return `${series[item.datasetIndex].label}: ${money(r.scenario.profit)} de lucro total; ${signed(r.delta.profit)} vs. manter`; } } } },
        scales: { x: { grid: { display: false }, ticks: { color: text, font: { size: 10 }, callback: (_, index) => `R$ ${rows[index].budget.toLocaleString('pt-BR')}` }, title: { display: true, text: 'Investimento diário', color: text }, border: { color: border } }, y: { beginAtZero: true, ticks: { color: text, maxTicksLimit: 6, callback: n => `R$ ${Number(n).toLocaleString('pt-BR', { notation: 'compact', maximumFractionDigits: 1 })}` }, title: { display: true, text: `${gain ? 'Ganho sobre manter' : 'Lucro líquido total'} · ${days} dias`, color: text }, grid: { color: c => c.tick.value === 0 ? text : border }, border: { display: false } } },
      } });
    }).catch(() => { if (!disposed) setFailed(true); });
    return () => { disposed = true; chart?.destroy(); };
  }, [rows, gain, days, theme]);
  return <div className="sim-cac-canvas">{failed && <p role="status">Gráfico indisponível; todos os valores estão na tabela abaixo.</p>}<canvas ref={canvas} role="img" aria-label={`${gain ? 'Ganho de lucro líquido sobre manter o investimento' : 'Lucro líquido total'}: CAC constante versus elasticidade 0,8, R$ 3.400, R$ 4.100 e R$ 5.000 por dia, com e sem aprendizado`}>Valores exatos na tabela comparativa.</canvas></div>;
}
function ProfitCell({ result, winner }: { result: CacProjection; winner: boolean }) {
  return <td className={`num ${winner ? 'sim-cac-winner' : ''}`}><strong>{money(result.scenario.profit)}</strong><small className={`sim-block ${result.delta.profit >= 0 ? 'sim-good' : 'sim-bad'}`}>{signed(result.delta.profit)} vs. manter</small>{winner && <small className="sim-block">Maior lucro entre as três verbas</small>}</td>;
}
export default function CacComparison({ rows, recent, settings, theme }: { rows: CacComparisonRow[]; recent: WindowOptimization; settings: OptimizationSettings; theme: string }) {
  const best = (key: typeof series[number]['key']) => Math.max(...rows.map(r => r[key].scenario.profit));
  return <section className="sim-cac-comparison" aria-label="Comparação CAC constante e elasticidade 0,8">
    <div className="sim-growth-heading"><div><h3>CAC constante × elasticidade 0,8</h3><p>Base da recuperação recente: {date(recent.start)}–{date(recent.end)} · {recent.base.days} dia(s) completo(s) · CAC interno observado: <b>{money(recent.base.ads / recent.base.orders)}</b></p></div><span className="sim-profit-badge"><LockKeyhole size={12} />Lucro líquido</span></div>
    <p className="sim-help">As três verbas são comparadas nas quatro condições, independentemente da base e da elasticidade destacadas no otimizador. Referência comum de investimento: <b>{money(settings.controlBudget)}/dia</b> durante {settings.days} dias, de {date(settings.start)} a {date(addDays(settings.start, settings.days - 1))}. A referência inicial usa gasto efetivamente observado, não orçamento configurado; você pode alterá-la acima.</p>
    <div className="sim-cac-controls"><div><span>Lucro ao manter · CAC constante</span><strong>{money(rows[0].constant.control.profit)}</strong></div><div><span>Lucro ao manter · elasticidade 0,8</span><strong>{money(rows[0].elastic.control.profit)}</strong></div><p>Ganhos sempre contra manter a mesma verba na respectiva hipótese. Manter não sofre novo aprendizado.</p></div>
    <div className="sim-cac-legend" aria-label="Condições comparadas">{series.map(s => <span key={s.key}><i style={{ background: s.color }} />{s.label}</span>)}</div>
    <div className="sim-cac-graphs"><div><h4>Lucro líquido total</h4><ComparisonGraph rows={rows} gain={false} days={settings.days} theme={theme} /></div><div><h4>Ganho sobre manter o investimento</h4><ComparisonGraph rows={rows} gain days={settings.days} theme={theme} /></div></div>
    <div className="table-scroll"><table><caption className="sr-only">Doze cenários: lucro líquido total e ganho sobre manter nas três verbas, com e sem aprendizado, por hipótese de CAC</caption><thead><tr><th rowSpan={2}>Verba por dia</th><th colSpan={2} className="num">CAC recente preservado após eventual transição</th><th colSpan={2} className="num">Elasticidade 0,8</th></tr><tr><th className="num">Sem aprendizado</th><th className="num">Com aprendizado</th><th className="num">Sem aprendizado</th><th className="num">Com aprendizado</th></tr></thead><tbody>{rows.map(r => <tr key={r.budget}><td><strong>{money(r.budget)}</strong><small className="sim-block">{money(r.constant.scenario.ads)} de mídia em {settings.days} dias</small></td>{series.map(s => <ProfitCell key={s.key} result={r[s.key]} winner={r[s.key].scenario.profit === best(s.key)} />)}</tr>)}</tbody></table></div>
    <p className="sim-help">CAC constante usa elasticidade 1 e 100% dos pedidos sensíveis à mídia; antes de uma nova perda inicial, Ads ÷ pedidos pagos permanece no nível observado. A hipótese 0,8 usa os mesmos pedidos sensíveis, ticket, bureau, imposto e custos extras. Nesta seção, as elasticidades 1 e 0,8 são fixas para preservar a comparação, mesmo se você editar a elasticidade do otimizador.</p>
    <p className="sim-help">Com aprendizado: {settings.learningPct}% menos pedidos nos primeiros {Math.min(settings.learningDays, settings.days)} dia(s) após mudar a verba. O CAC sobe temporariamente nesse intervalo; só depois retorna à resposta projetada. A perda é uma hipótese editável nos controles acima e já entra na receita e nos custos, sem desconto em duplicidade. Base recente curta: manter esse CAC ao escalar ainda precisa ser validado. O destaque indica o melhor dos três investimentos, não um ótimo global.</p>
  </section>;
}
