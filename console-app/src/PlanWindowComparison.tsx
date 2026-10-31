import { mobileChartTooltips } from './chart-tooltips';
import { balancedChartScales } from './domain/chart-scales';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { Chart } from 'chart.js';
import { Download } from 'lucide-react';
import { exportFile } from './Dashboard';
import { useUI } from './ui';
import { bestPlanStrategies, comparePlanWindows, type GradualComparison, type GradualLearning, type PlanStrategy } from './domain/gradual-comparison';
import type { Snapshot } from './domain/simulation';
const brl = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const signed = (n: number) => `${n > 0 ? '+' : n < 0 ? '−' : ''}${brl(Math.abs(n))}`;
const day = (s: string) => s.split('-').reverse().join('/');
const names: Record<PlanStrategy, string> = { keep: 'Manter R$ 2.790/dia', gradual: 'Plano gradual', direct: 'Direto a R$ 5.000/dia' };
type ChartRow = { label: string; row: GradualComparison['rows'][number] };
function WindowChart({ rows, gain, theme }: { rows: ChartRow[]; gain: boolean; theme: string }) {
  const canvas = useRef<HTMLCanvasElement>(null), [failed, setFailed] = useState(false);
  useEffect(() => {
    let disposed = false, chart: Chart<'bar'> | undefined;
    setFailed(false);
    void import('chart.js/auto').then(({ default: ChartJS }) => {
      if (disposed || !canvas.current || !rows.length) return;
      const css = getComputedStyle(document.documentElement), muted = css.getPropertyValue('--muted').trim(), border = css.getPropertyValue('--border').trim();
      chart = new ChartJS(canvas.current, { plugins: [balancedChartScales, mobileChartTooltips], type: 'bar', data: { labels: rows.map(r => r.label), datasets: [
        { label: 'Plano gradual', data: rows.map(r => gain ? r.row.result.delta : r.row.result.scenario.profit), backgroundColor: '#4c83ff', borderRadius: 3, maxBarThickness: 44 },
        { label: 'Direto a R$ 5.000/dia', data: rows.map(r => gain ? r.row.directResult.delta : r.row.directResult.scenario.profit), backgroundColor: '#18b892', borderRadius: 3, maxBarThickness: 44 },
        ...(!gain ? [{ label: 'Manter R$ 2.790/dia', data: rows.map(r => r.row.result.control.profit), backgroundColor: theme === 'dark' ? '#68758a' : '#a0aabe', borderRadius: 3, maxBarThickness: 30 }] : []),
      ] }, options: { responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: item => `${item.dataset.label}: ${brl(Number(item.raw))}`, afterBody: items => `Melhor resultado: ${bestPlanStrategies(rows[items[0].dataIndex].row).map(w => names[w]).join(' / ')}` } } }, scales: {
        x: { ticks: { color: muted, font: { size: 11 }, maxRotation: 0, autoSkip: false }, grid: { display: false } },
        y: { beginAtZero: true, title: { display: true, text: gain ? 'Ganho de lucro líquido sobre manter · R$' : 'Lucro líquido até 31/01/2027 · R$', color: muted }, ticks: { color: muted, maxTicksLimit: 6, callback: n => `R$ ${Number(n).toLocaleString('pt-BR', { notation: 'compact', maximumFractionDigits: 1 })}` }, grid: { color: c => c.tick.value === 0 ? muted : border }, border: { display: false } },
      } } });
    }).catch(() => { if (!disposed) setFailed(true); });
    return () => { disposed = true; chart?.destroy(); };
  }, [rows, gain, theme]);
  return <div className="sim-window-canvas">{failed && <p role="status">Gráfico indisponível. Os cartões e a tabela contêm os valores completos.</p>}<canvas ref={canvas} role="img" aria-label={`${gain ? 'Ganho sobre manter' : 'Lucro líquido total'}: ${rows.map(({ label, row }) => `${label}: gradual ${brl(gain ? row.result.delta : row.result.scenario.profit)}, direto ${brl(gain ? row.directResult.delta : row.directResult.scenario.profit)}, manter ${brl(gain ? 0 : row.result.control.profit)}`).join('; ')}`}>Valores na tabela abaixo.</canvas></div>;
}
export default function PlanWindowComparison({ snapshot, learning, loading, sourceError }: { snapshot: Snapshot | null; learning: GradualLearning; loading: boolean; sourceError: string }) {
  const { theme } = useUI();
  const [elasticity, setElasticity] = useState(.8), [withLearning, setWithLearning] = useState(true), [gain, setGain] = useState(true);
  const windows = useMemo(() => snapshot && !loading && !sourceError ? comparePlanWindows(snapshot, learning) : [], [snapshot, learning.days, learning.lossPct, loading, sourceError]);
  const selected = useMemo(() => windows.map(w => ({ ...w, row: w.data?.rows.find(r => r.elasticity === elasticity && r.withLearning === withLearning) })), [windows, elasticity, withLearning]);
  const chartRows = useMemo(() => selected.flatMap(w => w.row ? [{ label: w.label, row: w.row }] : []), [selected]);
  const keepCount = selected.filter(w => w.row && bestPlanStrategies(w.row).includes('keep')).length;
  return <section className="sim-window-comparison" aria-labelledby="sim-window-title">
    <h3 id="sim-window-title">A conclusão se mantém nas diferentes bases?</h3>
    <p className="sim-help">Mesmas estratégias de 26/10/2026 a 31/01/2027 (98 dias). Só muda a eficiência observada na base: CAC, ticket e bureau reais. O dia 16/10 fica fora da base após as mudanças.</p>
    <div className="sim-window-controls"><label className="sim-field"><span>Resposta à verba</span><select value={elasticity} onChange={e => setElasticity(Number(e.target.value))}><option value={.8}>Elasticidade 0,8</option><option value={1}>CAC constante</option></select></label><label className="sim-field"><span>Condição de aprendizado</span><select value={withLearning ? 'yes' : 'no'} onChange={e => setWithLearning(e.target.value === 'yes')}><option value="yes">Com aprendizado</option><option value="no">Sem aprendizado</option></select></label></div>
    {snapshot && !loading && !sourceError && <p className="sim-gradual-source">Fontes até <strong>{day(snapshot.end)}</strong> · leitura em {new Date(snapshot.loadedAt).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' })} (Brasília). {withLearning ? `${Number.isFinite(learning.days) ? learning.days : '—'} dias / ${Number.isFinite(learning.lossPct) ? learning.lossPct : '—'}% menos pedidos por aumento.` : 'Sem nova perda hipotética de pedidos.'}</p>}
    {(loading || sourceError || !snapshot) && <p className="sim-alert" role="status">{loading ? 'Atualizando as quatro bases reais…' : sourceError || 'Aguardando as fontes reais.'}</p>}
    {!!chartRows.length && <div className="sim-comparison-answer" role="status"><strong>{keepCount ? `Manter R$ 2.790/dia lidera ou empata em ${keepCount} de ${chartRows.length} bases disponíveis.` : `Escalar projeta mais lucro que manter nas ${chartRows.length} bases disponíveis.`}</strong><p>{chartRows.length < 4 ? 'Há bases indisponíveis; a comparação está incompleta. ' : ''}Válido para {elasticity === .8 ? 'elasticidade 0,8' : 'CAC constante'}, {withLearning ? 'com' : 'sem'} aprendizado. A liberação do próximo teste depende dos critérios realizados acima, mesmo quando a projeção favorece aumentar.</p></div>}
    <div className="sim-window-cards">{selected.map(w => { const winners = w.row ? bestPlanStrategies(w.row) : []; return <article key={w.id} className={winners.includes('keep') ? 'keep-best' : ''}>
      <h4>{w.label}</h4><p>{day(w.start)}–{day(w.end)}{w.data ? ` · ${w.data.base.days} dias` : ''}</p>
      {!w.row || !w.data ? <p className="sim-alert danger">Base indisponível: {w.error}</p> : <>
        <p>CAC real <strong>{brl(w.data.observedCac)}</strong> · lucro real/dia <strong>{brl(w.data.base.profit / w.data.base.days)}</strong></p>
        <div className="sim-window-best">{winners.length > 1 ? 'Empate: ' : 'Maior lucro: '}{winners.map(winner => names[winner]).join(' / ')}</div>
        <dl>{([['keep', w.row.result.control.profit, 0], ['gradual', w.row.result.scenario.profit, w.row.result.delta], ['direct', w.row.directResult.scenario.profit, w.row.directResult.delta]] as const).map(([strategy, profit, delta]) => <div key={strategy} className={winners.includes(strategy) ? 'is-winner' : ''}><dt>{strategy === 'keep' ? 'Manter' : strategy === 'gradual' ? 'Gradual' : 'Direto R$ 5.000'}</dt><dd>{brl(profit)}{strategy !== 'keep' && <small className={delta >= 0 ? 'sim-good' : 'sim-bad'}>{signed(delta)} vs. manter</small>}</dd></div>)}</dl><small>Lucro líquido projetado até 31/01</small>
      </>}
    </article>; })}</div>
    {!!chartRows.length && <><div className="sim-opt-graph-heading"><div className="sim-opt-legend"><span><i style={{ background: '#4c83ff' }} />Gradual</span><span><i style={{ background: '#18b892' }} />Direto R$ 5.000</span>{!gain && <span><i style={{ background: '#68758a' }} />Manter R$ 2.790</span>}</div><div className="sim-chart-modes" role="group" aria-label="Comparação financeira das bases"><button aria-pressed={gain} onClick={() => setGain(true)}>Ganho sobre manter</button><button aria-pressed={!gain} onClick={() => setGain(false)}>Lucro total</button></div></div><p className="sim-help">{gain ? 'Zero = manter R$ 2.790/dia. Barras negativas indicam menos lucro que manter.' : 'As três estratégias na mesma hipótese e horizonte.'}</p><WindowChart rows={chartRows} gain={gain} theme={theme} /></>}
    {!!windows.length && <details className="sim-ux-disclosure"><summary>Ver as quatro condições e os valores exatos de todas as bases</summary><div className="table-scroll"><table><caption className="sr-only">Todas as bases, respostas à verba e hipóteses de aprendizado; lucro líquido até 31/01/2027</caption><thead><tr><th>Base</th><th>Condição</th><th className="num">Lucro · manter</th><th className="num">Lucro · gradual</th><th className="num">Lucro · direto</th><th className="num">Ganho gradual</th><th className="num">Ganho direto</th><th>Melhor resultado</th></tr></thead><tbody>{windows.flatMap(w => w.data ? w.data.rows.map(r => <tr key={`${w.id}-${r.id}`}><td>{w.label}</td><td>{r.label}</td><td className="num">{brl(r.result.control.profit)}</td><td className="num">{brl(r.result.scenario.profit)}</td><td className="num">{brl(r.directResult.scenario.profit)}</td><td className={`num ${r.result.delta >= 0 ? 'sim-good' : 'sim-bad'}`}>{signed(r.result.delta)}</td><td className={`num ${r.directResult.delta >= 0 ? 'sim-good' : 'sim-bad'}`}>{signed(r.directResult.delta)}</td><td>{bestPlanStrategies(r).map(winner => names[winner]).join(' / ')}</td></tr>) : [<tr key={w.id}><td>{w.label}</td><td colSpan={7}>{w.error}</td></tr>])}</tbody></table></div></details>}
    <p className="sim-help">As bases de 14 e 30 dias misturam períodos antes e depois das alterações. Servem como sensibilidade, sem substituir o acompanhamento após 16/10. A hipótese de aprendizado trata cada aumento igualmente: seis transições no gradual e uma no direto; não mede o efeito do tamanho do salto. O controle não recebe nova perda.</p>
    <p className="sim-help">Mídia total: manter R$ 273.420 · gradual R$ 389.940 · direto R$ 490.000. O calendário integral é uma hipótese; reter etapas muda as datas e o resultado. Imposto das metas; sem custos extras não informados.</p>
    <button className="text-button" disabled={!chartRows.length} onClick={() => exportFile('verifica-placa-analise-bases-plano.json', JSON.stringify({ modelVersion: 1, exportedAt: new Date().toISOString(), selected: { elasticity, withLearning }, learning, windows, snapshot, definitions: { comparison: 'mesmo calendário e hipóteses, eficiência real de cada janela', winner: 'maior lucro entre manter 2790, gradual e direto 5000; empates preservados', decision: 'projeção não libera aumento; critérios realizados independentes', missing: 'base inválida sem projeções ou valores zerados' } }, null, 2), 'application/json;charset=utf-8')}><Download size={14} />Exportar análise das quatro bases</button>
  </section>;
}
