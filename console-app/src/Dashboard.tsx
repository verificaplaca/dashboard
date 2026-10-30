import { useRef, useState } from 'react';
import { ShieldCheck } from 'lucide-react';
import { Dialog, Empty } from './components';
import OriginalOverview, { type OverviewEvent, type OverviewHandle, type DayDetail } from './OriginalOverview';
import type { MetricDetail } from './domain/overview-insights';
import { money } from './format';
import { useUI } from './ui';

export const teamNames: Record<string, string> = { 'team-sales': 'Comercial', 'team-support': 'Atendimento', 'team-growth': 'Crescimento' };
export function exportFile(name: string, text: string, mime = 'text/csv;charset=utf-8') { const blob = new Blob(['\uFEFF', text], { type: mime }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url); }
export function TrendChart({ values, secondary, labels }: { values: number[]; secondary?: number[]; labels: string[] }) {
  if (!values.some(Boolean) && !secondary?.some(Boolean)) return <Empty title="Sem dados neste período" description="Selecione um período com pagamentos confirmados." />;
  const max = Math.max(...values, ...(secondary || []), 100), min = Math.min(0, ...(secondary || []));
  const point = (v: number, i: number) => `${64 + i * (710 / Math.max(1, values.length - 1))},${220 - (v - min) / (max - min) * 180}`;
  const points = values.map(point).join(' '), second = secondary?.map(point).join(' ');
  return <div className="trend-chart"><svg viewBox="0 0 820 260" role="img" aria-label={`Evolução diária. ${labels.map((l, i) => `${l}: ${money(values[i])}`).join('; ')}`}><defs><linearGradient id="chart-fill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#67e8f9" stopOpacity=".22" /><stop offset="100%" stopColor="#67e8f9" stopOpacity="0" /></linearGradient></defs>{[0, 1, 2, 3].map(i => <g key={i}><line x1="64" x2="776" y1={40 + i * 60} y2={40 + i * 60} className="chart-grid" /><text x="0" y={44 + i * 60} className="chart-label">{money(Math.round(max - i / 3 * (max - min)))}</text></g>)}<polygon points={`64,220 ${points} 774,220`} fill="url(#chart-fill)" /><polyline points={points} fill="none" stroke="var(--cyan)" strokeWidth="3" strokeLinejoin="round" />{second && <polyline points={second} fill="none" stroke="#a78bfa" strokeWidth="2" strokeDasharray="5 5" strokeLinejoin="round" />}{values.map((v, i) => <g key={i}><circle cx={point(v, i).split(',')[0]} cy={point(v, i).split(',')[1]} r="4" fill="var(--cyan)"><title>{labels[i]}: {money(v)}</title></circle><text x={64 + i * (710 / Math.max(1, values.length - 1))} y="249" textAnchor="middle" className="chart-label">{labels[i]}</text></g>)}</svg></div>;
}

export default function Dashboard() {
  const { notify } = useUI();
  const overview = useRef<OverviewHandle>(null);
  const [metric, setMetric] = useState<MetricDetail | null>(null);
  const [day, setDay] = useState<DayDetail | null>(null);
  const [rules, setRules] = useState(false);
  const [chartTable, setChartTable] = useState<{ title: string; headers: string[]; rows: string[][] } | null>(null);
  const handleOverview = (event: OverviewEvent) => {
    if (event.ready) window.dispatchEvent(new CustomEvent('vp-overview-source', { detail: event.example ? 'EXEMPLO · FONTE INDISPONÍVEL' : 'DADOS REAIS' }));
    if (event.action === 'rules') setRules(true);
    if (event.action === 'export') overview.current?.exportIndicators();
    if (event.metricDetail) setMetric(event.metricDetail);
    if (event.dayDetail) setDay(event.dayDetail);
    if (event.chartTable) setChartTable(event.chartTable);
    if (event.csv) { exportFile('verifica-placa-visao-geral.csv', event.csv); notify('Os 15 indicadores e comparativos foram exportados.'); }
  };
  return <div className="original-overview">
    <OriginalOverview ref={overview} onEvent={handleOverview} />
    <p className="overview-source-note"><ShieldCheck size={14} />Leitura das fontes originais. As metas, impostos e comparativos seguem as regras do dashboard atual.</p>
    {chartTable && <Dialog title={chartTable.title} close={() => setChartTable(null)} wide><p className="dialog-description">As mesmas séries do gráfico em formato de tabela. “—” indica ausência de valor; projeções e metas mantêm suas próprias colunas.</p><div className="table-scroll chart-data-table"><table><caption className="sr-only">Valores de {chartTable.title}</caption><thead><tr>{chartTable.headers.map((header, index) => <th key={index}>{header}</th>)}</tr></thead><tbody>{chartTable.rows.map((row, index) => <tr key={index}>{row.map((cell, column) => <td key={column} className={column ? 'num' : ''}>{cell}</td>)}</tr>)}</tbody></table></div><div className="dialog-footer"><button className="button secondary" onClick={() => { const quote = (value: string) => '"' + String(value).replace(/"/g, '""') + '"'; exportFile('verifica-placa-grafico.csv', [chartTable.headers, ...chartTable.rows].map(row => row.map(quote).join(';')).join('\n')); }}>Exportar valores</button><button className="button primary" onClick={() => setChartTable(null)}>Fechar tabela</button></div></Dialog>}
    {metric && <Dialog title={metric.title} close={() => setMetric(null)} side>
      <p className="detail-eyebrow">ENTENDA O INDICADOR</p>
      <div className={`detail-total ${metric.tone === 'negative' ? 'loss' : metric.tone === 'positive' ? 'gain' : ''}`}>{metric.value}</div>
      <p className="detail-formula">{metric.formula}</p>
      <dl className="detail-ledger">{metric.rows.map(row => <div key={row.label}><dt>{row.label}</dt><dd>{row.value}</dd></div>)}</dl>
      <div className="detail-comparison"><p><strong>Comparação</strong><br />{metric.comparison}</p><p><strong>Meta ou referência</strong><br />{metric.target}</p></div>
      <p className="detail-explanation">{metric.context}</p>
      <a href="#/metas" className="button primary detail-link" onClick={() => setMetric(null)}>Conferir metas ↗</a>
    </Dialog>}
    {day && <Dialog title="Detalhe do resultado" close={() => { setDay(null); overview.current?.clearSelection(); }} side>
      <p className="detail-eyebrow">{day.example ? 'DADOS DE EXEMPLO' : 'DADOS REAIS'} · {day.date.split('-').reverse().join('/')}</p>
      <h3>Lucro líquido do dia</h3><div className={`detail-total ${day.result < 0 ? 'loss' : 'gain'}`}>{money(Math.round(day.result * 100))}</div>
      <span className={`detail-status ${day.result < 0 ? 'loss' : 'gain'}`}>{day.result < 0 ? 'Resultado negativo' : day.result > 0 ? 'Resultado positivo' : 'Resultado neutro'}</span>
      <h3 className="detail-section-title">Como o resultado é formado</h3>
      <dl className="detail-ledger">{[['Receita', day.revenue], ['Imposto', -day.tax], ['Google Ads', -day.ads], ['Bureau', -day.bureau], ['Resultado', day.result]].map(([label, value]) => <div key={String(label)}><dt>{label}</dt><dd>{money(Math.round(Number(value) * 100))}</dd></div>)}</dl>
      <p className="detail-explanation">Resultado após imposto, Ads e bureau. Despesas complementares da DRE não incluídas. Os valores acima correspondem ao dia selecionado; metas mensais são avaliadas no resumo do período.</p>
      <a className="button primary detail-link" href="#/financeiro" onClick={() => setDay(null)}>Abrir financeiro ↗</a>
    </Dialog>}
    {rules && <Dialog title="Regras da visão geral" close={() => setRules(false)}><div className="overview-rules">
      <p>Esta visão utiliza as mesmas fontes, filtros e cálculos do dashboard original. O período anterior corresponde à janela de calendário imediatamente anterior, com a mesma duração.</p>
      <dl><dt>Custos e resultado</dt><dd>COGS = Google Ads + bureau. Lucro bruto = receita − COGS. Lucro líquido aplica o imposto vigente em cada mês. ROAS = receita ÷ COGS.</dd><dt>Aquisição e conversão</dt><dd>CAC = Google Ads ÷ pedidos pagos. CAC Google Ads usa as conversões reportadas pelo Google. CPV = COGS ÷ pedidos. Conversão = pedidos ÷ checkouts iniciados.</dd><dt>Metas e histórico</dt><dd>As metas e o imposto são carregados por mês, com a mesma herança do original. EBITDA usa todo o histórico disponível e não muda com o filtro de datas.</dd><dt>Escopo dos blocos</dt><dd>O resumo financeiro mantém o consolidado, inclusive ao selecionar uma campanha, como no original. O saldo usa a conta de mídia; bureau operacional usa o mês atual. Dias da semana exibem médias, de segunda a domingo.</dd><dt>Disponibilidade dos dados</dt><dd>O rodapé e as exportações identificam a fonte e a atualização dos dados, inclusive quando uma consulta fica indisponível e o dashboard utiliza dados de exemplo. Os módulos operacionais em integração estão disponíveis na demonstração.</dd></dl>
    </div><div className="dialog-footer"><button className="button primary" onClick={() => setRules(false)}>Entendi</button></div></Dialog>}
  </div>;
}
