import { useEffect, useRef, useState } from 'react';
import { CalendarDays, Download, ExternalLink, Info, ShieldCheck } from 'lucide-react';
import { Dialog, Empty, PageHeading } from './components';
import OriginalOverview, { type OverviewEvent, type OverviewHandle } from './OriginalOverview';
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
  const [rules, setRules] = useState(false);
  const [ready, setReady] = useState(false);
  const [period, setPeriod] = useState('Período');
  const [sections, setSections] = useState<Record<string, number>>({});
  const [activeSection, setActiveSection] = useState('financial');
  const [chartTable, setChartTable] = useState<{ title: string; headers: string[]; rows: string[][] } | null>(null);
  const handleOverview = (event: OverviewEvent) => {
    if (typeof event.period === 'string') setPeriod(event.period || 'Período');
    if (event.sections) setSections(event.sections);
    if (event.chartTable) setChartTable(event.chartTable);
    if (typeof event.ready === 'boolean') setReady(event.ready);
    if (event.csv) { exportFile('verifica-placa-visao-geral.csv', event.csv); notify('Os 15 indicadores e comparativos foram exportados.'); }
  };
  useEffect(() => {
    let pending = false;
    const update = () => {
      pending = false;
      const offset = document.querySelector('.overview-navigation')?.getBoundingClientRect().bottom || 128;
      const next = ['financial', 'products', 'operation', 'performance'].filter(id => Number.isFinite(sections[id]) && sections[id] - window.scrollY <= offset + 40).at(-1);
      setActiveSection(next || 'financial');
    };
    const onScroll = () => { if (!pending) { pending = true; requestAnimationFrame(update); } };
    window.addEventListener('scroll', onScroll, { passive: true }); update();
    return () => window.removeEventListener('scroll', onScroll);
  }, [sections]);
  return <div className="original-overview">
    <PageHeading eyebrow="INTELIGÊNCIA DA OPERAÇÃO" title="Visão Geral" description="Vendas, custos e performance. Todos os indicadores do dashboard atual, em uma leitura mais clara.">
      <button className="button secondary" onClick={() => setRules(true)}><Info size={15} />Regras e comparativos</button>
      <button className="button secondary" disabled={!ready} onClick={() => overview.current?.exportIndicators()}><Download size={15} />Exportar</button>
    </PageHeading>
    <div className="overview-navigation"><nav aria-label="Seções da visão geral">{[['financial', 'Financeiro'], ['products', 'Produtos'], ['operation', 'Operação'], ['performance', 'Performance']].map(([id, label]) => <button key={id} aria-current={activeSection === id ? 'location' : undefined} onClick={() => overview.current?.scrollToSection(id)}>{label}</button>)}</nav><button className="overview-period" onClick={() => overview.current?.scrollToSection('filters')}><CalendarDays size={14} />{period}</button><a href="https://verificaplaca.github.io/dashboard/dashboard.html" target="_blank" rel="noreferrer">Dashboard original<ExternalLink size={12} /></a></div>
    <OriginalOverview ref={overview} onEvent={handleOverview} />
    <p className="overview-source-note"><ShieldCheck size={14} />Leitura das fontes originais. As metas, impostos e comparativos seguem as regras do dashboard atual.</p>
    {chartTable && <Dialog title={chartTable.title} close={() => setChartTable(null)} wide><p className="dialog-description">As mesmas séries do gráfico em formato de tabela. “—” indica ausência de valor; projeções e metas mantêm suas próprias colunas.</p><div className="table-scroll chart-data-table"><table><caption className="sr-only">Valores de {chartTable.title}</caption><thead><tr>{chartTable.headers.map((header, index) => <th key={index}>{header}</th>)}</tr></thead><tbody>{chartTable.rows.map((row, index) => <tr key={index}>{row.map((cell, column) => <td key={column} className={column ? 'num' : ''}>{cell}</td>)}</tr>)}</tbody></table></div><div className="dialog-footer"><button className="button secondary" onClick={() => { const quote = (value: string) => '"' + String(value).replace(/"/g, '""') + '"'; exportFile('verifica-placa-grafico.csv', [chartTable.headers, ...chartTable.rows].map(row => row.map(quote).join(';')).join('\n')); }}>Exportar valores</button><button className="button primary" onClick={() => setChartTable(null)}>Fechar tabela</button></div></Dialog>}
    {rules && <Dialog title="Regras da visão geral" close={() => setRules(false)}><div className="overview-rules">
      <p>Esta visão utiliza as mesmas fontes, filtros e cálculos do dashboard original. O período anterior corresponde à janela de calendário imediatamente anterior, com a mesma duração.</p>
      <dl><dt>Custos e resultado</dt><dd>COGS = Google Ads + bureau. Lucro bruto = receita − COGS. Lucro líquido aplica o imposto vigente em cada mês. ROAS = receita ÷ COGS.</dd><dt>Aquisição e conversão</dt><dd>CAC = Google Ads ÷ pedidos pagos. CAC Google Ads usa as conversões reportadas pelo Google. CPV = COGS ÷ pedidos. Conversão = pedidos ÷ checkouts iniciados.</dd><dt>Metas e histórico</dt><dd>As metas e o imposto são carregados por mês, com a mesma herança do original. EBITDA usa todo o histórico disponível e não muda com o filtro de datas.</dd><dt>Escopo dos blocos</dt><dd>O resumo financeiro mantém o consolidado, inclusive ao selecionar uma campanha, como no original. O saldo usa a conta de mídia; bureau operacional usa o mês atual. Dias da semana exibem médias, de segunda a domingo.</dd><dt>Disponibilidade dos dados</dt><dd>Atualização e avisos de dados de exemplo aparecem na própria visão. Caso uma fonte fique indisponível, a sinalização e o comportamento seguem o dashboard original. Os módulos operacionais em integração estão disponíveis na demonstração.</dd></dl>
    </div><div className="dialog-footer"><button className="button primary" onClick={() => setRules(false)}>Entendi</button></div></Dialog>}
  </div>;
}
