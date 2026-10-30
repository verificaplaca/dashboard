import { useRef, useState } from 'react';
import { Dialog, PageHeading } from './components';
import OriginalOverview, { type OverviewEvent } from './OriginalOverview';
import { exportFile } from './Dashboard';
export default function LiveCampaigns() {
  const page = useRef<HTMLDivElement>(null);
  const [table, setTable] = useState<OverviewEvent['chartTable'] | null>(null);
  const handleEvent = (event: OverviewEvent) => {
    if (event.chartTable) setTable(event.chartTable);
    const root = page.current?.querySelector('.native-overview-host')?.shadowRoot;
    if (!event.ready || !root) return;
    const campaigns = root.getElementById('view-campanhas');
    if (!campaigns || campaigns.dataset.presentation === 'premium') return;
    campaigns.dataset.presentation = 'premium';
    const headings: Record<string, [string, string]> = {
      'Detalhamento por Campanha': ['Performance por campanha', 'Compare investimento, retorno e custo de aquisição. Selecione um título da tabela para ordenar.'],
      'Campanha do Teste': ['Experimento · Pacote Completo', 'O teste de R$ 79,90 usa uma referência própria de CAC e permanece separado das campanhas do produto base.'],
      'Impression Share Perdido': ['Oportunidades nos leilões Search', 'Veja as impressões perdidas por orçamento e classificação. Conversões e receita perdidas são estimativas, não resultados realizados.'],
    };
    campaigns.querySelectorAll<HTMLElement>('.section-label').forEach(heading => {
      const entry = Object.entries(headings).find(([prefix]) => heading.textContent?.trim().startsWith(prefix));
      if (!entry) return;
      heading.textContent = entry[1][0]; heading.setAttribute('role', 'heading'); heading.setAttribute('aria-level', '2');
      const note = document.createElement('p'); note.className = 'campaign-section-note'; note.textContent = entry[1][1]; heading.after(note);
    });
    const titles: Record<string, string> = {
      'Lost Impressions por Campanha': 'Diagnóstico por campanha',
      'Lost Impressions — Causa': 'Onde perdemos visibilidade',
      'Share de Investimento — Campanhas Search': 'Distribuição do investimento',
      'Evolução do Lost Impression Share': 'Impressões perdidas ao longo do tempo',
    };
    campaigns.querySelectorAll('.card-title').forEach(title => { const copy = titles[title.textContent?.trim() || '']; if (copy) title.textContent = copy; });
    const headers: Record<string, string> = { 'Lost IS Budget': 'Perda por orçamento', 'Lost IS Rank': 'Perda por classificação', 'Search Impr. Share': 'Participação nas impressões', 'Conv. Perdidas (aprox)': 'Conversões perdidas (est.)', 'R$ Perdido (aprox)': 'Receita perdida (est.)' };
    campaigns.querySelectorAll('#lostIsTable th').forEach(th => { const copy = headers[th.textContent?.trim() || '']; if (copy) th.textContent = copy; });
    campaigns.querySelectorAll<HTMLElement>('#campaignTable th').forEach(th => {
      th.tabIndex = 0; th.setAttribute('aria-label', `Ordenar por ${th.textContent?.trim()}`);
      th.onkeydown = event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); th.click(); } };
    });
  };
  return <div ref={page} className="campaigns-page"><PageHeading eyebrow="AQUISIÇÃO E RESULTADOS" title="Campanhas" description="Entenda o retorno do investimento e encontre oportunidades de aquisição." /><OriginalOverview view="campanhas" onEvent={handleEvent} />{table && <Dialog title={table.title} close={() => setTable(null)} wide><p className="dialog-description">Valores das mesmas séries do gráfico. “—” indica ausência de valor.</p><div className="table-scroll chart-data-table"><table><thead><tr>{table.headers.map((header, index) => <th key={index}>{header}</th>)}</tr></thead><tbody>{table.rows.map((row, index) => <tr key={index}>{row.map((cell, column) => <td key={column} className={column ? 'num' : ''}>{cell}</td>)}</tr>)}</tbody></table></div><div className="dialog-footer"><button className="button secondary" onClick={() => { const quote = (value: string) => '"' + value.replace(/"/g, '""') + '"'; exportFile('verifica-placa-campanha-grafico.csv', [table.headers, ...table.rows].map(row => row.map(quote).join(';')).join('\n')); }}>Exportar valores</button><button className="button primary" onClick={() => setTable(null)}>Fechar tabela</button></div></Dialog>}</div>;
}
