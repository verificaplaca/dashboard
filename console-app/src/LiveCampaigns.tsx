import { useState } from 'react';
import { Dialog, PageHeading } from './components';
import OriginalOverview, { type OverviewEvent } from './OriginalOverview';
import { exportFile } from './Dashboard';
export default function LiveCampaigns() {
  const [table, setTable] = useState<OverviewEvent['chartTable'] | null>(null);
  return <><PageHeading eyebrow="AQUISIÇÃO E RESULTADOS" title="Campanhas" description="Investimento, conversões e participação de impressão, com os dados e regras do dashboard atual." /><OriginalOverview view="campanhas" onEvent={event => { if (event.chartTable) setTable(event.chartTable); }} />{table && <Dialog title={table.title} close={() => setTable(null)} wide><p className="dialog-description">Valores das mesmas séries do gráfico. “—” indica ausência de valor.</p><div className="table-scroll chart-data-table"><table><thead><tr>{table.headers.map((header, index) => <th key={index}>{header}</th>)}</tr></thead><tbody>{table.rows.map((row, index) => <tr key={index}>{row.map((cell, column) => <td key={column} className={column ? 'num' : ''}>{cell}</td>)}</tr>)}</tbody></table></div><div className="dialog-footer"><button className="button secondary" onClick={() => { const quote = (value: string) => '"' + value.replace(/"/g, '""') + '"'; exportFile('verifica-placa-campanha-grafico.csv', [table.headers, ...table.rows].map(row => row.map(quote).join(';')).join('\n')); }}>Exportar valores</button><button className="button primary" onClick={() => setTable(null)}>Fechar tabela</button></div></Dialog>}</>;
}
