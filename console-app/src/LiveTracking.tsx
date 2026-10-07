import { useEffect, useState } from 'react';
import { ChevronLeft, ChevronRight, RefreshCw } from 'lucide-react';
import { Card, Empty, PageHeading } from './components';
import { readPublic } from './api';
type Dispatch = { checkout_id: string; event_id: string; order_nsu: string; value: number; ga4_status: string; ads_status: string; meta_status: string; attempts: number; created_at: string; paid_at: string | null };
const label: Record<string, string> = { success: 'Enviado', failed: 'Falhou', pending: 'Pendente', skipped: 'Não enviado' };
function Channel({ status }: { status: string }) { return <span className={`badge ${status === 'success' ? 'paid' : status === 'failed' ? 'failed' : 'pending'}`}><i />{label[status] || status}</span>; }
export default function LiveTracking() {
  const [rows, setRows] = useState<Dispatch[]>([]), [page, setPage] = useState(0), [filter, setFilter] = useState('all'), [loading, setLoading] = useState(true), [error, setError] = useState(''), [version, setVersion] = useState(0);
  const [from] = useState(() => { const date = new Date(); date.setUTCDate(date.getUTCDate() - 90); return date.toISOString(); });
  useEffect(() => {
    const abort = new AbortController(); setLoading(true); setError('');
    const query = new URLSearchParams({ select: 'checkout_id,event_id,order_nsu,value,ga4_status,ads_status,meta_status,attempts,created_at,paid_at', order: 'created_at.desc,checkout_id.asc', created_at: `gte.${from}`, limit: '51', offset: String(page * 50) });
    if (filter === 'failed') query.set('or', '(ga4_status.eq.failed,ads_status.eq.failed,meta_status.eq.failed)');
    if (filter === 'pending') query.set('or', '(ga4_status.eq.pending,ads_status.eq.pending,meta_status.eq.pending)');
    readPublic<Dispatch>('conversion_dispatches_public', query, abort.signal).then(setRows).catch(reason => { if (!abort.signal.aborted) { setRows([]); setError(reason.message); } }).finally(() => { if (!abort.signal.aborted) setLoading(false); });
    return () => abort.abort();
  }, [page, filter, from, version]);
  return <><PageHeading eyebrow="QUALIDADE DOS EVENTOS" title="Tracking" description="Envios reais para GA4, Google Ads e Meta. Janela de 90 dias, como na origem."><button className="button secondary" disabled={loading} onClick={() => setVersion(version + 1)}><RefreshCw size={15} />Atualizar</button></PageHeading>
    <div className="toolbar"><span>Pedidos e eventos · somente leitura</span><select aria-label="Filtrar envios" value={filter} onChange={event => { setFilter(event.target.value); setPage(0); }}><option value="all">Todos os envios</option><option value="failed">Com falha</option><option value="pending">Pendentes</option></select></div>
    {error && <p role="alert" className="info-note">{error}</p>}
    <Card title="Histórico de envios" subtitle={loading ? 'Atualizando…' : `Página ${page + 1} · ${Math.min(rows.length, 50)} registros`}><div className="table-scroll"><table aria-busy={loading}><thead><tr><th>Pedido / evento</th><th>Data</th><th>Valor</th><th>GA4</th><th>Google Ads</th><th>Meta</th><th>Tentativas</th></tr></thead><tbody>{!loading && rows.slice(0, 50).map(row => <tr key={row.checkout_id}><td><strong>{row.order_nsu || row.checkout_id}</strong><small>{row.event_id}</small></td><td>{new Date(row.created_at).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' })}</td><td>{Number(row.value).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td><td><Channel status={row.ga4_status} /></td><td><Channel status={row.ads_status} /></td><td><Channel status={row.meta_status} /></td><td className="num">{row.attempts}</td></tr>)}</tbody></table></div>{!loading && !rows.length && <Empty title={error ? 'Fonte indisponível' : 'Nenhum envio encontrado'} description={error ? 'Atualize para tentar novamente.' : 'Ajuste o filtro ou aguarde novos eventos.'} />}<div className="live-pager"><button className="button secondary" disabled={page === 0 || loading} onClick={() => setPage(page - 1)}><ChevronLeft size={15} />Anterior</button><span>Página {page + 1}</span><button className="button secondary" disabled={rows.length <= 50 || loading} onClick={() => setPage(page + 1)}>Próxima<ChevronRight size={15} /></button></div></Card><p className="info-note">“Não enviado” corresponde ao status skipped; não equivale a envio confirmado. Reenvios e reconciliação continuam no backend existente.</p></>;
}
