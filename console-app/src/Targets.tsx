import { useEffect, useState } from 'react';
import { CalendarDays, Check, RefreshCw } from 'lucide-react';
import { Card, Dialog, PageHeading } from './components';
import { getAuthClient, readPublic } from './api';
import { useAuth } from './Auth';
import { useUI } from './ui';
import { targetFields, targetPayload, type Target } from './domain/targets';
import { saveTarget } from './domain/save-target';

export default function Targets() {
  const { user, openLogin } = useAuth(), { notify } = useUI();
  const [rows, setRows] = useState<Target[]>([]), [month, setMonth] = useState(() => new Intl.DateTimeFormat('sv-SE', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit' }).format(new Date()));
  const [values, setValues] = useState<Record<string, string>>({}), [loaded, setLoaded] = useState(false), [loading, setLoading] = useState(true), [busy, setBusy] = useState(false), [error, setError] = useState(''), [pending, setPending] = useState<Target | null>(null);
  const load = async (signal?: AbortSignal) => {
    setLoaded(false); setLoading(true); setError('');
    try { const data = await readPublic<Target>('monthly_targets', new URLSearchParams({ select: '*', order: 'month.asc' }), signal); setRows(data); setLoaded(true); }
    catch (reason) { if (!signal?.aborted) setError(reason instanceof Error ? reason.message : 'Falha ao carregar metas.'); }
    finally { if (!signal?.aborted) setLoading(false); }
  };
  useEffect(() => { const abort = new AbortController(); void load(abort.signal); return () => abort.abort(); }, []);
  const current = rows.find(row => row.month === `${month}-01`);
  const inherited = (key: typeof targetFields[number][0]) => rows.filter(row => row.month <= `${month}-01` && row[key] != null).at(-1);
  useEffect(() => { setValues(Object.fromEntries(targetFields.map(([key]) => [key, current?.[key] == null ? '' : String(current[key])]))); }, [month, current]);
  const save = async () => {
    if (!pending || busy) return;
    setBusy(true); setError('');
    try {
      const client = await getAuthClient();
      await saveTarget(client, pending, current);
      setPending(null); await load(); notify('Metas salvas no Supabase. A Visão Geral usará os valores atualizados.');
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Falha ao salvar. Tente novamente.'); }
    finally { setBusy(false); }
  };
  return <>
    <PageHeading eyebrow="GESTÃO DE RESULTADOS" title="Metas mensais" description="Metas e imposto do mesmo banco usado no dashboard atual."><button className="button secondary" disabled={loading || busy} onClick={() => void load()}><RefreshCw size={15} />Atualizar</button>{!user && <button className="button primary" onClick={openLogin}>Entrar para editar</button>}</PageHeading>
    {error && <p role="alert" className="info-note">{error}</p>}
    <Card title="Planejamento do mês" subtitle="Campos vazios herdam o último valor anterior. Sem cadastro, permanecem os padrões do dashboard.">
      <form className="live-form" onSubmit={event => { event.preventDefault(); if (!user) { openLogin(); return; } try { setPending(targetPayload(month, values)); setError(''); } catch (reason) { setError((reason as Error).message); } }}>
        <label className="month-field"><span><CalendarDays size={14} />Mês</span><input type="month" value={month} onChange={event => setMonth(event.target.value)} required disabled={busy} /></label>
        <div className="live-field-grid">{targetFields.map(([key, label, step]) => <label key={key}>{label}<input type="number" step={step} min="0" max={['tax_pct', 'target_upsell_pct'].includes(key) ? 100 : undefined} value={values[key] || ''} placeholder={key === 'tax_pct' ? 'Herança · padrão 8%' : 'Herdar valor anterior'} onChange={event => setValues({ ...values, [key]: event.target.value })} disabled={!user || busy || loading || !loaded} /><small>{loaded ? inherited(key) ? `Em vigor: ${inherited(key)![key]!.toLocaleString('pt-BR')} · ${inherited(key)!.month.slice(0, 7)}` : 'Em vigor: padrão do dashboard' : 'Carregando valor vigente…'}</small></label>)}</div>
        <div className="live-form-footer"><span>{user ? `Edição autenticada · ${user.email}` : 'Você está em modo de leitura.'}</span><button className="button primary" disabled={!user || loading || busy || !loaded}><Check size={15} />Revisar e salvar</button></div>
      </form>
    </Card>
    <Card title="Histórico de metas" subtitle={loading ? 'Carregando…' : `${rows.length} meses cadastrados`}><div className="table-scroll"><table><thead><tr><th>Mês</th>{targetFields.map(([key, label]) => <th key={key}>{label}</th>)}<th /></tr></thead><tbody>{rows.slice().reverse().map(row => <tr key={row.month}><td>{row.month.slice(0, 7)}</td>{targetFields.map(([key]) => <td key={key} className="num">{row[key] == null ? 'Herdado' : row[key]!.toLocaleString('pt-BR')}</td>)}<td><button className="text-button" onClick={() => { setMonth(row.month.slice(0, 7)); window.scrollTo({ top: 0, behavior: 'smooth' }); }}>Abrir mês</button></td></tr>)}</tbody></table></div>{!loading && !rows.length && <p className="info-note">{error ? 'Histórico indisponível.' : 'Nenhuma meta cadastrada. O dashboard usa os valores padrão.'}</p>}</Card>
    {pending && <Dialog title="Salvar metas em produção" close={() => { if (!busy) setPending(null); }}><p className="dialog-description">Mês {pending.month.slice(0, 7)}. Os valores serão usados também no dashboard atual; o imposto altera o cálculo do lucro líquido.</p><dl className="live-review">{targetFields.map(([key, label]) => <div key={key}><dt>{label}</dt><dd>{pending[key] == null ? 'Herdar valor anterior' : pending[key]!.toLocaleString('pt-BR')}</dd></div>)}</dl>{error && <p role="alert">{error}</p>}<div className="dialog-footer"><button className="button secondary" disabled={busy} onClick={() => setPending(null)}>Voltar</button><button className="button primary" disabled={busy} onClick={() => void save()}>{busy ? 'Salvando…' : 'Confirmar alteração'}</button></div></Dialog>}
  </>;
}
