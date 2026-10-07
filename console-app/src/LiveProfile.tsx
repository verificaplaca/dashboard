import { useEffect, useState } from 'react';
import { Avatar, Card, PageHeading } from './components';
import { useAuth } from './Auth';
import { useUI } from './ui';
import { getAuthClient } from './api';
export default function LiveProfile() {
  const { user, openLogin, refresh, signOut } = useAuth(), { notify } = useUI();
  const [name, setName] = useState(''), [busy, setBusy] = useState(false), [error, setError] = useState('');
  useEffect(() => { setName(typeof user?.user_metadata?.full_name === 'string' ? user.user_metadata.full_name : ''); }, [user]);
  return <><PageHeading eyebrow="SUA CONTA" title="Meu perfil" description="Identidade autenticada no Supabase da operação." />{!user ? <Card title="Entre para acessar seu perfil"><p>Use a mesma conta que edita as metas.</p><button className="button primary" onClick={openLogin}>Entrar</button></Card> : <Card title="Informações pessoais"><div className="live-identity"><Avatar name={name || user.email || 'Conta'} /><div><strong>{name || 'Minha conta'}</strong><small>{user.email}</small></div></div><form className="live-form" onSubmit={async event => {
    event.preventDefault(); if (busy) return;
    setBusy(true); setError('');
    try { const client = await getAuthClient(); const { error } = await client.auth.updateUser({ data: { full_name: name.trim() } }); if (error) throw new Error('Não foi possível salvar o nome. Tente novamente.'); await refresh(); notify('Perfil atualizado.'); }
    catch (reason) { setError((reason as Error).message); } finally { setBusy(false); }
  }}><label>Nome de exibição<input value={name} onChange={event => setName(event.target.value)} maxLength={120} required /></label><label>Email<input value={user.email || ''} disabled /></label>{error && <p role="alert">{error}</p>}<div className="live-form-footer"><button type="button" className="button secondary" onClick={() => void signOut()}>Sair desta sessão</button><button className="button primary" disabled={busy}>{busy ? 'Salvando…' : 'Salvar perfil'}</button></div></form></Card>}</>;
}
