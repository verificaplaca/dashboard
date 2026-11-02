import { useId, useState, type FormEvent } from 'react';
import { ArrowRight, Eye, EyeOff, LockKeyhole, ShieldCheck } from 'lucide-react';
import './sign-in-page.css';

type LoginPageProps = {
  onSignIn: (email: string, password: string) => Promise<void>;
  compact?: boolean;
};

/** Login presentation shared by the access page and the reauthentication dialog. */
export function LoginPage({ onSignIn, compact = false }: LoginPageProps) {
  const id = useId();
  const [showPassword, setShowPassword] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [help, setHelp] = useState(false);

  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (busy) return;
    const form = event.currentTarget;
    const values = new FormData(form);
    setBusy(true);
    setError('');
    try {
      await onSignIn(String(values.get('email')).trim(), String(values.get('password')));
      form.reset();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Não foi possível entrar. Tente novamente.');
    } finally {
      setBusy(false);
    }
  }

  return <div className={`vf-sign-in${compact ? ' vf-sign-in--compact' : ''}`}>
    {!compact && <section className="vf-sign-in-art" aria-label="Verifica Placa · Console de gestão">
      <img className="vf-sign-in-image" src={`${import.meta.env.BASE_URL}images/sign-in-art.png`} alt="" fetchPriority="high" />
      <div className="vf-sign-in-art-shade" />
      <div className="vf-sign-in-brand vf-sign-in-brand--light"><img src={`${import.meta.env.BASE_URL}verifica-placa-logo.png`} alt="" /><span>Verifica Placa<small>CONSOLE DE GESTÃO</small></span></div>
      <div className="vf-sign-in-story"><span className="vf-sign-in-pill"><span /> SUA OPERAÇÃO, CONECTADA</span><h2>Uma visão clara.<br />Novas possibilidades.</h2><p>Resultados, campanhas e metas.<br />Tudo conectado para a próxima decisão.</p></div>
      <div className="vf-sign-in-art-footer"><span>Verifica Placa · Business Console</span><ShieldCheck size={18} aria-hidden="true" /></div>
    </section>}
    <section className="vf-sign-in-panel" aria-labelledby={`${id}-title`}>
      <div className="vf-sign-in-form-wrap">
        <div className="vf-sign-in-brand"><img src={`${import.meta.env.BASE_URL}verifica-placa-logo.png`} alt="" /><span>Verifica Placa<small>CONSOLE DE GESTÃO</small></span></div>
        <header className="vf-sign-in-heading"><span className="vf-sign-in-eyebrow">BEM-VINDO AO SEU PAINEL</span><h1 id={`${id}-title`}>Bom ter você de volta.</h1><p>Entre na sua conta para acompanhar a operação.</p></header>
        <form onSubmit={submit} aria-busy={busy} className="vf-sign-in-form">
          <div className="vf-sign-in-field"><label htmlFor={`${id}-email`}>Email</label><input id={`${id}-email`} name="email" type="email" autoComplete="username" placeholder="seu@email.com.br" required disabled={busy} aria-describedby={error ? `${id}-error` : undefined} /></div>
          <div className="vf-sign-in-field"><label htmlFor={`${id}-password`}>Senha</label><div className="vf-sign-in-password"><input id={`${id}-password`} name="password" type={showPassword ? 'text' : 'password'} autoComplete="current-password" placeholder="Digite sua senha" required disabled={busy} aria-describedby={error ? `${id}-error` : undefined} /><button type="button" onClick={() => setShowPassword(value => !value)} aria-label={showPassword ? 'Ocultar senha' : 'Mostrar senha'} aria-pressed={showPassword} disabled={busy}>{showPassword ? <EyeOff size={19} /> : <Eye size={19} />}</button></div></div>
          <div className="vf-sign-in-options"><span><LockKeyhole size={13} aria-hidden="true" /> Acesso autorizado</span><button type="button" aria-expanded={help} aria-controls={`${id}-help`} onClick={() => setHelp(value => !value)}>Esqueceu a senha?</button></div>
          {help && <p id={`${id}-help`} className="vf-sign-in-help" role="status">Peça a um administrador do painel o envio de um link de recuperação para o seu email.</p>}
          {error && <p id={`${id}-error`} className="vf-sign-in-error" role="alert">{error}</p>}
          <button type="submit" className="vf-sign-in-submit" disabled={busy}><span>{busy ? 'Entrando…' : 'Entrar no painel'}</span>{!busy && <ArrowRight size={18} aria-hidden="true" />}</button>
        </form>
        <p className="vf-sign-in-invite">Ainda não tem acesso?<br /><span>Solicite sua conta a um administrador.</span></p>
        <div className="vf-sign-in-security"><ShieldCheck size={16} aria-hidden="true" /><span>Acesso exclusivo para a equipe Verifica Placa.</span></div>
      </div>
      {!compact && <footer className="vf-sign-in-footer">Verifica Placa <span>·</span> Gestão com mais clareza.</footer>}
    </section>
  </div>;
}
