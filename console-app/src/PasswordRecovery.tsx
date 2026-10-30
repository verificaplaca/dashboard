import {useEffect,useRef,useState} from 'react';
import {createClient,type SupabaseClient} from '@supabase/supabase-js';
import {CheckCircle2,KeyRound} from 'lucide-react';
import config from './production-config.json';
import {beginPasswordRecovery,finishPasswordRecovery,type RecoveryEntry} from './domain/password-recovery';
import './password-recovery.css';
export default function PasswordRecovery({entry}:{entry:RecoveryEntry}){
 const client=useRef<SupabaseClient|null>(null),attempt=useRef<Promise<unknown>|null>(null);
 const [stage,setStage]=useState<'link'|'ready'|'done'>('link'),[busy,setBusy]=useState(false),[error,setError]=useState(entry.kind==='invalid'?'Este link é inválido ou expirou. Solicite um novo link ao administrador.':'');
 const getClient=()=>client.current??=createClient(config.url,config.key,{auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false,storageKey:'vp-password-recovery'}});
 useEffect(()=>()=>{void client.current?.auth.stopAutoRefresh();},[]);
 const validate=async()=>{
  if(busy)return;setBusy(true);setError('');
  try{await (attempt.current??=beginPasswordRecovery(getClient(),entry));setStage('ready');}catch(reason){setError((reason as Error).message);}finally{setBusy(false);}
 };
 return <main className="password-recovery"><section className="recovery-card" aria-labelledby="recovery-title"><div className="recovery-brand"><img src={`${import.meta.env.BASE_URL}verifica-placa-logo.png`} alt="Verifica Placa"/><span>Verifica Placa</span></div><div className="recovery-icon">{stage==='done'?<CheckCircle2 size={28}/>:<KeyRound size={28}/>}</div><h1 id="recovery-title">{stage==='done'?'Senha criada':'Crie sua nova senha'}</h1>
 {stage==='done'?<><p role="status">Sua senha foi alterada. Entre no painel com a nova senha.</p><a className="button primary" href={`${import.meta.env.BASE_URL}#/inicio`}>Ir para o painel</a></>:stage==='link'?<><p>Continue para validar seu link e escolher uma senha que só você conhece.</p>{error&&<p className="recovery-error" role="alert">{error}</p>}{entry.kind!=='invalid'&&!error&&<button className="button primary" disabled={busy} onClick={()=>void validate()}>{busy?'Validando link…':'Continuar'}</button>}<a className="text-button" href={`${import.meta.env.BASE_URL}#/inicio`}>Voltar ao painel</a></>:<form onSubmit={async event=>{
  event.preventDefault();if(busy)return;const form=event.currentTarget,values=new FormData(form);setBusy(true);setError('');
  try{await finishPasswordRecovery(getClient(),String(values.get('password')),String(values.get('confirmation')));form.reset();setStage('done');void getClient().auth.signOut({scope:'local'}).catch(()=>{});}catch(reason){setError((reason as Error).message);}finally{setBusy(false);}
 }}><p>Use de 12 a 128 caracteres. Evite reutilizar senhas de outros serviços.</p><label>Nova senha<input autoFocus name="password" type="password" autoComplete="new-password" required minLength={12} maxLength={128} disabled={busy}/></label><label>Confirmar nova senha<input name="confirmation" type="password" autoComplete="new-password" required minLength={12} maxLength={128} disabled={busy}/></label>{error&&<p className="recovery-error" role="alert">{error}</p>}<button className="button primary" disabled={busy}>{busy?'Salvando…':'Salvar nova senha'}</button></form>}
 </section></main>;
}
