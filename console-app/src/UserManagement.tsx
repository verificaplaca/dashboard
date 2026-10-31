import { useGlobalLoading } from './Loading';
import { useEffect, useRef, useState } from 'react';
import { KeyRound, Mail, Plus, RefreshCw, Users } from 'lucide-react';
import { Card, Dialog, Empty, PageHeading, Toolbar } from './components';
import { useAuth } from './Auth';
import { getAuthClient } from './api';
import { useUI } from './ui';
import { createConsoleUser, listConsoleUsers, roleLabels, changeConsoleUser, sendRecoveryEmail, type ConsoleRole, type ConsoleUser } from './domain/users';
import './users.css';
export default function UserManagement() {
 const {role,refresh}=useAuth(),{notify}=useUI();
 const [rows,setRows]=useState<ConsoleUser[]>([]),[count,setCount]=useState(0),[query,setQuery]=useState(''),[busy,setBusy]=useState(true),[error,setError]=useState('');
 useGlobalLoading(busy);
 const [recoveryUser,setRecoveryUser]=useState<ConsoleUser|null>(null);
 const [passwordUser,setPasswordUser]=useState<ConsoleUser|null>(null);
 const [editor,setEditor]=useState<ConsoleUser|'new'|null>(null),revision=useRef(0);
 const load=async(append=false)=>{
  const request=++revision.current; setBusy(true); setError('');
  try { const result=await listConsoleUsers(await getAuthClient(),query,append?rows.length:0); if(request===revision.current){setRows(previous=>append?[...previous,...result.users]:result.users);setCount(result.count);} }
  catch(reason){if(request===revision.current)setError((reason as Error).message);}
  finally{if(request===revision.current)setBusy(false);}
 };
 useEffect(()=>{const timer=setTimeout(()=>void load(),250);return()=>{clearTimeout(timer);revision.current++;};},[query]);
 if(role!=='administrator') return <Empty title="Acesso restrito" description="Somente administradores podem gerenciar usuários." />;
 return <div className="module-page users-page"><PageHeading eyebrow="ADMINISTRAÇÃO" title="Usuários" description="Controle quem pode acessar o painel do Verifica Placa."><button className="button primary" onClick={()=>setEditor('new')}><Plus size={16}/>Novo usuário</button></PageHeading>
 <details className="mobile-records users-role-guide"><summary>Como funcionam os perfis</summary><p><strong>Gerente:</strong> acessa o painel, tarefas e financeiro. Consulta as metas.</p><p><strong>Administrador:</strong> também altera metas, gerencia usuários e redefine senhas.</p></details>
 <div className="users-permissions desktop-records"><Card title="Gerente"><p>Acessa o painel, tarefas e financeiro. Consulta as metas em modo de leitura.</p></Card><Card title="Administrador"><p>Também altera metas e dados dos usuários, redefine senhas e gerencia os acessos.</p></Card></div>
 <Card title="Acessos ao painel" subtitle={`${count} usuário${count===1?'':'s'}`}><Toolbar value={query} onChange={setQuery} placeholder="Buscar por nome"><button className="button secondary" disabled={busy} onClick={()=>void load()}><RefreshCw size={15}/>Atualizar</button></Toolbar>
 {error&&<p role="alert">{error}</p>}{busy&&<p role="status">Carregando usuários…</p>}
 {!busy&&!error&&!rows.length&&<Empty title="Nenhum usuário encontrado" description="Tente outro nome na busca."/>}
 <ul className="mobile-records users-mobile-list" aria-label="Usuários do painel">{rows.map(row=><li key={row.user_id}><header><strong>{row.display_name}</strong><span className={`badge ${row.active?'active':'suspended'}`}>{row.active?'Ativo':'Suspenso'}</span></header><p>{row.email}</p><span className="record-role">{roleLabels[row.role]}</span><div className="users-actions"><button className="button secondary" onClick={()=>setEditor(row)}>Gerenciar</button><button className="button secondary" onClick={()=>setPasswordUser(row)}><KeyRound size={15}/>Redefinir senha</button><button className="button secondary" disabled={!row.active} onClick={()=>setRecoveryUser(row)}><Mail size={15}/>Enviar link por email</button></div></li>)}</ul>
 <div className="users-table-wrap desktop-records"><table className="users-table"><thead><tr><th>Usuário</th><th>Perfil</th><th>Situação</th><th><span className="sr-only">Ações</span></th></tr></thead><tbody>{rows.map(row=><tr key={row.user_id}><td><strong>{row.display_name}</strong><small>{row.email}</small></td><td>{roleLabels[row.role]}</td><td><span className={`badge ${row.active?'active':'suspended'}`}>{row.active?'Ativo':'Suspenso'}</span></td><td><div className="users-actions"><button className="button secondary" onClick={()=>setEditor(row)}>Gerenciar</button><button className="button secondary" onClick={()=>setPasswordUser(row)} aria-label={`Redefinir senha de ${row.display_name}`}><KeyRound size={15}/>Redefinir senha</button><button className="button secondary" disabled={!row.active} title={!row.active?'Reative o usuário para enviar o link':undefined} onClick={()=>setRecoveryUser(row)} aria-label={`Enviar link por email para ${row.display_name}`}><Mail size={15}/>Enviar link por email</button></div></td></tr>)}</tbody></table></div>
 {rows.length<count&&<button className="button secondary" disabled={busy} onClick={()=>void load(true)}>Carregar mais</button>}</Card>
 {editor&&<UserEditor member={editor==='new'?null:editor} close={()=>setEditor(null)} saved={async()=>{setEditor(null);await refresh();await load();notify(editor==='new'?'Usuário criado.':'Usuário atualizado.');}}/>}{passwordUser&&<PasswordEditor member={passwordUser} close={()=>setPasswordUser(null)} saved={async()=>{setPasswordUser(null);await refresh();await load();notify('Senha redefinida. Entregue a nova senha ao usuário por um canal privado.');}}/>}{recoveryUser&&<RecoveryEmailDialog member={recoveryUser} close={()=>setRecoveryUser(null)}/>}</div>;
}
function UserEditor({member,close,saved}:{member:ConsoleUser|null;close:()=>void;saved:()=>Promise<void>}) {
 const [review,setReview]=useState(false),[reviewCopy,setReviewCopy]=useState({name:'',email:''});
 const [role,setRole]=useState<ConsoleRole>(member?.role??'manager'),[active,setActive]=useState(member?.active??true),[busy,setBusy]=useState(false),[error,setError]=useState('');
 return <Dialog title={member?'Gerenciar usuário':'Novo usuário'} close={()=>{if(!busy)close();}}><form className={`live-auth-form ${review?'user-editor-reviewing':''}`} onSubmit={async event=>{
  event.preventDefault();if(busy)return;setBusy(true);setError('');
  const form=event.currentTarget,values=new FormData(form);
  if(member&&!review){setReviewCopy({name:String(values.get('name')),email:String(values.get('email'))});setReview(true);setBusy(false);return;}
  try { const client=await getAuthClient(); if(member) await changeConsoleUser(client,member,{action:'update',name:String(values.get('name')),email:String(values.get('email')),role,active});else await createConsoleUser(client,{name:String(values.get('name')),email:String(values.get('email')),password:String(values.get('password')),role});form.reset();await saved(); }
  catch(reason){setError((reason as Error).message);}finally{setBusy(false);}
 }}><div className="form-fields live-form">{review&&<dl className="mobile-user-review"><div><dt>Nome</dt><dd>{reviewCopy.name}</dd></div><div><dt>Email</dt><dd>{reviewCopy.email}</dd></div><div><dt>Perfil</dt><dd>{roleLabels[role]}</dd></div><div><dt>Situação</dt><dd>{active?'Ativo':'Suspenso'}</dd></div></dl>}<label>Nome<input name="name" autoComplete="name" defaultValue={member?.display_name||''} required maxLength={120} readOnly={busy||review}/></label><label>Email<input name="email" type="email" autoComplete="off" defaultValue={member?.email||''} required maxLength={254} readOnly={busy||review}/><small>{member?'Este será o email usado para entrar no painel. A alteração é imediata.':''}</small></label>{!member&&<label>Senha inicial<input name="password" type="password" autoComplete="new-password" required minLength={12} maxLength={128} disabled={busy}/><small>De 12 a 128 caracteres. Entregue a senha ao usuário por um canal privado.</small></label>}
 <label>Perfil<select aria-label="Perfil" value={role} onChange={event=>setRole(event.target.value as ConsoleRole)} disabled={busy||review}><option value="manager">Gerente</option><option value="administrator">Administrador</option></select></label>
 {member&&<label>Situação<select aria-label="Situação" value={String(active)} onChange={event=>setActive(event.target.value==='true')} disabled={busy||review}><option value="true">Ativo</option><option value="false">Suspenso</option></select></label>}
 <p className="info-note"><Users size={16}/>{role==='administrator'?'Pode gerenciar os usuários e todos os módulos do painel.':'Acessa o painel sem permissão para criar ou gerenciar usuários.'}</p>
 {member&&<p>É necessário manter pelo menos um administrador ativo. A suspensão bloqueia também tarefas e alterações financeiras.</p>}{review&&<p className="info-note">Confira nome, email, perfil e situação acima. Ao confirmar, o email de acesso e os dados serão atualizados.</p>}{error&&<p role="alert">{error}</p>}</div><div className="dialog-footer"><button className="button secondary" type="button" disabled={busy} onClick={()=>review?setReview(false):close()}>{review?'Voltar':'Cancelar'}</button><button className="button primary" disabled={busy}>{busy?'Salvando…':member?(review?'Confirmar alteração':'Revisar alterações'):'Criar usuário'}</button></div></form></Dialog>;
}

function PasswordEditor({member,close,saved}:{member:ConsoleUser;close:()=>void;saved:()=>Promise<void>}){
 const [busy,setBusy]=useState(false),[error,setError]=useState('');
 return <Dialog title="Redefinir senha" close={()=>{if(!busy)close();}}><form className="live-auth-form" onSubmit={async event=>{
  event.preventDefault();if(busy)return;const form=event.currentTarget,values=new FormData(form),password=String(values.get('password'));
  if(password!==values.get('confirmation')){setError('As senhas devem ser iguais.');return;}
  setBusy(true);setError('');
  try{await changeConsoleUser(await getAuthClient(),member,{action:'reset_password',password});form.reset();await saved();}catch(reason){setError((reason as Error).message);}finally{setBusy(false);}
 }}><div className="form-fields live-form"><p><strong>{member.display_name}</strong><br/>{member.email}</p><p>A nova senha substitui a atual imediatamente. Entregue-a ao usuário por um canal privado.</p><label>Nova senha<input autoFocus name="password" type="password" autoComplete="new-password" required minLength={12} maxLength={128} disabled={busy}/><small>Use de 12 a 128 caracteres.</small></label><label>Confirmar nova senha<input name="confirmation" type="password" autoComplete="new-password" required minLength={12} maxLength={128} disabled={busy}/></label>{error&&<p role="alert">{error}</p>}</div><div className="dialog-footer"><button className="button secondary" type="button" disabled={busy} onClick={close}>Cancelar</button><button className="button primary" disabled={busy}>{busy?'Redefinindo…':'Confirmar nova senha'}</button></div></form></Dialog>;
}

function RecoveryEmailDialog({member,close}:{member:ConsoleUser;close:()=>void}){
 const [busy,setBusy]=useState(false),[sent,setSent]=useState(false),[error,setError]=useState('');
 return <Dialog title="Enviar link de redefinição" close={()=>{if(!busy)close();}}><div className="form-fields live-form"><p><strong>{member.display_name}</strong><br/>{member.email}</p>{sent?<p role="status">Envio solicitado para {member.email}. O usuário poderá criar a própria senha pelo link. Oriente-o a verificar também a pasta de spam.</p>:<p>O usuário receberá um link para criar a própria senha. A senha atual continua válida até ele concluir a alteração.</p>}{error&&<p role="alert">{error}</p>}</div><div className="dialog-footer"><button className="button secondary" disabled={busy} onClick={close}>{sent?'Fechar':'Cancelar'}</button>{!sent&&<button className="button primary" disabled={busy} onClick={async()=>{if(busy)return;setBusy(true);setError('');try{await sendRecoveryEmail(await getAuthClient(),member);setSent(true);}catch(reason){setError((reason as Error).message);}finally{setBusy(false);}}}>{busy?'Enviando…':'Enviar link por email'}</button>}</div></Dialog>;
}
