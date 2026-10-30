import { useEffect, useRef, useState } from 'react';
import { Plus, RefreshCw, Users } from 'lucide-react';
import { Card, Dialog, Empty, PageHeading, Toolbar } from './components';
import { useAuth } from './Auth';
import { getAuthClient } from './api';
import { useUI } from './ui';
import { createConsoleUser, listConsoleUsers, roleLabels, updateConsoleUser, type ConsoleRole, type ConsoleUser } from './domain/users';
import './users.css';
export default function UserManagement() {
 const {role,refresh}=useAuth(),{notify}=useUI();
 const [rows,setRows]=useState<ConsoleUser[]>([]),[count,setCount]=useState(0),[query,setQuery]=useState(''),[busy,setBusy]=useState(false),[error,setError]=useState('');
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
 <div className="users-permissions"><Card title="Gerente"><p>Acessa os módulos do painel, tarefas, metas e financeiro.</p></Card><Card title="Administrador"><p>Também cria usuários, altera perfis e suspende ou reativa acessos.</p></Card></div>
 <Card title="Acessos ao painel" subtitle={`${count} usuário${count===1?'':'s'}`}><Toolbar value={query} onChange={setQuery} placeholder="Buscar por nome"><button className="button secondary" disabled={busy} onClick={()=>void load()}><RefreshCw size={15}/>Atualizar</button></Toolbar>
 {error&&<p role="alert">{error}</p>}{busy&&<p role="status">Carregando usuários…</p>}
 {!busy&&!error&&!rows.length&&<Empty title="Nenhum usuário encontrado" description="Tente outro nome na busca."/>}
 <div className="users-table-wrap"><table className="users-table"><thead><tr><th>Usuário</th><th>Perfil</th><th>Situação</th><th><span className="sr-only">Ações</span></th></tr></thead><tbody>{rows.map(row=><tr key={row.user_id}><td><strong>{row.display_name}</strong><small>{row.email}</small></td><td>{roleLabels[row.role]}</td><td><span className={`badge ${row.active?'active':'suspended'}`}>{row.active?'Ativo':'Suspenso'}</span></td><td><button className="button secondary" onClick={()=>setEditor(row)}>Gerenciar</button></td></tr>)}</tbody></table></div>
 {rows.length<count&&<button className="button secondary" disabled={busy} onClick={()=>void load(true)}>Carregar mais</button>}</Card>
 {editor&&<UserEditor member={editor==='new'?null:editor} close={()=>setEditor(null)} saved={async()=>{setEditor(null);await refresh();await load();notify(editor==='new'?'Usuário criado.':'Acesso atualizado.');}}/>}</div>;
}
function UserEditor({member,close,saved}:{member:ConsoleUser|null;close:()=>void;saved:()=>Promise<void>}) {
 const [role,setRole]=useState<ConsoleRole>(member?.role??'manager'),[active,setActive]=useState(member?.active??true),[busy,setBusy]=useState(false),[error,setError]=useState('');
 return <Dialog title={member?'Gerenciar acesso':'Novo usuário'} close={()=>{if(!busy)close();}}><form className="live-auth-form" onSubmit={async event=>{
  event.preventDefault();if(busy)return;setBusy(true);setError('');
  const form=event.currentTarget,values=new FormData(form);
  try { const client=await getAuthClient(); if(member) await updateConsoleUser(client,member,role,active);else await createConsoleUser(client,{name:String(values.get('name')),email:String(values.get('email')),password:String(values.get('password')),role});form.reset();await saved(); }
  catch(reason){setError((reason as Error).message);}finally{setBusy(false);}
 }}><div className="form-fields live-form">{member?<p><strong>{member.display_name}</strong><br/>{member.email}</p>:<><label>Nome<input name="name" autoComplete="name" required maxLength={120}/></label><label>Email<input name="email" type="email" autoComplete="off" required maxLength={254}/></label><label>Senha inicial<input name="password" type="password" autoComplete="new-password" required minLength={12} maxLength={128}/><small>De 12 a 128 caracteres. Entregue a senha ao usuário por um canal privado.</small></label></>}
 <label>Perfil<select aria-label="Perfil" value={role} onChange={event=>setRole(event.target.value as ConsoleRole)} disabled={busy}><option value="manager">Gerente</option><option value="administrator">Administrador</option></select></label>
 {member&&<label>Situação<select aria-label="Situação" value={String(active)} onChange={event=>setActive(event.target.value==='true')} disabled={busy}><option value="true">Ativo</option><option value="false">Suspenso</option></select></label>}
 <p className="info-note"><Users size={16}/>{role==='administrator'?'Pode gerenciar os usuários e todos os módulos do painel.':'Acessa o painel sem permissão para criar ou gerenciar usuários.'}</p>
 {member&&<p>É necessário manter pelo menos um administrador ativo. A suspensão bloqueia também tarefas e alterações financeiras.</p>}{error&&<p role="alert">{error}</p>}</div><div className="dialog-footer"><button className="button secondary" type="button" disabled={busy} onClick={close}>Cancelar</button><button className="button primary" disabled={busy}>{busy?'Salvando…':member?'Salvar acesso':'Criar usuário'}</button></div></form></Dialog>;
}
