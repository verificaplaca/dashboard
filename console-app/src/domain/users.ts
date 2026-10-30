import type { SupabaseClient } from '@supabase/supabase-js';
export type ConsoleRole = 'administrator' | 'manager';
export const roleLabels: Record<ConsoleRole,string> = {administrator:'Administrador',manager:'Gerente'};
export interface ConsoleUser { user_id: string; display_name: string; email: string; role: ConsoleRole; active: boolean; version: number; created_at: string; updated_at: string }
export async function readConsoleAccess(client: SupabaseClient, id: string): Promise<ConsoleUser | null> {
 const {data,error}=await client.from('console_users').select('*').eq('user_id',id).maybeSingle();
 if(error) throw new Error('Não foi possível verificar seu acesso ao painel.');
 return data;
}
export async function listConsoleUsers(client: SupabaseClient, search: string, offset: number) {
 let query=client.from('console_users').select('*',{count:'exact'}).order('created_at',{ascending:false}).order('user_id').range(offset,offset+49);
 // Escape PostgREST filter delimiters; search by name only, email remains visible.
 if(search.trim()) query=query.ilike('display_name',`%${search.trim().replace(/[\\%_]/g,'\\$&')}%`);
 const {data,error,count}=await query;
 if(error) throw new Error('Não foi possível carregar os usuários. Tente atualizar.');
 return {users:(data??[]) as ConsoleUser[],count:count??0};
}
export async function createConsoleUser(client: SupabaseClient, values: {name:string;email:string;password:string;role:ConsoleRole}) {
 const {data,error}=await client.functions.invoke('console-users',{body:values});
 if(error) {
  let message='Não foi possível confirmar a criação. Atualize a lista antes de tentar novamente.';
  if(error.context instanceof Response) { try { const body=await error.context.json(); if(typeof body.error==='string') message=body.error; } catch {} }
  throw new Error(message);
 }
 if(!data?.user) throw new Error('Resposta inválida. Atualize a lista de usuários.');
 return data.user as ConsoleUser;
}
export async function updateConsoleUser(client: SupabaseClient, user: ConsoleUser, role: ConsoleRole, active: boolean) {
 const {error}=await client.rpc('console_update_user',{p_id:user.user_id,p_role:role,p_active:active,p_version:user.version});
 if(error) throw new Error(error.code==='40001'?'O usuário mudou. Atualize a lista antes de salvar.':error.code==='22023'?error.message:'Sem permissão para alterar este usuário. Atualize sua sessão.');
}
export type ConsoleUserChange={action:'update';name:string;email:string;role:ConsoleRole;active:boolean}|{action:'reset_password';password:string};
export async function changeConsoleUser(client:SupabaseClient,user:ConsoleUser,change:ConsoleUserChange){
 const {data,error}=await client.functions.invoke('console-users',{body:{...change,id:user.user_id,version:user.version}});
 if(error){
  let message='Não foi possível confirmar a alteração. Atualize a lista antes de tentar novamente.';
  if(error.context instanceof Response){try{const body=await error.context.json();if(typeof body.error==='string')message=body.error;}catch{}}
  throw new Error(message);
 }
 if(!data?.user)throw new Error('Resposta inválida. Atualize a lista antes de tentar novamente.');
 return data.user as ConsoleUser;
}
