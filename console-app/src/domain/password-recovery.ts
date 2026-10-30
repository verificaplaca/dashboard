import type { SupabaseClient } from '@supabase/supabase-js';
import {readConsoleAccess} from './users';
export type RecoveryEntry={kind:'hash';token_hash:string}|{kind:'session';access_token:string;refresh_token:string}|{kind:'invalid'};
export function parseRecoveryEntry(search:string,hash:string):RecoveryEntry|null{
 const query=new URLSearchParams(search),fragment=new URLSearchParams(hash.replace(/^#/,''));
 if(query.get('recovery')!=='1'&&fragment.get('type')!=='recovery')return null;
 if(fragment.get('error')||query.get('error'))return {kind:'invalid'};
 if(fragment.get('type')!=='recovery')return {kind:'invalid'};
 const token_hash=fragment.get('token_hash');
 if(token_hash&&/^[a-zA-Z0-9_-]{20,512}$/.test(token_hash))return {kind:'hash',token_hash};
 const access_token=fragment.get('access_token'),refresh_token=fragment.get('refresh_token');
 if(access_token&&refresh_token&&access_token.length<8192&&refresh_token.length<1024)return {kind:'session',access_token,refresh_token};
 return {kind:'invalid'};
}
export async function beginPasswordRecovery(client:SupabaseClient,entry:RecoveryEntry){
 if(entry.kind==='invalid')throw new Error('Este link é inválido ou expirou. Solicite um novo link ao administrador.');
 const result=entry.kind==='hash'?await client.auth.verifyOtp({type:'recovery',token_hash:entry.token_hash}):await client.auth.setSession({access_token:entry.access_token,refresh_token:entry.refresh_token});
 if(result.error||!result.data.session||!result.data.user)throw new Error('Este link é inválido ou expirou. Solicite um novo link ao administrador.');
 const {data,error}=await client.auth.getUser();
 if(error||!data.user)throw new Error('Não foi possível validar o link. Solicite outro ao administrador.');
 const member=await readConsoleAccess(client,data.user.id);
 if(!member?.active)throw new Error('Sua conta não está ativa. Procure um administrador.');
 return data.user;
}
export function validateRecoveryPassword(password:string,confirmation:string){
 if(password.length<12||password.length>128)throw new Error('Use uma senha entre 12 e 128 caracteres.');
 if(password!==confirmation)throw new Error('As senhas devem ser iguais.');
}
export async function finishPasswordRecovery(client:SupabaseClient,password:string,confirmation:string){
 validateRecoveryPassword(password,confirmation);
 // Recheck current identity and membership immediately before the Auth write.
 const {data,error}=await client.auth.getUser();
 if(error||!data.user)throw new Error('O link expirou. Solicite outro ao administrador.');
 const member=await readConsoleAccess(client,data.user.id);
 if(!member?.active)throw new Error('Sua conta não está ativa. Procure um administrador.');
 const result=await client.auth.updateUser({password});
 if(result.error)throw new Error('Não foi possível salvar. Use uma senha diferente da atual ou solicite um novo link se ele expirou.');
}
