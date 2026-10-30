import test from 'node:test';
import assert from 'node:assert/strict';
import type {SupabaseClient} from '@supabase/supabase-js';
import {parseRecoveryEntry,beginPasswordRecovery,finishPasswordRecovery,validateRecoveryPassword} from './password-recovery';
import {sendRecoveryEmail,type ConsoleUser} from './users';
test('recovery entry accepts only recovery links and never uses an unrelated existing session',()=>{
 assert.equal(parseRecoveryEntry('','#/inicio'),null);
 assert.equal(parseRecoveryEntry('?demo=1','#/inicio'),null);
 assert.deepEqual(parseRecoveryEntry('?recovery=1','#token_hash='+ 'a'.repeat(64)+'&type=recovery'),{kind:'hash',token_hash:'a'.repeat(64)});
 assert.deepEqual(parseRecoveryEntry('?recovery=1','#access_token=fixture&refresh_token=fixture-refresh&type=recovery'),{kind:'session',access_token:'fixture',refresh_token:'fixture-refresh'});
 for(const hash of ['','#type=signup&access_token=wrong&refresh_token=wrong','#type=recovery&token_hash=short','#error=access_denied&error_code=otp_expired'])assert.deepEqual(parseRecoveryEntry('?recovery=1',hash),{kind:'invalid'});
});
function fixture({active=true,verifyError=false,writeError=false}={}){
 const calls:unknown[]=[];
 const identity={id:'recovered-user'};
 const client={auth:{verifyOtp:async(args:unknown)=>{calls.push(args);return {data:{session:{},user:identity},error:verifyError?{}:null};},setSession:async(args:unknown)=>{calls.push(args);return {data:{session:{},user:identity},error:null};},getUser:async()=>({data:{user:identity},error:null}),updateUser:async(args:unknown)=>{calls.push(args);return {error:writeError?{}:null};}},from:()=>({select:()=>({eq:(_key:string,id:string)=>{assert.equal(id,'recovered-user');return {maybeSingle:async()=>({data:{active,role:'manager'},error:null})};}})})} as unknown as SupabaseClient;
 return {client,calls};
}
test('single-use token is exchanged as recovery and verified server identity controls eligibility',async()=>{
 const f=fixture();assert.equal((await beginPasswordRecovery(f.client,{kind:'hash',token_hash:'fixture-hash'})).id,'recovered-user');
 assert.deepEqual(f.calls,[{type:'recovery',token_hash:'fixture-hash'}]);
 for(const f of [fixture({active:false}),fixture({verifyError:true})])await assert.rejects(beginPasswordRecovery(f.client,{kind:'hash',token_hash:'fixture-hash'}));
 const invalid=fixture();await assert.rejects(beginPasswordRecovery(invalid.client,{kind:'invalid'}));assert.equal(invalid.calls.length,0);
});
test('password validation and fresh membership prevent invalid or suspended writes',async()=>{
 for(const [a,b] of [['short','short'],['a'.repeat(129),'a'.repeat(129)],['password 12345','different 12345']])assert.throws(()=>validateRecoveryPassword(a,b));
 const blocked=fixture({active:false});await assert.rejects(finishPasswordRecovery(blocked.client,'replacement 12345','replacement 12345'),/ativa/);assert.equal(blocked.calls.length,0);
 const f=fixture();await finishPasswordRecovery(f.client,'replacement 12345','replacement 12345');assert.deepEqual(f.calls,[{password:'replacement 12345'}]);
 await assert.rejects(finishPasswordRecovery(fixture({writeError:true}).client,'replacement 12345','replacement 12345'),/Não foi possível/);
});
test('email request sends only target/version and preserves transport failure without false success',async()=>{
 let body:unknown;
 const client={functions:{invoke:async(_name:string,args:unknown)=>{body=args;return {data:{sent:true,email:'recipient@example.invalid'},error:null};}}} as unknown as SupabaseClient;
 const member={user_id:'target',version:5} as ConsoleUser;
 assert.equal(await sendRecoveryEmail(client,member),'recipient@example.invalid');
 assert.deepEqual(body,{body:{action:'send_recovery',id:'target',version:5}});
 client.functions.invoke=(async()=>({data:null,error:{context:new Response(JSON.stringify({error:'Remetente pendente.'}))}})) as typeof client.functions.invoke;
 await assert.rejects(sendRecoveryEmail(client,member),/pendente/);
});
