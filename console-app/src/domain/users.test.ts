import test from 'node:test';
import assert from 'node:assert/strict';
import type {SupabaseClient} from '@supabase/supabase-js';
import {readConsoleAccess,createConsoleUser,updateConsoleUser,type ConsoleUser} from './users';
import {validateNewUser} from '../../../supabase/functions/console-users/validation';
test('creation accepts only the two server roles and rejects forged fields',()=>{
 const base={name:' Ana ',email:'ANA@example.com',password:'long password 123',role:'manager'};
 assert.equal(validateNewUser(base).email,'ana@example.com');
 assert.equal(validateNewUser(base).name,'Ana');
 assert.equal(validateNewUser({...base,role:'administrator'}).role,'administrator');
 for(const changes of [{role:'owner'},{role:null},{user_id:'forged'},{password:'short'},{name:' '},{email:'invalid'},{password:'a'.repeat(129)}]) assert.throws(()=>validateNewUser({...base,...changes}));
});
test('permission lookup fails closed when database is unavailable',async()=>{
 const query={eq(){return this;},maybeSingle:async()=>({data:null,error:{code:'503'}})};
 const client={from:()=>({select:()=>query})} as unknown as SupabaseClient;
 await assert.rejects(readConsoleAccess(client,'operator'),/verificar/);
});
test('create does not report success on an edge error and preserves useful denial',async()=>{
 const client={functions:{invoke:async()=>({data:null,error:{context:new Response(JSON.stringify({error:'Somente administradores podem criar usuários.'}))}})}} as unknown as SupabaseClient;
 await assert.rejects(createConsoleUser(client,{name:'Ana',email:'ana@example.com',password:'password 12345',role:'manager'}),/Somente administradores/);
});
test('access update sends observed version; concurrent change preserves server conflict',async()=>{
 const member={user_id:'target',version:3} as ConsoleUser;
 let args:unknown;
 const client={rpc:async(_name:string,values:unknown)=>{args=values;return {error:{code:'40001'}};}} as unknown as SupabaseClient;
 await assert.rejects(updateConsoleUser(client,member,'manager',false),/mudou/);
 assert.deepEqual(args,{p_id:'target',p_role:'manager',p_active:false,p_version:3});
});

test('profile/reset validation rejects forged authority and keeps passwords out of profile edits',async()=>{
 const {validateUserChange}=await import('../../../supabase/functions/console-users/validation');
 const base={action:'update',id:'90000000-0000-4000-8000-000000000071',version:2,name:' Ana ',email:'ANA@example.com',role:'manager',active:true};
 assert.deepEqual(validateUserChange(base),{...base,name:'Ana',email:'ana@example.com'});
 for(const changes of [{actor:'forged'},{id:'other'},{version:0},{version:'2'},{active:'true'},{password:'secret'},{role:'owner'}])assert.throws(()=>validateUserChange({...base,...changes}));
 const reset={action:'reset_password',id:base.id,version:2,password:'valid password 123'};
 assert.equal(validateUserChange(reset).action,'reset_password');
 for(const changes of [{password:'short'},{password:'x'.repeat(129)},{role:'administrator'},{email:'other@example.com'}])assert.throws(()=>validateUserChange({...reset,...changes}));
});
test('user edits/reset send target version to server and never report false success',async()=>{
 const {changeConsoleUser}=await import('./users');
 let sent:unknown;
 const member={user_id:'target',version:4} as ConsoleUser;
 const client={functions:{invoke:async(_name:string,args:unknown)=>{sent=args;return {data:{user:{...member,version:5}},error:null};}}} as unknown as SupabaseClient;
 await changeConsoleUser(client,member,{action:'reset_password',password:'valid password 123'});
 assert.deepEqual(sent,{body:{action:'reset_password',password:'valid password 123',id:'target',version:4}});
 client.functions.invoke=(async()=>({data:null,error:{context:new Response(JSON.stringify({error:'Este email já tem uma conta.'}))}})) as typeof client.functions.invoke;
 await assert.rejects(changeConsoleUser(client,member,{action:'update',name:'Ana',email:'used@example.com',role:'manager',active:true}),/já tem/);
});
