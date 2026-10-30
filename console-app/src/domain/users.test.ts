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
