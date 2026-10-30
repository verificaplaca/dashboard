import { handleRequest } from './index.ts';
import type { SupabaseClient } from 'npm:@supabase/supabase-js@2.117.3';
const assert = (value: unknown, message='Assertion failed') => { if(!value) throw new Error(message); };
function fixture(role='administrator',active=true,failProvision=false) {
 const calls: string[]=[];
 const client={
  auth:{getUser:async()=>({data:{user:{id:'verified-admin'}},error:null}),admin:{createUser:async()=>{calls.push('create');return {data:{user:{id:'created-user'}},error:null};},deleteUser:async()=>{calls.push('cleanup');return {error:null};}}},
  from:()=>({select:()=>({eq:()=>({maybeSingle:async()=>({data:{role,active},error:null})})})}),
  rpc:async(_name:string,args:Record<string,unknown>)=>{calls.push('provision');assert(args.p_actor==='verified-admin','Actor comes from verified token');return {data:{user_id:'created-user'},error:failProvision?{code:'42501'}:null};},
 } as unknown as SupabaseClient;
 return {client,calls};
}
const body={name:'Manager',email:'manager@example.invalid',password:'long password 123',role:'manager'};
function request(value:unknown=body,token=true){return new Request('https://local.invalid/console-users',{method:'POST',headers:{...(token?{Authorization:'Bearer user-token'}:{}),'Content-Type':'application/json'},body:JSON.stringify(value)});}
Deno.test('anonymous requests cannot create Auth accounts',async()=>{const f=fixture();assert((await handleRequest(request(body,false),f.client)).status===401);assert(f.calls.length===0);});
Deno.test('manager and inactive administrator cannot create accounts',async()=>{for(const [role,active] of [['manager',true],['administrator',false]] as const){const f=fixture(role,active);assert((await handleRequest(request(),f.client)).status===403);assert(f.calls.length===0);}});
Deno.test('caller cannot forge actor, id or role via body',async()=>{for(const extra of [{actor:'forged'},{user_id:'forged'},{role:'owner'}]){const f=fixture();assert((await handleRequest(request({...body,...extra}),f.client)).status===400);assert(f.calls.length===0);}});
Deno.test('valid administrator creates Auth and finalizes with verified caller',async()=>{const f=fixture();assert((await handleRequest(request(),f.client)).status===201);assert(f.calls.join(',')==='create,provision');});
Deno.test('permission revoked during creation rolls back the new Auth account',async()=>{const f=fixture('administrator',true,true);assert((await handleRequest(request(),f.client)).status===409);assert(f.calls.join(',')==='create,provision,cleanup');});
Deno.test('failed token validation does not reach privileged mutation',async()=>{const f=fixture();f.client.auth.getUser=(()=>Promise.resolve({data:{user:null},error:new Error('Invalid')})) as typeof f.client.auth.getUser;assert((await handleRequest(request(),f.client)).status===401);assert(f.calls.length===0);});
