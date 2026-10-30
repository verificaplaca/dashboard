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
const edit={action:'update',id:'90000000-0000-4000-8000-000000000071',version:3,name:' New name ',email:'NEW@example.invalid',role:'manager',active:true};
const reset={action:'reset_password',id:edit.id,version:3,password:'replacement password 123'};
function changeFixture({role='administrator',active=true,prepareError=null,updateError=null}={} as {role?:string;active?:boolean;prepareError?:unknown;updateError?:unknown}){
 const calls:{kind:string;args:unknown}[]=[];
 const client={auth:{getUser:async()=>({data:{user:{id:'verified-admin'}},error:null}),admin:{updateUserById:async(id:string,attributes:unknown)=>{calls.push({kind:'auth',args:{id,attributes}});return {data:{user:{id}},error:updateError};}}},
 from:()=>({select:()=>({eq:()=>({maybeSingle:async()=>({data:{role,active},error:null}),single:async()=>({data:{user_id:edit.id,version:4},error:null})})})}),
 rpc:async(name:string,args:unknown)=>{calls.push({kind:name,args});return {data:name==='console_prepare_user_change'?'operation-nonce':null,error:name==='console_prepare_user_change'?prepareError:null};}} as unknown as SupabaseClient;
 return {client,calls};
}
Deno.test('editing and password reset reject anonymous, manager and inactive admin before mutation',async()=>{
 for(const body of [edit,reset])for(const [role,active] of [['manager',true],['administrator',false]] as const){const f=changeFixture({role,active});assert((await handleRequest(request(body),f.client)).status===403);assert(f.calls.length===0);}
 const f=changeFixture();assert((await handleRequest(request(reset,false),f.client)).status===401);assert(f.calls.length===0);
});
Deno.test('profile changes use verified actor, observed version and transactional marker',async()=>{
 const f=changeFixture(),response=await handleRequest(request(edit),f.client);assert(response.status===200);
 const prepare=f.calls[0].args as Record<string,unknown>;assert(prepare.p_actor==='verified-admin'&&prepare.p_version===3&&prepare.p_email==='new@example.invalid');
 const auth=(f.calls[1].args as {attributes:Record<string,unknown>}).attributes;assert(auth.email==='new@example.invalid'&&auth.email_confirm===true&&!('password' in auth));
 assert((auth.app_metadata as Record<string,unknown>).console_operation==='operation-nonce');
});
Deno.test('reset sends password only to Auth, never to preparation or returned profile',async()=>{
 const f=changeFixture(),response=await handleRequest(request(reset),f.client);assert(response.status===200);
 assert(!JSON.stringify(f.calls[0]).includes(reset.password));
 const auth=(f.calls[1].args as {attributes:Record<string,unknown>}).attributes;assert(auth.password===reset.password&&!('email' in auth)&&!('user_metadata' in auth));
 assert(!(await response.text()).includes(reset.password));
});
Deno.test('stale or revoked preparation never reaches Auth; Auth failure cancels pending operation',async()=>{
 const stale=changeFixture({prepareError:{code:'40001'}});assert((await handleRequest(request(edit),stale.client)).status===409);assert(stale.calls.length===1);
 const fail=changeFixture({updateError:{code:'email_exists'}});assert((await handleRequest(request(edit),fail.client)).status===409);assert(fail.calls[2].kind==='console_cancel_user_change');
});
Deno.test('profile/reset forged fields are rejected before privileged actions',async()=>{
 for(const body of [{...edit,actor:'forged'},{...reset,role:'administrator'},{...reset,password:'short'}]){const f=changeFixture();assert((await handleRequest(request(body),f.client)).status===400);assert(f.calls.length===0);}
});
