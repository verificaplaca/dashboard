import { createClient, type SupabaseClient } from 'npm:@supabase/supabase-js@2.117.3';
import { validateNewUser } from './validation.ts';
const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-client-info', 'Access-Control-Allow-Methods': 'POST, OPTIONS' };
export async function handleRequest(request: Request, admin: SupabaseClient) {
 const respond = (status: number, body: unknown) => new Response(JSON.stringify(body), { status, headers: { ...cors, 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } });
 if(request.method==='OPTIONS') return new Response(null,{status:204,headers:cors});
 if(request.method!=='POST') return respond(405,{error:'Método não permitido.'});
 const token=request.headers.get('Authorization')?.match(/^Bearer (.+)$/i)?.[1];
 if(!token) return respond(401,{error:'Entre novamente para continuar.'});
 try {
  const {data:identity,error:authError}=await admin.auth.getUser(token);
  if(authError||!identity.user) return respond(401,{error:'Sessão inválida. Entre novamente.'});
  const {data:actor,error:accessError}=await admin.from('console_users').select('role,active').eq('user_id',identity.user.id).maybeSingle();
  if(accessError) return respond(503,{error:'Não foi possível verificar a permissão. Tente novamente.'});
  if(!actor?.active||actor.role!=='administrator') return respond(403,{error:'Somente administradores podem criar usuários.'});
  const raw=await request.text();
  if(raw.length>4096) return respond(400,{error:'Formulário muito longo.'});
  let values;
  try { values=validateNewUser(JSON.parse(raw)); } catch(error) { return respond(400,{error:error instanceof Error?error.message:'Formulário inválido.'}); }
  const {data:created,error:createError}=await admin.auth.admin.createUser({email:values.email,password:values.password,email_confirm:true,user_metadata:{full_name:values.name}});
  if(createError||!created.user) return respond(createError?.code==='email_exists'?409:400,{error:createError?.code==='email_exists'?'Este email já tem uma conta. Atualize a lista de usuários.':'Não foi possível criar a conta. Confira o email e os requisitos de senha.'});
  const {data:member,error:provisionError}=await admin.rpc('console_provision_user',{p_actor:identity.user.id,p_id:created.user.id,p_name:values.name,p_role:values.role});
  if(provisionError) {
   const {error:cleanupError}=await admin.auth.admin.deleteUser(created.user.id);
   if(cleanupError) console.error('console-users: Auth cleanup required for',created.user.id);
   return respond(409,{error:'Não foi possível confirmar o acesso. Atualize a lista antes de tentar novamente.'});
  }
  return respond(201,{user:member});
 } catch { return respond(503,{error:'Serviço indisponível. Atualize a lista antes de tentar novamente.'}); }
}
if(import.meta.main) Deno.serve(request => handleRequest(request,createClient(Deno.env.get('SUPABASE_URL')!,Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,{auth:{persistSession:false,autoRefreshToken:false}})));
