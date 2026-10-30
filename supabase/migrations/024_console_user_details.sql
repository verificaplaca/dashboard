begin;
-- A short-lived server-only intent joins Auth's transaction to membership/audit.
-- Neither passwords nor password hashes are stored in this table or audit.
create table console_private.user_changes (
 token uuid primary key default gen_random_uuid(),
 actor_id uuid not null references public.console_users(user_id),
 user_id uuid not null references public.console_users(user_id),
 version integer not null,
 action text not null check(action in ('update','reset_password')),
 name text, email text, role text, active boolean,
 expires_at timestamptz not null default (clock_timestamp()+interval '2 minutes')
);
revoke all on console_private.user_changes from public,anon,authenticated,service_role;

create function public.console_prepare_user_change(p_actor uuid,p_id uuid,p_version integer,p_action text,p_name text default null,p_email text default null,p_role text default null,p_active boolean default null)
returns uuid language plpgsql security definer set search_path='' as $$
declare v_before public.console_users; v_token uuid;
begin
 perform pg_advisory_xact_lock(734221,22);
 if not exists(select 1 from public.console_users where user_id=p_actor and active and role='administrator') then
  raise exception 'Somente administradores podem gerenciar usuários.' using errcode='42501';
 end if;
 select * into v_before from public.console_users where user_id=p_id for update;
 if not found then raise exception 'Usuário não encontrado.' using errcode='22023';end if;
 if v_before.version is distinct from p_version then raise exception 'O usuário mudou. Atualize a lista antes de salvar.' using errcode='40001';end if;
 if p_action is null or p_action not in ('update','reset_password') then raise exception 'Ação inválida.' using errcode='22023';end if;
 if p_action='update' then
  if p_name is null or length(btrim(p_name)) not between 1 and 120 or p_email is null or length(p_email)>254 or p_email !~ '^[^[:space:]@]+@[^[:space:]@]+\.[^[:space:]@]+$' or p_role is null or p_role not in ('administrator','manager') or p_active is null then
   raise exception 'Dados de usuário inválidos.' using errcode='22023';
  end if;
  if v_before.role='administrator' and v_before.active and (p_role<>'administrator' or not p_active)
  and (select count(*) from public.console_users where active and role='administrator')=1 then
   raise exception 'Mantenha pelo menos um administrador ativo.' using errcode='22023';
  end if;
 end if;
 delete from console_private.user_changes where expires_at<clock_timestamp();
 insert into console_private.user_changes(actor_id,user_id,version,action,name,email,role,active)
 values(p_actor,p_id,p_version,p_action,btrim(p_name),lower(btrim(p_email)),p_role,p_active) returning token into v_token;
 return v_token;
end $$;
revoke all on function public.console_prepare_user_change(uuid,uuid,integer,text,text,text,text,boolean) from public,anon,authenticated;
grant execute on function public.console_prepare_user_change(uuid,uuid,integer,text,text,text,text,boolean) to service_role;

create function public.console_cancel_user_change(p_token uuid) returns void language sql security definer set search_path='' as $$
 delete from console_private.user_changes where token=p_token
$$;
revoke all on function public.console_cancel_user_change(uuid) from public,anon,authenticated;
grant execute on function public.console_cancel_user_change(uuid) to service_role;

-- GoTrue's admin update writes password/email/app metadata in ONE transaction.
-- Any failure here rolls back Auth, membership and audit together.
create function console_private.finish_user_change() returns trigger language plpgsql security definer set search_path='' as $$
declare v_change console_private.user_changes; v_before public.console_users;v_after public.console_users;v_marker text;
begin
 v_marker:=new.raw_app_meta_data->>'console_operation';
 if v_marker is null then return new;end if;
 perform pg_advisory_xact_lock(734221,22);
 select * into v_change from console_private.user_changes where token::text=v_marker and user_id=new.id for update;
 if not found or v_change.expires_at<clock_timestamp() then raise exception 'Alteração expirada ou inválida.' using errcode='42501';end if;
 if not exists(select 1 from public.console_users where user_id=v_change.actor_id and active and role='administrator') then
  raise exception 'Permissão de administrador revogada.' using errcode='42501';end if;
 select * into v_before from public.console_users where user_id=new.id for update;
 if v_before.version is distinct from v_change.version then raise exception 'O usuário mudou. Atualize a lista antes de salvar.' using errcode='40001';end if;
 if v_change.action='update' then
  if new.email is distinct from v_change.email then raise exception 'Email não confirmado na alteração.' using errcode='22023';end if;
  if v_before.role='administrator' and v_before.active and (v_change.role<>'administrator' or not v_change.active)
  and (select count(*) from public.console_users where active and role='administrator')=1 then
   raise exception 'Mantenha pelo menos um administrador ativo.' using errcode='22023';end if;
  new.raw_user_meta_data:=coalesce(new.raw_user_meta_data,'{}'::jsonb)||jsonb_build_object('full_name',v_change.name);
  update public.console_users set display_name=v_change.name,email=v_change.email,role=v_change.role,active=v_change.active,version=version+1,updated_at=clock_timestamp() where user_id=new.id returning * into v_after;
 else
  update public.console_users set version=version+1,updated_at=clock_timestamp() where user_id=new.id returning * into v_after;
 end if;
 insert into public.console_user_events(actor_id,user_id,before_value,after_value)
 values(v_change.actor_id,new.id,to_jsonb(v_before),to_jsonb(v_after)||jsonb_build_object('action',v_change.action));
 delete from console_private.user_changes where token=v_change.token;
 new.raw_app_meta_data:=new.raw_app_meta_data-'console_operation';
 return new;
end $$;
revoke all on function console_private.finish_user_change() from public,anon,authenticated,service_role;
create trigger console_finish_user_change before update of raw_app_meta_data on auth.users for each row execute function console_private.finish_user_change();
notify pgrst,'reload schema';
commit;
