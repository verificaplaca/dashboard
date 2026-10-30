-- Console identity is managed here, never in editable Auth metadata.
begin;
create table public.console_users (
 user_id uuid primary key references auth.users(id),
 display_name text not null check(length(btrim(display_name)) between 1 and 120),
 email text not null,
 role text not null check(role in ('administrator','manager')),
 active boolean not null default true,
 version integer not null default 1,
 created_at timestamptz not null default now(),
 updated_at timestamptz not null default now()
);
alter table public.console_users enable row level security;
revoke all on public.console_users from public,anon,authenticated;
grant select on public.console_users to authenticated;
grant all on public.console_users to service_role;
create function console_private.console_role() returns text language sql stable security definer set search_path='' as $$
 select role from public.console_users where user_id=auth.uid() and active
$$;
revoke all on function console_private.console_role() from public,anon;
grant execute on function console_private.console_role() to authenticated;
create policy console_users_read on public.console_users for select to authenticated
using(user_id=auth.uid() or (select console_private.console_role())='administrator');
create table public.console_user_events (
 id bigint generated always as identity primary key,
 actor_id uuid not null references auth.users(id),
 user_id uuid not null references auth.users(id),
 before_value jsonb,
 after_value jsonb not null,
 created_at timestamptz not null default now()
);
alter table public.console_user_events enable row level security;
revoke all on public.console_user_events from public,anon,authenticated;
grant select on public.console_user_events to authenticated;
grant all on public.console_user_events to service_role;
create policy console_user_events_read on public.console_user_events for select to authenticated
using((select console_private.console_role())='administrator');

-- Keep the existing task and finance contracts; revocation applies to both.
create function console_private.sync_console_membership() returns trigger language plpgsql security definer set search_path='' as $$
begin
 insert into public.console_task_members(user_id,display_name,role,active)
 values(new.user_id,new.display_name,'manager',new.active)
 on conflict(user_id) do update set display_name=excluded.display_name,role=excluded.role,active=excluded.active;
 return new;
end $$;
revoke all on function console_private.sync_console_membership() from public,anon,authenticated;
create trigger console_membership after insert or update on public.console_users
for each row execute function console_private.sync_console_membership();

-- Only the single pre-existing operator may be bootstrapped automatically.
do $$ begin
 if (select count(*) from auth.users where deleted_at is null)<>1
 or (select count(*) from public.console_task_members where active and role='manager')<>1 then
  raise exception 'Bootstrap requires one existing operator. Explicitly review administrator identity.';
 end if;
 insert into public.console_users(user_id,display_name,email,role)
 select u.id,left(m.display_name,120),u.email,'administrator' from auth.users u
 join public.console_task_members m on m.user_id=u.id where u.deleted_at is null and m.active and m.role='manager';
 if not found then raise exception 'Existing operator and membership do not match.'; end if;
end $$;

-- A valid Auth session alone does not grant access to private modules.
create or replace function console_private.task_role() returns text language sql stable security definer set search_path='' as $$
 select m.role from public.console_task_members m join public.console_users u on u.user_id=m.user_id
 where m.user_id=auth.uid() and m.active and u.active
$$;
-- Restrict the former authenticated-write policy without changing public aggregates.
create policy console_targets_insert_gate on public.monthly_targets as restrictive for insert to authenticated
with check((select console_private.console_role()) is not null);
create policy console_targets_update_gate on public.monthly_targets as restrictive for update to authenticated
using((select console_private.console_role()) is not null) with check((select console_private.console_role()) is not null);

-- One serialization lock protects concurrent last-admin checks and role changes.
create function public.console_update_user(p_id uuid,p_role text,p_active boolean,p_version integer)
returns public.console_users language plpgsql security definer set search_path='' as $$
declare v_before public.console_users; v_after public.console_users;
begin
 perform pg_advisory_xact_lock(734221,22);
 if console_private.console_role() is distinct from 'administrator' then raise exception 'Somente administradores podem gerenciar usuários.' using errcode='42501'; end if;
 select * into v_before from public.console_users where user_id=p_id for update;
 if not found then raise exception 'Usuário não encontrado.' using errcode='22023'; end if;
 if p_version is distinct from v_before.version then raise exception 'O usuário mudou. Atualize a lista antes de salvar.' using errcode='40001'; end if;
 if p_role is null or p_role not in ('administrator','manager') or p_active is null then raise exception 'Perfil ou situação inválidos.' using errcode='22023'; end if;
 if v_before.role='administrator' and v_before.active and (p_role<>'administrator' or not p_active)
 and (select count(*) from public.console_users where role='administrator' and active)=1 then
  raise exception 'Mantenha pelo menos um administrador ativo.' using errcode='22023';
 end if;
 update public.console_users set role=p_role,active=p_active,version=version+1,updated_at=clock_timestamp() where user_id=p_id returning * into v_after;
 insert into public.console_user_events(actor_id,user_id,before_value,after_value) values(auth.uid(),p_id,to_jsonb(v_before),to_jsonb(v_after));
 return v_after;
end $$;
revoke all on function public.console_update_user(uuid,text,boolean,integer) from public,anon;
grant execute on function public.console_update_user(uuid,text,boolean,integer) to authenticated;

-- Service-only finalization, after Auth creation. Recheck the verified caller in
-- the same transaction as membership + audit; rolled-back Auth accounts have no access.
create function public.console_provision_user(p_actor uuid,p_id uuid,p_name text,p_role text)
returns public.console_users language plpgsql security definer set search_path='' as $$
declare v_row public.console_users; v_email text;
begin
 perform pg_advisory_xact_lock(734221,22);
 if not exists(select 1 from public.console_users where user_id=p_actor and active and role='administrator') then
  raise exception 'Somente administradores podem criar usuários.' using errcode='42501';
 end if;
 select email into v_email from auth.users where id=p_id and deleted_at is null;
 if v_email is null then raise exception 'Conta inválida.' using errcode='22023'; end if;
 insert into public.console_users(user_id,display_name,email,role) values(p_id,btrim(p_name),v_email,p_role) returning * into v_row;
 insert into public.console_user_events(actor_id,user_id,after_value) values(p_actor,p_id,to_jsonb(v_row));
 return v_row;
end $$;
revoke all on function public.console_provision_user(uuid,uuid,text,text) from public,anon,authenticated;
grant execute on function public.console_provision_user(uuid,uuid,text,text) to service_role;
notify pgrst,'reload schema';
commit;
