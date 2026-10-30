begin;
create table console_private.recovery_emails (
 id uuid primary key default gen_random_uuid(),
 actor_id uuid not null references public.console_users(user_id),
 user_id uuid not null references public.console_users(user_id),
 email text not null,
 status text not null default 'pending' check(status in ('pending','sent','failed')),
 created_at timestamptz not null default clock_timestamp()
);
create index recovery_emails_user_time on console_private.recovery_emails(user_id,created_at desc);
revoke all on console_private.recovery_emails from public,anon,authenticated,service_role;
create function public.console_request_recovery_email(p_actor uuid,p_id uuid,p_version integer)
returns jsonb language plpgsql security definer set search_path='' as $$
declare v_user public.console_users;v_id uuid;
begin
 perform pg_advisory_xact_lock(734221,22);
 if not exists(select 1 from public.console_users where user_id=p_actor and active and role='administrator') then
  raise exception 'Somente administradores podem enviar links de redefinição.' using errcode='42501';end if;
 select * into v_user from public.console_users where user_id=p_id for update;
 if not found or not v_user.active then raise exception 'O usuário precisa estar ativo para receber o link.' using errcode='22023';end if;
 if v_user.version is distinct from p_version then raise exception 'O usuário mudou. Atualize a lista antes de enviar.' using errcode='40001';end if;
 if exists(select 1 from console_private.recovery_emails where user_id=p_id and created_at>clock_timestamp()-interval '60 seconds')
 or (select count(*) from console_private.recovery_emails where user_id=p_id and created_at>clock_timestamp()-interval '1 hour')>=5 then
  raise exception 'Aguarde antes de enviar outro link para este usuário.' using errcode='P0001';end if;
 insert into console_private.recovery_emails(actor_id,user_id,email) values(p_actor,p_id,v_user.email) returning id into v_id;
 return jsonb_build_object('request_id',v_id,'email',v_user.email);
end $$;
revoke all on function public.console_request_recovery_email(uuid,uuid,integer) from public,anon,authenticated;
grant execute on function public.console_request_recovery_email(uuid,uuid,integer) to service_role;
create function public.console_finish_recovery_email(p_id uuid,p_sent boolean) returns void language plpgsql security definer set search_path='' as $$
declare v_request console_private.recovery_emails;
begin
 update console_private.recovery_emails set status=case when p_sent then 'sent' else 'failed' end where id=p_id and status='pending' returning * into v_request;
 if found and p_sent then
  insert into public.console_user_events(actor_id,user_id,after_value)
  values(v_request.actor_id,v_request.user_id,jsonb_build_object('action','recovery_email_sent','email',v_request.email,'request_id',v_request.id));
 end if;
end $$;
revoke all on function public.console_finish_recovery_email(uuid,boolean) from public,anon,authenticated;
grant execute on function public.console_finish_recovery_email(uuid,boolean) to service_role;
notify pgrst,'reload schema';
commit;
