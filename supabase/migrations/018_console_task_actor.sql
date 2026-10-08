-- Bind each mutation to the identity the UI verified before submitting.
-- Implementation functions are no longer reachable through the public API.
begin;
alter function public.console_create_task(jsonb) set schema console_private;
alter function public.console_update_task(uuid,integer,jsonb) set schema console_private;
alter function public.console_add_task_comment(uuid,text) set schema console_private;
revoke all on function console_private.console_create_task(jsonb),console_private.console_update_task(uuid,integer,jsonb),console_private.console_add_task_comment(uuid,text) from public,anon,authenticated;
create function public.console_create_task(p_task jsonb,p_actor uuid) returns public.console_tasks
language plpgsql security definer set search_path='' as $$ begin
  if p_actor is null or p_actor is distinct from auth.uid() then raise exception 'A sessão mudou. Entre novamente e atualize.' using errcode='42501'; end if;
  return console_private.console_create_task(p_task);
end $$;
create function public.console_update_task(p_id uuid,p_version integer,p_changes jsonb,p_actor uuid) returns public.console_tasks
language plpgsql security definer set search_path='' as $$ begin
  if p_actor is null or p_actor is distinct from auth.uid() then raise exception 'A sessão mudou. Entre novamente e atualize.' using errcode='42501'; end if;
  return console_private.console_update_task(p_id,p_version,p_changes);
end $$;
create function public.console_add_task_comment(p_id uuid,p_body text,p_actor uuid) returns public.console_task_events
language plpgsql security definer set search_path='' as $$ begin
  if p_actor is null or p_actor is distinct from auth.uid() then raise exception 'A sessão mudou. Entre novamente e atualize.' using errcode='42501'; end if;
  return console_private.console_add_task_comment(p_id,p_body);
end $$;
revoke all on function public.console_create_task(jsonb,uuid),public.console_update_task(uuid,integer,jsonb,uuid),public.console_add_task_comment(uuid,text,uuid) from public,anon;
grant execute on function public.console_create_task(jsonb,uuid),public.console_update_task(uuid,integer,jsonb,uuid),public.console_add_task_comment(uuid,text,uuid) to authenticated;
notify pgrst,'reload schema';
commit;
