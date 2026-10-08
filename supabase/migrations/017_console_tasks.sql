-- Additive task module, principal project ftmgmfdqdqxboiktxcoj.
-- No analytics, Auth policies, syncs or monthly targets are changed.
begin;
create schema console_private;
revoke all on schema console_private from public, anon;
grant usage on schema console_private to authenticated;

create table public.console_task_members (
  user_id uuid primary key references auth.users(id),
  display_name text not null check (length(btrim(display_name)) between 1 and 160),
  role text not null default 'member' check (role in ('manager','member')),
  active boolean not null default true
);
alter table public.console_task_members enable row level security;
revoke all on public.console_task_members from public, anon, authenticated;
grant select on public.console_task_members to authenticated;
grant all on public.console_task_members to service_role;

-- Explicit bootstrap for the single operator already managing monthly targets.
-- A fresh/multi-user installation must provision membership deliberately.
do $$ begin
  if (select count(*) from auth.users where deleted_at is null) <> 1 then
    raise exception 'Expected one existing operator. Review task membership before applying.';
  end if;
  insert into public.console_task_members(user_id,display_name,role)
  select id, left(coalesce(nullif(btrim(raw_user_meta_data->>'full_name'),''),'Operador Verifica Placa'),160), 'manager'
  from auth.users where deleted_at is null;
end $$;

create function console_private.task_role() returns text
language sql stable security definer set search_path = ''
as $$ select role from public.console_task_members where user_id = (select auth.uid()) and active $$;
revoke all on function console_private.task_role() from public, anon;
grant execute on function console_private.task_role() to authenticated;
create policy task_directory on public.console_task_members for select to authenticated
using ((select console_private.task_role()) is not null and (active or user_id = (select auth.uid())));

create table public.console_tasks (
  id uuid primary key default gen_random_uuid(),
  title text not null check (length(btrim(title)) between 2 and 160),
  description text not null default '' check (length(description) <= 6000),
  status text not null default 'todo' check (status in ('todo','doing','blocked','done')),
  priority text not null default 'normal' check (priority in ('low','normal','high')),
  assignee_id uuid references public.console_task_members(user_id),
  created_by uuid not null references public.console_task_members(user_id),
  due_date date check (due_date between date '2000-01-01' and date '2200-12-31'),
  completed_at timestamptz,
  archived_at timestamptz,
  version integer not null default 1 check (version > 0),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check ((status = 'done') = (completed_at is not null))
);
create index console_tasks_assignee on public.console_tasks(assignee_id,created_at desc);
create index console_tasks_creator on public.console_tasks(created_by,created_at desc);
create index console_tasks_open_due on public.console_tasks(due_date) where archived_at is null and status <> 'done';
create index console_tasks_created on public.console_tasks(created_at desc,id);
alter table public.console_tasks enable row level security;
revoke all on public.console_tasks from public, anon, authenticated;
grant select on public.console_tasks to authenticated;
grant all on public.console_tasks to service_role;
create policy task_visibility on public.console_tasks for select to authenticated using (
  (select console_private.task_role()) is not null and (
    (select console_private.task_role()) = 'manager' or created_by = (select auth.uid()) or assignee_id = (select auth.uid())
  )
);

create table public.console_task_events (
  id uuid primary key default gen_random_uuid(),
  task_id uuid not null references public.console_tasks(id),
  actor_id uuid not null references public.console_task_members(user_id),
  actor_name text not null,
  kind text not null check (kind in ('created','updated','comment')),
  changes jsonb not null default '{}'::jsonb,
  body text not null default '' check (length(body) <= 2000),
  task_version integer not null check (task_version > 0),
  created_at timestamptz not null default clock_timestamp()
);
create index console_task_events_task on public.console_task_events(task_id,created_at desc,id);
alter table public.console_task_events enable row level security;
revoke all on public.console_task_events from public, anon, authenticated;
grant select on public.console_task_events to authenticated;
grant all on public.console_task_events to service_role;
create policy task_event_visibility on public.console_task_events for select to authenticated
using (exists (select 1 from public.console_tasks t where t.id = task_id));

create function public.console_create_task(p_task jsonb) returns public.console_tasks
language plpgsql security definer set search_path = '' as $$
declare v_uid uuid := auth.uid(); v_role text; v_row public.console_tasks;
begin
  -- Hold membership against concurrent revocation until this transaction ends.
  select role into v_role from public.console_task_members where user_id=v_uid and active for share;
  if v_uid is null or v_role is null then raise exception 'Sem acesso ao módulo tarefas.' using errcode = '42501'; end if;
  if p_task is null or jsonb_typeof(p_task) <> 'object' or p_task - array['id','title','description','priority','assignee_id','due_date'] <> '{}'::jsonb then
    raise exception 'Campos inválidos.' using errcode = '22023';
  end if;
  if nullif(p_task->>'assignee_id','') is not null and not exists (
    select 1 from public.console_task_members where user_id = (p_task->>'assignee_id')::uuid and active for share
  ) then raise exception 'Responsável indisponível.' using errcode = '22023'; end if;
  if v_role <> 'manager' and nullif(p_task->>'assignee_id','') is null then
    raise exception 'Escolha um responsável ativo.' using errcode = '22023';
  end if;
  insert into public.console_tasks(id,title,description,priority,assignee_id,due_date,created_by)
  values (coalesce((p_task->>'id')::uuid,gen_random_uuid()), btrim(p_task->>'title'),coalesce(p_task->>'description',''),
    coalesce(p_task->>'priority','normal'),nullif(p_task->>'assignee_id','')::uuid,nullif(p_task->>'due_date','')::date,v_uid)
  on conflict(id) do nothing returning * into v_row;
  if not found then
    select * into v_row from public.console_tasks where id=(p_task->>'id')::uuid;
    if v_row.created_by=v_uid and v_row.title=btrim(p_task->>'title') and v_row.description=coalesce(p_task->>'description','')
      and v_row.priority=coalesce(p_task->>'priority','normal') and v_row.assignee_id is not distinct from nullif(p_task->>'assignee_id','')::uuid
      and v_row.due_date is not distinct from nullif(p_task->>'due_date','')::date then return v_row; end if;
    raise exception 'Esta tarefa já foi registrada. Atualize antes de tentar novamente.' using errcode='23505';
  end if;
  insert into public.console_task_events(task_id,actor_id,actor_name,kind,task_version)
  select v_row.id,v_uid,display_name,'created',v_row.version from public.console_task_members where user_id=v_uid;
  return v_row;
end $$;

create function public.console_update_task(p_id uuid,p_version integer,p_changes jsonb) returns public.console_tasks
language plpgsql security definer set search_path = '' as $$
declare v_uid uuid := auth.uid(); v_role text; v_before public.console_tasks; v_row public.console_tasks; v_changes jsonb;
begin
  -- Hold membership against concurrent revocation until this transaction ends.
  select role into v_role from public.console_task_members where user_id=v_uid and active for share;
  if v_uid is null or v_role is null then raise exception 'Sem acesso ao módulo tarefas.' using errcode = '42501'; end if;
  select * into v_before from public.console_tasks where id=p_id for update;
  if not found or not (v_role='manager' or v_before.created_by=v_uid or v_before.assignee_id=v_uid) then
    raise exception 'Tarefa indisponível para esta conta.' using errcode = '42501';
  end if;
  if p_version is distinct from v_before.version then raise exception 'A tarefa mudou. Atualize antes de salvar.' using errcode = '40001'; end if;
  if p_changes is null or jsonb_typeof(p_changes) <> 'object' or p_changes='{}'::jsonb or
    p_changes - array['title','description','priority','assignee_id','due_date','status','archived'] <> '{}'::jsonb then
    raise exception 'Campos inválidos.' using errcode = '22023';
  end if;
  -- Assignees can advance execution and comment; creator/manager manages scope.
  if v_role <> 'manager' and v_before.created_by <> v_uid and p_changes - 'status' <> '{}'::jsonb then
    raise exception 'Somente o criador ou gestor pode editar os detalhes.' using errcode = '42501';
  end if;
  if v_before.archived_at is not null and p_changes <> '{"archived":false}'::jsonb then
    raise exception 'Restaure a tarefa antes de alterar.' using errcode = '22023';
  end if;
  if p_changes ? 'archived' and (jsonb_typeof(p_changes->'archived') <> 'boolean' or p_changes - 'archived' <> '{}'::jsonb) then
    raise exception 'Arquivamento deve ser uma ação separada.' using errcode = '22023';
  end if;
  if p_changes ? 'assignee_id' and nullif(p_changes->>'assignee_id','') is not null and not exists (
    select 1 from public.console_task_members where user_id=(p_changes->>'assignee_id')::uuid and active for share
  ) then raise exception 'Responsável indisponível.' using errcode = '22023'; end if;
  if v_role <> 'manager' and p_changes ? 'assignee_id' and nullif(p_changes->>'assignee_id','') is null then
    raise exception 'Escolha um responsável ativo.' using errcode = '22023';
  end if;
  update public.console_tasks set
    title=case when p_changes ? 'title' then btrim(p_changes->>'title') else title end,
    description=case when p_changes ? 'description' then p_changes->>'description' else description end,
    priority=case when p_changes ? 'priority' then p_changes->>'priority' else priority end,
    assignee_id=case when p_changes ? 'assignee_id' then nullif(p_changes->>'assignee_id','')::uuid else assignee_id end,
    due_date=case when p_changes ? 'due_date' then nullif(p_changes->>'due_date','')::date else due_date end,
    status=case when p_changes ? 'status' then p_changes->>'status' else status end,
    completed_at=case when p_changes ? 'status' then case when p_changes->>'status'='done' then coalesce(completed_at,now()) else null end else completed_at end,
    archived_at=case when p_changes ? 'archived' then case when (p_changes->>'archived')::boolean then now() else null end else archived_at end,
    updated_at=clock_timestamp(), version=version+1 where id=p_id returning * into v_row;
  select coalesce(jsonb_object_agg(key,jsonb_build_object('before',to_jsonb(v_before)->key,'after',to_jsonb(v_row)->key)),'{}'::jsonb)
  into v_changes from jsonb_object_keys(to_jsonb(v_row)) as key
  where key in ('title','description','priority','assignee_id','due_date','status','archived_at') and to_jsonb(v_before)->key is distinct from to_jsonb(v_row)->key;
  insert into public.console_task_events(task_id,actor_id,actor_name,kind,changes,task_version)
  select p_id,v_uid,display_name,'updated',v_changes,v_row.version from public.console_task_members where user_id=v_uid;
  return v_row;
end $$;

create function public.console_add_task_comment(p_id uuid,p_body text) returns public.console_task_events
language plpgsql security definer set search_path = '' as $$
declare v_uid uuid := auth.uid(); v_role text; v_task public.console_tasks; v_row public.console_task_events;
begin
  -- Hold membership against concurrent revocation until this transaction ends.
  select role into v_role from public.console_task_members where user_id=v_uid and active for share;
  if v_uid is null or v_role is null then raise exception 'Sem acesso ao módulo tarefas.' using errcode = '42501'; end if;
  select * into v_task from public.console_tasks where id=p_id for update;
  if not found or not (v_role='manager' or v_task.created_by=v_uid or v_task.assignee_id=v_uid) then
    raise exception 'Tarefa indisponível para esta conta.' using errcode = '42501';
  end if;
  if v_task.archived_at is not null then raise exception 'Restaure a tarefa antes de comentar.' using errcode = '22023'; end if;
  if p_body is null or length(btrim(p_body)) not between 1 and 2000 then raise exception 'Comentário inválido.' using errcode = '22023'; end if;
  insert into public.console_task_events(task_id,actor_id,actor_name,kind,body,task_version)
  select p_id,v_uid,display_name,'comment',btrim(p_body),v_task.version from public.console_task_members where user_id=v_uid returning * into v_row;
  return v_row;
end $$;

-- Invoker preserves RLS; date-only deadlines follow the São Paulo calendar.
create function public.console_task_summary() returns jsonb language sql stable security invoker set search_path = '' as $$
select jsonb_build_object(
  'open',count(*) filter (where archived_at is null and status <> 'done'),
  'mine',count(*) filter (where archived_at is null and status <> 'done' and assignee_id=(select auth.uid())),
  'overdue',count(*) filter (where archived_at is null and status <> 'done' and due_date < (now() at time zone 'America/Sao_Paulo')::date),
  'today',count(*) filter (where archived_at is null and status <> 'done' and due_date = (now() at time zone 'America/Sao_Paulo')::date),
  'done',count(*) filter (where archived_at is null and status='done')
) from public.console_tasks $$;
revoke all on function public.console_create_task(jsonb),public.console_update_task(uuid,integer,jsonb),public.console_add_task_comment(uuid,text),public.console_task_summary() from public,anon;
grant execute on function public.console_create_task(jsonb),public.console_update_task(uuid,integer,jsonb),public.console_add_task_comment(uuid,text),public.console_task_summary() to authenticated;
notify pgrst, 'reload schema';
commit;
