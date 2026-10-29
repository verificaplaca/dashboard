-- Private monthly DRE complements. Existing financial sources remain read-only.
create table public.console_finance_months (
 month date primary key check (extract(day from month)=1 and month between date '2020-01-01' and date '2100-12-01'),
 fees numeric(14,2) check(fees between 0 and 1000000000),
 people numeric(14,2) check(people between 0 and 1000000000),
 systems numeric(14,2) check(systems between 0 and 1000000000),
 other numeric(14,2) check(other between 0 and 1000000000),
 financial numeric(14,2) check(financial between 0 and 1000000000),
 version integer not null default 1,
 updated_at timestamptz not null default clock_timestamp(),
 updated_by uuid not null references auth.users(id)
);
create table public.console_finance_history (
 id bigint generated always as identity primary key,
 month date not null references public.console_finance_months(month),
 version integer not null,
 amounts jsonb not null,
 actor_id uuid not null references auth.users(id),
 recorded_at timestamptz not null default clock_timestamp(),
 unique(month,version)
);
alter table public.console_finance_months enable row level security;
alter table public.console_finance_history enable row level security;
revoke all on public.console_finance_months,public.console_finance_history from public,anon,authenticated;
grant select on public.console_finance_months,public.console_finance_history to authenticated;
grant all on public.console_finance_months,public.console_finance_history to service_role;
create policy finance_manager_read on public.console_finance_months for select to authenticated using ((select console_private.task_role())='manager');
create policy finance_history_manager_read on public.console_finance_history for select to authenticated using ((select console_private.task_role())='manager');
create function public.console_finance_access() returns boolean language sql stable security definer set search_path='' as $$
 select coalesce(console_private.task_role()='manager',false)
$$;
revoke all on function public.console_finance_access() from public,anon;
grant execute on function public.console_finance_access() to authenticated;
create function public.console_save_finance_month(p_month date,p_values jsonb,p_version integer) returns public.console_finance_months
language plpgsql security definer set search_path='' as $$
declare v_uid uuid:=auth.uid(); v_role text; v_key text; v_value jsonb; v_row public.console_finance_months;
begin
 select role into v_role from public.console_task_members where user_id=v_uid and active for share;
 if v_uid is null or v_role is distinct from 'manager' then raise exception 'Somente o gestor autorizado pode salvar despesas.' using errcode='42501'; end if;
 if p_month is null or extract(day from p_month)<>1 or p_month not between date '2020-01-01' and date '2100-12-01' or p_version is null or p_version<0 or p_values is null or jsonb_typeof(p_values)<>'object' or not p_values ?& array['fees','people','systems','other','financial'] or p_values-array['fees','people','systems','other','financial']<>'{}'::jsonb then
  raise exception 'Confira mês, versão e categorias da DRE.' using errcode='22023';
 end if;
 for v_key,v_value in select * from jsonb_each(p_values) loop
  if v_value<>'null'::jsonb and (jsonb_typeof(v_value)<>'number') then raise exception 'Valores devem ser números ou ficar em aberto.' using errcode='22023'; end if;
  if v_value<>'null'::jsonb and ((v_value::text)::numeric<0 or (v_value::text)::numeric>1000000000 or round((v_value::text)::numeric,2)<>(v_value::text)::numeric) then raise exception 'Informe valores não negativos com até duas casas decimais.' using errcode='22023'; end if;
 end loop;
 if p_version=0 then
  begin
   insert into public.console_finance_months(month,fees,people,systems,other,financial,updated_by)
   values(p_month,(p_values->>'fees')::numeric,(p_values->>'people')::numeric,(p_values->>'systems')::numeric,(p_values->>'other')::numeric,(p_values->>'financial')::numeric,v_uid) returning * into v_row;
  exception when unique_violation then raise exception 'O mês foi alterado por outra pessoa. Recarregue antes de salvar.' using errcode='40001'; end;
 else
  update public.console_finance_months set fees=(p_values->>'fees')::numeric,people=(p_values->>'people')::numeric,systems=(p_values->>'systems')::numeric,other=(p_values->>'other')::numeric,financial=(p_values->>'financial')::numeric,version=version+1,updated_at=clock_timestamp(),updated_by=v_uid
  where month=p_month and version=p_version returning * into v_row;
  if not found then raise exception 'O mês foi alterado. Recarregue antes de salvar.' using errcode='40001'; end if;
 end if;
 insert into public.console_finance_history(month,version,amounts,actor_id) values(p_month,v_row.version,p_values,v_uid);
 return v_row;
end $$;
revoke all on function public.console_save_finance_month(date,jsonb,integer) from public,anon;
grant execute on function public.console_save_finance_month(date,jsonb,integer) to authenticated;
