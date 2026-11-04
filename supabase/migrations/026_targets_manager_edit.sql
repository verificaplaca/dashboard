begin;
-- Active managers may create and update business targets, without tax changes.
alter policy console_targets_insert_gate on public.monthly_targets
 with check ((select console_private.console_role()) in ('administrator','manager'));
alter policy console_targets_update_gate on public.monthly_targets
 using ((select console_private.console_role()) in ('administrator','manager'))
 with check ((select console_private.console_role()) in ('administrator','manager'));
-- Deletion remains administrator-only (migration 023).
create function console_private.guard_manager_target_tax() returns trigger
language plpgsql security definer set search_path='' as $$
begin
 if console_private.console_role() = 'manager' then
  if tg_op = 'INSERT' then
   if new.tax_pct is not null then
    raise exception 'Somente administradores podem alterar o imposto.' using errcode='42501';
   end if;
  elsif new.tax_pct is distinct from old.tax_pct or new.month is distinct from old.month then
   raise exception 'Somente administradores podem alterar o imposto ou mover um mês cadastrado.' using errcode='42501';
  end if;
 end if;
 return new;
end $$;
revoke all on function console_private.guard_manager_target_tax() from public,anon,authenticated;
create trigger console_targets_manager_tax before insert or update on public.monthly_targets
for each row execute function console_private.guard_manager_target_tax();
notify pgrst,'reload schema';
commit;
