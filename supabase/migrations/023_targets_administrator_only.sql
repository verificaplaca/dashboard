begin;
-- Reading is unchanged; every browser/client mutation requires a live admin.
alter policy console_targets_insert_gate on public.monthly_targets
 with check ((select console_private.console_role()) = 'administrator');
alter policy console_targets_update_gate on public.monthly_targets
 using ((select console_private.console_role()) = 'administrator')
 with check ((select console_private.console_role()) = 'administrator');
create policy console_targets_delete_gate on public.monthly_targets as restrictive
 for delete to authenticated using ((select console_private.console_role()) = 'administrator');
notify pgrst,'reload schema';
commit;
