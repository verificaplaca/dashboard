begin;
insert into auth.users(id,email) values
 ('90000000-0000-4000-8000-000000000061','recovery-admin@example.invalid'),
 ('90000000-0000-4000-8000-000000000062','recovery-manager@example.invalid');
insert into public.console_users(user_id,display_name,email,role) values
 ('90000000-0000-4000-8000-000000000061','Recovery admin','recovery-admin@example.invalid','administrator'),
 ('90000000-0000-4000-8000-000000000062','Recovery manager','recovery-manager@example.invalid','manager');
set local role authenticated;
do $$begin
 begin perform public.console_request_recovery_email('90000000-0000-4000-8000-000000000061','90000000-0000-4000-8000-000000000062',1);raise exception 'Public recovery RPC';exception when insufficient_privilege then null;end;
 begin perform * from console_private.recovery_emails;raise exception 'Public request ledger';exception when insufficient_privilege then null;end;
end$$;
reset role;
set local role service_role;
do $$begin
 begin perform public.console_request_recovery_email('90000000-0000-4000-8000-000000000062','90000000-0000-4000-8000-000000000061',1);raise exception 'Manager sent recovery';exception when insufficient_privilege then null;end;
 begin perform public.console_request_recovery_email('90000000-0000-4000-8000-000000000061','90000000-0000-4000-8000-000000000062',9);raise exception 'Stale email request';exception when serialization_failure then null;end;
end$$;
select set_config('test.recovery',public.console_request_recovery_email('90000000-0000-4000-8000-000000000061','90000000-0000-4000-8000-000000000062',1)::text,true);
do $$begin
 assert current_setting('test.recovery')::jsonb->>'email'='recovery-manager@example.invalid';
 begin perform public.console_request_recovery_email('90000000-0000-4000-8000-000000000061','90000000-0000-4000-8000-000000000062',1);raise exception 'Cooldown bypass';exception when raise_exception then assert sqlerrm like 'Aguarde%';end;
end$$;
select public.console_finish_recovery_email((current_setting('test.recovery')::jsonb->>'request_id')::uuid,true);
select public.console_finish_recovery_email((current_setting('test.recovery')::jsonb->>'request_id')::uuid,true);
reset role;
do $$begin
 assert (select count(*) from public.console_user_events where user_id='90000000-0000-4000-8000-000000000062' and after_value->>'action'='recovery_email_sent')=1,'No duplicate audit';
 assert (select version=1 from public.console_users where user_id='90000000-0000-4000-8000-000000000062'),'Sending does not alter user';
end$$;
update console_private.recovery_emails set created_at=clock_timestamp()-interval '61 seconds' where user_id='90000000-0000-4000-8000-000000000062';
set local role service_role;
select set_config('test.recovery',public.console_request_recovery_email('90000000-0000-4000-8000-000000000061','90000000-0000-4000-8000-000000000062',1)::text,true);
select public.console_finish_recovery_email((current_setting('test.recovery')::jsonb->>'request_id')::uuid,false);
reset role;
do $$begin
 assert (select count(*) from public.console_user_events where user_id='90000000-0000-4000-8000-000000000062')=1,'Failed transport not audited as sent';
end$$;
update public.console_users set active=false where user_id='90000000-0000-4000-8000-000000000062';
set local role service_role;
do $$begin
 begin perform public.console_request_recovery_email('90000000-0000-4000-8000-000000000061','90000000-0000-4000-8000-000000000062',1);raise exception 'Suspended recipient accepted';exception when invalid_parameter_value then null;end;
end$$;
reset role;
select 'Recovery email authorization, destination, cooldown, audit and suspended recipient checks passed; fixtures rolled back.' as result;
rollback;
