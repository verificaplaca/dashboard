begin;
select set_config('test.admin',(select user_id::text from public.console_users where active and role='administrator' limit 1),true);
insert into auth.users(id,email,raw_user_meta_data) values
 ('90000000-0000-4000-8000-000000000081','manager-users-test@example.invalid','{"role":"administrator"}'),
 ('90000000-0000-4000-8000-000000000082','admin-users-test@example.invalid','{}'),
 ('90000000-0000-4000-8000-000000000083','outsider-users-test@example.invalid','{"role":"administrator"}');
set local role anon;
do $$begin
 begin perform * from public.console_users;raise exception 'Anon read allowed';exception when insufficient_privilege then null;end;
 begin perform public.console_update_user(null,'administrator',true,1);raise exception 'Anon RPC allowed';exception when insufficient_privilege then null;end;
end$$;
reset role;
set local role service_role;
select public.console_provision_user(current_setting('test.admin')::uuid,'90000000-0000-4000-8000-000000000081','Test manager','manager');
do $$begin
 begin perform public.console_provision_user('90000000-0000-4000-8000-000000000081','90000000-0000-4000-8000-000000000082','Forbidden','administrator');raise exception 'Manager provisioned';exception when insufficient_privilege then null;end;
end$$;
reset role;
set local role authenticated;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000081","role":"authenticated"}',true);
do $$begin
 assert console_private.console_role()='manager','Editable metadata cannot elevate';
 assert (select count(*) from public.console_users)=1,'Manager only sees own account';
 assert (select count(*) from public.console_user_events)=0,'Manager cannot read audit';
 assert public.console_finance_access(),'Manager has financial access';
 begin update public.console_users set role='administrator';raise exception 'Direct escalation';exception when insufficient_privilege then null;end;
 begin perform public.console_update_user(auth.uid(),'administrator',true,1);raise exception 'Manager escalation';exception when insufficient_privilege then null;end;
 begin perform public.console_provision_user(auth.uid(),'90000000-0000-4000-8000-000000000082','Forged','administrator');raise exception 'Public provision';exception when insufficient_privilege then null;end;
end$$;
select set_config('request.jwt.claims',jsonb_build_object('sub',current_setting('test.admin'),'role','authenticated')::text,true);
do $$declare r public.console_users;begin
 begin perform public.console_update_user(auth.uid(),'manager',true,1);raise exception 'Last admin demoted';exception when invalid_parameter_value then null;end;
 begin perform public.console_update_user(auth.uid(),'administrator',false,1);raise exception 'Last admin suspended';exception when invalid_parameter_value then null;end;
 r:=public.console_update_user('90000000-0000-4000-8000-000000000081','manager',false,1);
 assert r.version=2 and not r.active,'Versioned suspension';
 assert (select count(*) from public.console_user_events where user_id=r.user_id)=2,'Creation and suspension audited';
 begin perform public.console_update_user(r.user_id,'administrator',true,1);raise exception 'Stale update';exception when serialization_failure then null;end;
end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000081","role":"authenticated"}',true);
do $$begin
 assert console_private.console_role() is null,'Suspended access';
 assert console_private.task_role() is null,'Task suspension';
 assert not public.console_finance_access(),'Financial suspension';
 begin perform public.console_save_finance_month('2099-12-01','{"fees":0,"people":0,"systems":0,"other":0,"financial":0}',0);raise exception 'Suspended finance write';exception when insufficient_privilege then null;end;
 begin insert into public.monthly_targets(month) values('2099-12-01');raise exception 'Suspended target write';exception when insufficient_privilege then null;end;
 begin perform public.console_create_task('{"title":"Forbidden task"}',auth.uid());raise exception 'Suspended task write';exception when insufficient_privilege then null;end;
end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000083","role":"authenticated"}',true);
do $$begin
 assert console_private.console_role() is null,'Auth account alone insufficient';
 begin insert into public.monthly_targets(month) values('2099-12-01');raise exception 'Outsider target write';exception when insufficient_privilege then null;end;
end$$;
select set_config('request.jwt.claims',jsonb_build_object('sub',current_setting('test.admin'),'role','authenticated')::text,true);
select public.console_update_user('90000000-0000-4000-8000-000000000081','administrator',true,2);
-- A second active administrator allows deliberate demotion, preserving access.
select public.console_update_user(current_setting('test.admin')::uuid,'manager',true,1);
reset role;
rollback;
