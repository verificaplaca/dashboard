begin;
do $$begin
 assert not exists(select 1 from auth.users where id in ('90000000-0000-4000-8000-000000000071','90000000-0000-4000-8000-000000000072')),'Fixture IDs must be unused';
end$$;
insert into auth.users(id,email,raw_user_meta_data,raw_app_meta_data,encrypted_password) values
 ('90000000-0000-4000-8000-000000000071','details-admin@example.invalid','{}','{}','fixture-old'),
 ('90000000-0000-4000-8000-000000000072','details-manager@example.invalid','{"role":"administrator"}','{}','fixture-old');
insert into public.console_users(user_id,display_name,email,role) values
 ('90000000-0000-4000-8000-000000000071','Fixture admin','details-admin@example.invalid','administrator'),
 ('90000000-0000-4000-8000-000000000072','Fixture manager','details-manager@example.invalid','manager');
set local role authenticated;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000072","role":"authenticated"}',true);
do $$begin
 begin perform public.console_prepare_user_change(auth.uid(),auth.uid(),1,'reset_password');raise exception 'Authenticated preparation allowed';exception when insufficient_privilege then null;end;
 begin perform * from console_private.user_changes;raise exception 'Read operation intents allowed';exception when insufficient_privilege then null;end;
 begin perform public.console_cancel_user_change(gen_random_uuid());raise exception 'Public cancellation allowed';exception when insufficient_privilege then null;end;
end$$;
reset role;
set local role service_role;
do $$begin
 begin perform public.console_prepare_user_change('90000000-0000-4000-8000-000000000072','90000000-0000-4000-8000-000000000072',1,'reset_password');raise exception 'Manager prepared reset';exception when insufficient_privilege then null;end;
 begin perform public.console_prepare_user_change('90000000-0000-4000-8000-000000000071','90000000-0000-4000-8000-000000000072',0,'reset_password');raise exception 'Stale preparation allowed';exception when serialization_failure then null;end;
end$$;
select set_config('test.details_token',public.console_prepare_user_change('90000000-0000-4000-8000-000000000071','90000000-0000-4000-8000-000000000072',1,'update','Renamed manager','renamed-details@example.invalid','manager',true)::text,true);
reset role;
-- Emulate Auth admin transaction: email first, app metadata then user metadata.
update auth.users set email='renamed-details@example.invalid' where id='90000000-0000-4000-8000-000000000072';
update auth.users set raw_app_meta_data=coalesce(raw_app_meta_data,'{}')||jsonb_build_object('console_operation',current_setting('test.details_token')) where id='90000000-0000-4000-8000-000000000072';
update auth.users set raw_user_meta_data=raw_user_meta_data||'{"full_name":"Renamed manager"}'::jsonb where id='90000000-0000-4000-8000-000000000072';
do $$declare m public.console_users;begin
 select * into m from public.console_users where user_id='90000000-0000-4000-8000-000000000072';
 assert m.display_name='Renamed manager' and m.email='renamed-details@example.invalid' and m.version=2;
 assert (select display_name='Renamed manager' from public.console_task_members where user_id=m.user_id),'Task name synchronized';
 assert (select not(raw_app_meta_data?'console_operation') from auth.users where id=m.user_id),'Marker stripped';
 assert not exists(select 1 from console_private.user_changes where token::text=current_setting('test.details_token')),'Intent consumed';
 assert exists(select 1 from public.console_user_events where user_id=m.user_id and after_value->>'action'='update'),'Profile audited';
 begin update auth.users set raw_app_meta_data=jsonb_build_object('console_operation',current_setting('test.details_token')) where id=m.user_id;raise exception 'Marker replay allowed';exception when insufficient_privilege then null;end;
end$$;
set local role service_role;
select set_config('test.details_token',public.console_prepare_user_change('90000000-0000-4000-8000-000000000071','90000000-0000-4000-8000-000000000072',2,'reset_password')::text,true);
reset role;
-- One transaction with a forced failure must also roll back the Auth password.
do $$begin
 begin
  update auth.users set encrypted_password='fixture-replacement' where id='90000000-0000-4000-8000-000000000072';
  update public.console_users set version=version+1 where user_id='90000000-0000-4000-8000-000000000072';
  update auth.users set raw_app_meta_data=jsonb_build_object('console_operation',current_setting('test.details_token')) where id='90000000-0000-4000-8000-000000000072';
  raise exception 'Stale update accepted';
 exception when serialization_failure then null;end;
 assert (select encrypted_password='fixture-old' from auth.users where id='90000000-0000-4000-8000-000000000072'),'Auth password rolled back';
 assert (select version=2 from public.console_users where user_id='90000000-0000-4000-8000-000000000072');
end$$;
-- Revocation between preparation and Auth must fail atomically, too.
update public.console_users set active=false where user_id='90000000-0000-4000-8000-000000000071';
do $$begin
 begin update auth.users set encrypted_password='fixture-replacement',raw_app_meta_data=jsonb_build_object('console_operation',current_setting('test.details_token')) where id='90000000-0000-4000-8000-000000000072';raise exception 'Revoked administrator reset';exception when insufficient_privilege then null;end;
 assert (select encrypted_password='fixture-old' from auth.users where id='90000000-0000-4000-8000-000000000072');
end$$;
update public.console_users set active=true where user_id='90000000-0000-4000-8000-000000000071';
update auth.users set encrypted_password='fixture-replacement',raw_app_meta_data=jsonb_build_object('console_operation',current_setting('test.details_token')) where id='90000000-0000-4000-8000-000000000072';
do $$begin
 assert (select version=3 and email='renamed-details@example.invalid' from public.console_users where user_id='90000000-0000-4000-8000-000000000072');
 assert exists(select 1 from public.console_user_events where user_id='90000000-0000-4000-8000-000000000072' and after_value->>'action'='reset_password');
 assert not exists(select 1 from public.console_user_events where user_id='90000000-0000-4000-8000-000000000072' and (after_value::text like '%fixture-replacement%' or after_value::text like '%console_operation%')),'Audit excludes credentials';
end$$;
-- Invalid/expired intents cannot change Auth.
set local role service_role;
select set_config('test.details_token',public.console_prepare_user_change('90000000-0000-4000-8000-000000000071','90000000-0000-4000-8000-000000000072',3,'reset_password')::text,true);
reset role;
update console_private.user_changes set expires_at=clock_timestamp()-interval '1 second' where token::text=current_setting('test.details_token');
do $$begin
 begin update auth.users set encrypted_password='fixture-expired',raw_app_meta_data=jsonb_build_object('console_operation',current_setting('test.details_token')) where id='90000000-0000-4000-8000-000000000072';raise exception 'Expired reset accepted';exception when insufficient_privilege then null;end;
 assert (select encrypted_password='fixture-replacement' from auth.users where id='90000000-0000-4000-8000-000000000072');
end$$;
select 'User changes: admin-only preparation, Auth/profile/audit consistency, stale/revoked/expired rollback, marker replay denied. Fixtures rolled back.' as result;
rollback;
