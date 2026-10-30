begin;
do $$begin
 assert not exists(select 1 from public.monthly_targets where month in ('2098-10-01','2098-11-01')),'Fixture months must be unused';
end$$;
insert into auth.users(id,email,raw_user_meta_data) values
 ('90000000-0000-4000-8000-000000000091','targets-admin-test@example.invalid','{}'),
 ('90000000-0000-4000-8000-000000000092','targets-manager-test@example.invalid','{"role":"administrator"}'),
 ('90000000-0000-4000-8000-000000000093','targets-inactive-test@example.invalid','{}');
insert into public.console_users(user_id,display_name,email,role,active) values
 ('90000000-0000-4000-8000-000000000091','Targets admin','targets-admin-test@example.invalid','administrator',true),
 ('90000000-0000-4000-8000-000000000092','Targets manager','targets-manager-test@example.invalid','manager',true),
 ('90000000-0000-4000-8000-000000000093','Targets inactive','targets-inactive-test@example.invalid','administrator',false);
set local role authenticated;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000091","role":"authenticated"}',true);
insert into public.monthly_targets(month,revenue_target,tax_pct) values('2098-10-01',1000,8);
update public.monthly_targets set revenue_target=1200 where month='2098-10-01';
do $$begin assert (select revenue_target from public.monthly_targets where month='2098-10-01')=1200,'Admin edits targets';end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000092","role":"authenticated"}',true);
do $$declare n integer;begin
 assert (select revenue_target from public.monthly_targets where month='2098-10-01')=1200,'Manager can consult';
 begin insert into public.monthly_targets(month) values('2098-11-01');raise exception 'Manager inserted target';exception when insufficient_privilege then null;end;
 update public.monthly_targets set revenue_target=999,tax_pct=99 where month='2098-10-01';get diagnostics n=row_count;
 assert n=0,'Manager update blocked despite forged metadata';
 begin delete from public.monthly_targets where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Manager delete blocked';exception when insufficient_privilege then null;end;
end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000093","role":"authenticated"}',true);
do $$declare n integer;begin
 begin insert into public.monthly_targets(month) values('2098-11-01');raise exception 'Inactive admin inserted';exception when insufficient_privilege then null;end;
 update public.monthly_targets set revenue_target=999 where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Inactive admin cannot edit';
end$$;
reset role;
set local role anon;
do $$begin
 assert (select revenue_target from public.monthly_targets where month='2098-10-01')=1200,'Public read preserved';
 begin insert into public.monthly_targets(month) values('2098-11-01');raise exception 'Anonymous insert';exception when insufficient_privilege then null;end;
end$$;
reset role;
select 'Targets: admin write; manager read-only; forged metadata, inactive admin and anon mutations denied. Fixtures rolled back.' as result;
rollback;
