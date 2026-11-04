begin;
do $$begin
 assert not exists(select 1 from public.monthly_targets where month in ('2098-10-01','2098-11-01','2098-12-01')),'Fixture months must be unused';
end$$;
insert into auth.users(id,email,raw_user_meta_data) values
 ('90000000-0000-4000-8000-000000000091','targets-admin-test@example.invalid','{}'),
 ('90000000-0000-4000-8000-000000000092','targets-manager-test@example.invalid','{"role":"administrator"}'),
 ('90000000-0000-4000-8000-000000000093','targets-inactive-test@example.invalid','{}'),
 ('90000000-0000-4000-8000-000000000094','targets-inactive-manager@example.invalid','{}'),
 ('90000000-0000-4000-8000-000000000095','targets-unlisted@example.invalid','{"role":"administrator"}');
insert into public.console_users(user_id,display_name,email,role,active) values
 ('90000000-0000-4000-8000-000000000091','Targets admin','targets-admin-test@example.invalid','administrator',true),
 ('90000000-0000-4000-8000-000000000092','Targets manager','targets-manager-test@example.invalid','manager',true),
 ('90000000-0000-4000-8000-000000000093','Targets inactive','targets-inactive-test@example.invalid','administrator',false),
 ('90000000-0000-4000-8000-000000000094','Inactive manager','targets-inactive-manager@example.invalid','manager',false);
set local role authenticated;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000091","role":"authenticated"}',true);
insert into public.monthly_targets(month,revenue_target,tax_pct) values('2098-10-01',1000,8);
update public.monthly_targets set revenue_target=1200 where month='2098-10-01';
do $$begin assert (select revenue_target from public.monthly_targets where month='2098-10-01')=1200,'Admin edits targets';end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000092","role":"authenticated"}',true);
do $$declare n integer;begin
 assert (select revenue_target from public.monthly_targets where month='2098-10-01')=1200,'Manager can consult';
 insert into public.monthly_targets(month,revenue_target) values('2098-11-01',2000);
 assert (select tax_pct from public.monthly_targets where month='2098-11-01') is null,'New month inherits tax';
 update public.monthly_targets set revenue_target=1500,target_cpa=10,monthly_budget=5000 where month='2098-10-01';
 get diagnostics n=row_count;assert n=1,'Manager updates targets';
 assert (select tax_pct from public.monthly_targets where month='2098-10-01')=8,'Manager preserves tax';
 begin update public.monthly_targets set tax_pct=99 where month='2098-10-01';raise exception 'Manager changed tax';exception when insufficient_privilege then null;end;
 begin insert into public.monthly_targets(month,tax_pct) values('2098-12-01',99);raise exception 'Manager inserted tax';exception when insufficient_privilege then null;end;
 begin update public.monthly_targets set month='2098-12-01' where month='2098-10-01';raise exception 'Manager moved taxed month';exception when insufficient_privilege then null;end;
 begin delete from public.monthly_targets where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Manager delete blocked';exception when insufficient_privilege then null;end;
end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000093","role":"authenticated"}',true);
do $$declare n integer;begin
 begin insert into public.monthly_targets(month) values('2098-11-01');raise exception 'Inactive admin inserted';exception when insufficient_privilege then null;end;
 update public.monthly_targets set revenue_target=999 where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Inactive admin cannot edit';
end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000094","role":"authenticated"}',true);
do $$declare n integer;begin
 update public.monthly_targets set revenue_target=999 where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Inactive manager cannot edit';
 begin insert into public.monthly_targets(month) values('2098-12-01');raise exception 'Inactive manager inserted';exception when insufficient_privilege then null;end;
end$$;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000095","role":"authenticated"}',true);
do $$declare n integer;begin
 update public.monthly_targets set revenue_target=999 where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Forged metadata without membership cannot edit';
end$$;
reset role;
set local role anon;
do $$declare n integer;begin
 assert (select revenue_target from public.monthly_targets where month='2098-10-01')=1500,'Public read preserved';
 begin insert into public.monthly_targets(month) values('2098-11-01');raise exception 'Anonymous insert';exception when insufficient_privilege then null;end;
 update public.monthly_targets set revenue_target=999 where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Anonymous cannot update';
 delete from public.monthly_targets where month='2098-10-01';get diagnostics n=row_count;assert n=0,'Anonymous cannot delete';
end$$;
reset role;
set local role authenticated;
select set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000091","role":"authenticated"}',true);
update public.monthly_targets set tax_pct=9 where month='2098-10-01';
do $$declare n integer;begin
 assert (select tax_pct from public.monthly_targets where month='2098-10-01')=9,'Admin retains tax editing';
end$$;
reset role;
select 'Targets: active manager edits goals; tax remains administrator-only; inactive, unlisted and anonymous users denied. Fixtures rolled back.' as result;
rollback;
