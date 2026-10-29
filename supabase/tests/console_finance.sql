begin;
do $$begin
 perform set_config('test.finance_manager',(select user_id::text from public.console_task_members where active and role='manager' limit 1),true);
 assert current_setting('test.finance_manager',true) is not null,'Existing manager required';
 assert not exists(select 1 from public.console_finance_months where month='2099-12-01'),'Fixture month must be unused';
end$$;
set local role anon;
do $$begin
 begin perform * from public.console_finance_months;raise exception 'Anonymous read allowed';exception when insufficient_privilege then null;end;
 begin perform public.console_finance_access();raise exception 'Anonymous RPC allowed';exception when insufficient_privilege then null;end;
end$$;
reset role;
set local role authenticated;
do $$declare r public.console_finance_months;begin
 perform set_config('request.jwt.claims',jsonb_build_object('sub',current_setting('test.finance_manager'),'role','authenticated')::text,true);
 assert public.console_finance_access(),'Manager access';
 r:=public.console_save_finance_month('2099-12-01','{"fees":0,"people":100,"systems":null,"other":20,"financial":0}',0);
 assert r.fees=0 and r.systems is null and r.version=1,'Zero and missing differ';
 assert r.updated_by=auth.uid(),'Server identity';
 assert (select count(*) from public.console_finance_history where month='2099-12-01')=1,'Audit initial';
 r:=public.console_save_finance_month('2099-12-01','{"fees":10.25,"people":100,"systems":50,"other":20,"financial":0}',1);
 assert r.version=2 and r.fees=10.25,'Versioned update';
 assert (select count(*) from public.console_finance_history where month='2099-12-01')=2,'Audit update';
 begin perform public.console_save_finance_month('2099-12-01','{"fees":0,"people":0,"systems":0,"other":0,"financial":0}',1);raise exception 'Stale update';exception when serialization_failure then null;end;
 begin perform public.console_save_finance_month('2099-12-01','{"fees":0,"people":0,"systems":0,"other":0,"financial":0}',0);raise exception 'Duplicate create';exception when serialization_failure then null;end;
 begin perform public.console_save_finance_month('2099-12-02','{"fees":0,"people":0,"systems":0,"other":0,"financial":0}',0);raise exception 'Wrong month';exception when invalid_parameter_value then null;end;
 begin perform public.console_save_finance_month('2099-12-01','{"fees":-1,"people":0,"systems":0,"other":0,"financial":0}',2);raise exception 'Negative fee';exception when invalid_parameter_value then null;end;
 begin perform public.console_save_finance_month('2099-12-01','{"fees":"NaN","people":0,"systems":0,"other":0,"financial":0}',2);raise exception 'String amount';exception when invalid_parameter_value then null;end;
 begin perform public.console_save_finance_month('2099-12-01','{"fees":1.001,"people":0,"systems":0,"other":0,"financial":0}',2);raise exception 'Fraction cents';exception when invalid_parameter_value then null;end;
 begin perform public.console_save_finance_month('2099-12-01','{"fees":0,"people":0,"systems":0,"other":0,"financial":0,"updated_by":"forged"}',2);raise exception 'Unknown field';exception when invalid_parameter_value then null;end;
 begin update public.console_finance_months set fees=999 where month='2099-12-01';raise exception 'Direct write';exception when insufficient_privilege then null;end;
 begin delete from public.console_finance_history where month='2099-12-01';raise exception 'Audit deletion';exception when insufficient_privilege then null;end;
 perform set_config('request.jwt.claims','{"sub":"90000000-0000-4000-8000-000000000099","role":"authenticated"}',true);
 assert not public.console_finance_access(),'Outsider blocked';
 assert (select count(*) from public.console_finance_months)=0,'Outsider RLS';
 assert (select count(*) from public.console_finance_history)=0,'Outsider audit RLS';
 begin perform public.console_save_finance_month('2099-12-01','{"fees":0,"people":0,"systems":0,"other":0,"financial":0}',2);raise exception 'Outsider save';exception when insufficient_privilege then null;end;
end$$;
reset role;
update public.console_task_members set active=false where user_id=current_setting('test.finance_manager')::uuid;
set local role authenticated;
do $$begin
 perform set_config('request.jwt.claims',jsonb_build_object('sub',current_setting('test.finance_manager'),'role','authenticated')::text,true);
 assert not public.console_finance_access(),'Revocation blocks access';
 assert (select count(*) from public.console_finance_months)=0,'Revocation RLS';
 begin perform public.console_save_finance_month('2099-12-01','{"fees":0,"people":0,"systems":0,"other":0,"financial":0}',2);raise exception 'Revoked save';exception when insufficient_privilege then null;end;
end$$;
rollback;
