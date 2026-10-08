-- Transactional integration test. Never commits users, tasks or events.
begin;
select set_config('test.task_manager',(select user_id::text from public.console_task_members where role='manager' and active limit 1),true);
insert into auth.users(id,aud,role,email,created_at,updated_at) values
 ('10000000-0000-4000-8000-000000000001','authenticated','authenticated','task-member@example.invalid',now(),now()),
 ('10000000-0000-4000-8000-000000000002','authenticated','authenticated','task-outsider@example.invalid',now(),now()),
 ('10000000-0000-4000-8000-000000000003','authenticated','authenticated','task-revoked@example.invalid',now(),now());
insert into public.console_task_members(user_id,display_name,role,active) values
 ('10000000-0000-4000-8000-000000000001','Test member','member',true),
 ('10000000-0000-4000-8000-000000000002','Test outsider','member',true),
 ('10000000-0000-4000-8000-000000000003','Test revoked','member',false);

set local role authenticated;
select set_config('request.jwt.claims',jsonb_build_object('sub',current_setting('test.task_manager'),'role','authenticated')::text,true);
do $$ declare t public.console_tasks; begin
 t:=public.console_create_task(jsonb_build_object('id','20000000-0000-4000-8000-000000000001','title','Transactional task test','assignee_id','10000000-0000-4000-8000-000000000001','due_date',((now() at time zone 'America/Sao_Paulo')::date-1)::text),auth.uid());
 assert t.version=1 and t.created_by::text=current_setting('test.task_manager') and t.completed_at is null, 'create identity/version';
 assert (select count(*) from public.console_task_events where task_id=t.id)=1,'atomic creation audit';
 t:=public.console_create_task(jsonb_build_object('id',t.id,'title','Transactional task test','assignee_id','10000000-0000-4000-8000-000000000001','due_date',((now() at time zone 'America/Sao_Paulo')::date-1)::text),auth.uid());
 assert (select count(*) from public.console_task_events where task_id=t.id)=1,'idempotent creation';
 begin perform public.console_create_task('{"title":"Changed identity"}','10000000-0000-4000-8000-000000000002');raise exception 'identity substitution accepted';exception when insufficient_privilege then null;end;
 t:=public.console_update_task(t.id,1,'{"priority":"high","description":"Criteria","status":"blocked"}',auth.uid());
 assert t.version=2 and t.priority='high' and t.status='blocked','manager update';
 begin perform public.console_update_task(t.id,1,'{"title":"Stale write"}',auth.uid());raise exception 'stale write accepted';exception when serialization_failure then null;end;
 begin perform public.console_update_task(t.id,null,'{"status":"done"}',auth.uid());raise exception 'null version accepted';exception when serialization_failure then null;end;
 begin perform public.console_create_task('{"title":"Spoof creator","created_by":"10000000-0000-4000-8000-000000000002"}',auth.uid());raise exception 'creator spoof accepted';exception when invalid_parameter_value then null;end;
end $$;

select set_config('request.jwt.claims','{"sub":"10000000-0000-4000-8000-000000000001","role":"authenticated"}',true);
do $$ declare t public.console_tasks; begin
 assert (select count(*) from public.console_tasks where id='20000000-0000-4000-8000-000000000001')=1,'assignee reads task';
 assert (public.console_task_summary()->>'overdue')::int >= 1,'overdue summary';
 begin perform public.console_update_task('20000000-0000-4000-8000-000000000001',2,'{"priority":"low"}',auth.uid());raise exception 'assignee edited scope';exception when insufficient_privilege then null;end;
 begin update public.console_tasks set title='Bypassed RPC' where id='20000000-0000-4000-8000-000000000001';raise exception 'direct write accepted';exception when insufficient_privilege then null;end;
 begin update public.console_task_members set role='manager' where user_id=auth.uid();raise exception 'self promotion accepted';exception when insufficient_privilege then null;end;
 t:=public.console_update_task('20000000-0000-4000-8000-000000000001',2,'{"status":"doing"}',auth.uid());assert t.version=3,'assignee starts';
 t:=public.console_update_task(t.id,3,'{"status":"done"}',auth.uid());assert t.completed_at is not null and t.version=4,'completion';
 t:=public.console_update_task(t.id,4,'{"status":"todo"}',auth.uid());assert t.completed_at is null and t.version=5,'reopening';
 perform public.console_add_task_comment(t.id,'  Comment test  ',auth.uid());
 assert (select count(*) from public.console_task_events where task_id=t.id)=6,'audit/comment sequence';
 assert exists(select 1 from public.console_task_events where task_id=t.id and body='Comment test' and actor_id=auth.uid() and task_version=5),'comment identity/trim/version';
 begin perform public.console_add_task_comment(t.id,'  ',auth.uid());raise exception 'empty comment accepted';exception when invalid_parameter_value then null;end;
 t:=public.console_create_task('{"id":"20000000-0000-4000-8000-000000000002","title":"Own task","assignee_id":"10000000-0000-4000-8000-000000000001"}',auth.uid());
 t:=public.console_update_task(t.id,1,'{"title":"Own edited task","priority":"low"}',auth.uid());assert t.title='Own edited task','creator edit';
 begin perform public.console_update_task(t.id,2,'{"created_by":"10000000-0000-4000-8000-000000000002"}',auth.uid());raise exception 'immutable owner accepted';exception when invalid_parameter_value then null;end;
 begin perform public.console_create_task('{"title":"No assignee"}',auth.uid());raise exception 'unassigned member task accepted';exception when invalid_parameter_value then null;end;
 begin perform public.console_update_task(t.id,2,'{"assignee_id":"10000000-0000-4000-8000-000000000003"}',auth.uid());raise exception 'inactive assignee accepted';exception when invalid_parameter_value then null;end;
end $$;

select set_config('request.jwt.claims','{"sub":"10000000-0000-4000-8000-000000000002","role":"authenticated"}',true);
do $$ begin
 assert (select count(*) from public.console_tasks where id in ('20000000-0000-4000-8000-000000000001','20000000-0000-4000-8000-000000000002'))=0,'outsider RLS';
 assert (select count(*) from public.console_task_events where task_id='20000000-0000-4000-8000-000000000001')=0,'outsider history RLS';
 begin perform public.console_update_task('20000000-0000-4000-8000-000000000001',5,'{"status":"done"}',auth.uid());raise exception 'outsider write accepted';exception when insufficient_privilege then null;end;
 begin perform public.console_add_task_comment('20000000-0000-4000-8000-000000000001','Intrusion',auth.uid());raise exception 'outsider comment accepted';exception when insufficient_privilege then null;end;
end $$;
select set_config('request.jwt.claims','{"sub":"10000000-0000-4000-8000-000000000003","role":"authenticated"}',true);
do $$ begin
 assert (select count(*) from public.console_task_members)=0,'revoked directory RLS';
 assert (select count(*) from public.console_tasks)=0,'revoked task RLS';
 begin perform public.console_create_task('{"title":"Revoked user task"}',auth.uid());raise exception 'revoked member created';exception when insufficient_privilege then null;end;
end $$;

select set_config('request.jwt.claims',jsonb_build_object('sub',current_setting('test.task_manager'),'role','authenticated')::text,true);
do $$ declare t public.console_tasks; begin
 t:=public.console_update_task('20000000-0000-4000-8000-000000000001',5,'{"archived":true}',auth.uid());assert t.archived_at is not null and t.version=6,'archive';
 begin perform public.console_update_task(t.id,6,'{"status":"done"}',auth.uid());raise exception 'archived task changed';exception when invalid_parameter_value then null;end;
 begin perform public.console_add_task_comment(t.id,'Archived comment',auth.uid());raise exception 'archived comment accepted';exception when invalid_parameter_value then null;end;
 t:=public.console_update_task(t.id,6,'{"archived":false}',auth.uid());assert t.archived_at is null and t.version=7,'restore';
 assert (select count(*) from public.console_task_events where task_id=t.id)=8,'archive/restore audit';
 begin delete from public.console_tasks where id=t.id;raise exception 'hard delete accepted';exception when insufficient_privilege then null;end;
end $$;
reset role;
set local role anon;
select set_config('request.jwt.claims','{"role":"anon"}',true);
do $$ begin
 begin perform 1 from public.console_tasks;raise exception 'anonymous read accepted';exception when insufficient_privilege then null;end;
 begin perform public.console_task_summary();raise exception 'anonymous summary accepted';exception when insufficient_privilege then null;end;
 begin perform public.console_create_task('{"title":"Anon task"}',auth.uid());raise exception 'anonymous RPC accepted';exception when insufficient_privilege then null;end;
end $$;
reset role;
select 'Tasks: identity, RLS, scoped permissions, immutable fields, audit, comments, deadlines, completion, reopening, archive/restore, stale writes and anon denial passed; ALL fixtures rolled back.' as result;
rollback;
