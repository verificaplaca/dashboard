-- Cache closed São Paulo days only. Each live tail starts at the last
-- cached date, so a delayed refresh cannot lose a day at midnight.
begin;
create extension if not exists pg_cron;
create schema if not exists vp_analytics_cache;
revoke all on schema vp_analytics_cache from public, anon, authenticated;

create materialized view vp_analytics_cache.revenue_daily as
  select * from public.revenue_daily
  where date < (now() at time zone 'America/Sao_Paulo')::date;
create unique index on vp_analytics_cache.revenue_daily (date);

create view public.console_revenue_daily as
  select * from vp_analytics_cache.revenue_daily
  union all
  select * from public.revenue_daily
  where date > coalesce((select max(date) from vp_analytics_cache.revenue_daily), '-infinity'::date);
revoke all on public.console_revenue_daily from public, anon, authenticated;
grant select on public.console_revenue_daily to anon, authenticated;

create materialized view vp_analytics_cache.upsell_daily as
  select * from public.upsell_daily
  where date < (now() at time zone 'America/Sao_Paulo')::date;
create unique index on vp_analytics_cache.upsell_daily (date);

create view public.console_upsell_daily as
  select * from vp_analytics_cache.upsell_daily
  union all
  select * from public.upsell_daily
  where date > coalesce((select max(date) from vp_analytics_cache.upsell_daily), '-infinity'::date);
revoke all on public.console_upsell_daily from public, anon, authenticated;
grant select on public.console_upsell_daily to anon, authenticated;

create materialized view vp_analytics_cache.upsell_by_type as
  select * from public.upsell_by_type
  where date < (now() at time zone 'America/Sao_Paulo')::date;
create unique index on vp_analytics_cache.upsell_by_type (date, addon_key);

create view public.console_upsell_by_type as
  select * from vp_analytics_cache.upsell_by_type
  union all
  select * from public.upsell_by_type
  where date > coalesce((select max(date) from vp_analytics_cache.upsell_by_type), '-infinity'::date);
revoke all on public.console_upsell_by_type from public, anon, authenticated;
grant select on public.console_upsell_by_type to anon, authenticated;

select cron.schedule('vp-refresh-dashboard-history', '7,22,37,52 * * * *', $$
  refresh materialized view concurrently vp_analytics_cache.revenue_daily;
  refresh materialized view concurrently vp_analytics_cache.upsell_daily;
  refresh materialized view concurrently vp_analytics_cache.upsell_by_type;
$$);
notify pgrst, 'reload schema';
commit;
