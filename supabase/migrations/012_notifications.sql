-- 012_notifications.sql
-- Projeto do DASHBOARD (ftmgmfdqdqxboiktxcoj). Rodar no SQL Editor desse projeto.
--
-- (1) ads_balance_history.alert_sent — rastreia se o snapshot efetivamente
-- disparou uma mensagem no Telegram (check-ads-balance passa a reenviar 1x/dia
-- quando o nível fica estagnado, em vez de silenciar pra sempre — ver P0 do
-- plano de notificações).
--
-- (2) health_alerts — dedup dos alertas de check-tracking-health (watchdog de
-- frescor de sync). RLS habilitado sem policies, mesmo padrão de
-- conversion_dispatches (007/008): só service_role acessa.

alter table public.ads_balance_history
  add column if not exists alert_sent boolean not null default false;

create table if not exists health_alerts (
  id         bigserial primary key,
  alert_key  text not null,
  sent_at    timestamptz not null default now()
);

create index if not exists idx_health_alerts_key_sent_at
  on health_alerts (alert_key, sent_at desc);

alter table public.health_alerts enable row level security;
