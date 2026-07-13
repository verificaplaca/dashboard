-- 007_conversion_dispatches.sql
-- Tabela de controle de dispatch de conversões (GA4 MP / Google Ads Enhanced
-- Conversions / Meta CAPI) — projeto do DASHBOARD (ftmgmfdqdqxboiktxcoj).
--
-- Uma linha por checkout pago. Cada canal tem status próprio porque os 3
-- dispatches são independentes (um pode falhar sem bloquear os outros).
-- O job de retry (cron-job.org) varre linhas com algum canal != 'success'
-- e attempts < 5.

create table if not exists conversion_dispatches (
  checkout_id     uuid primary key,
  order_nsu       text not null,
  event_id        text not null,
  value           numeric not null,
  currency        text not null default 'BRL',
  email           text,
  phone           text,

  ga4_status      text not null default 'pending',   -- pending | success | failed | skipped
  ga4_error       text,
  ads_status      text not null default 'pending',
  ads_error       text,
  meta_status     text not null default 'pending',
  meta_error      text,

  attempts        int not null default 0,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create index if not exists idx_conversion_dispatches_pending
  on conversion_dispatches (attempts)
  where ga4_status <> 'success' or ads_status <> 'success' or meta_status <> 'success';
