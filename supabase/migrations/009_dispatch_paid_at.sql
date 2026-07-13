-- 009 — adiciona paid_at em conversion_dispatches (P0-b do plano-tracking-profissional.md)
-- Rodar no SQL Editor do projeto do DASHBOARD (ftmgmfdqdqxboiktxcoj).
-- Usado pelos dispatchers para enviar timestamp real da conversão
-- (GA4 MP timestamp_micros / Meta CAPI event_time) em vez da hora do dispatch/retry.

alter table public.conversion_dispatches add column if not exists paid_at timestamptz;
