-- 008_conversion_dispatches_public_view.sql
-- Projeto do DASHBOARD (ftmgmfdqdqxboiktxcoj). Rodar no SQL Editor desse projeto.
--
-- conversion_dispatches tem RLS habilitado sem policies (só service_role acessa)
-- porque guarda email/phone. O módulo Tracking Gateway do dashboard.html usa a
-- mesma anon key client-side que o resto do dashboard — não pode ler a tabela
-- direto. Esta view expõe só as colunas sem PII; dono da view é o role que a
-- criar (Supabase SQL Editor roda como postgres), que ignora RLS da tabela base,
-- então a view funciona mesmo com a tabela travada.

-- P3: expõe também atribuição (migration 010) + flags booleanas de email/phone
-- (nunca os hashes). create or replace só permite ADICIONAR colunas no fim —
-- a ordem abaixo preserva as existentes.

create or replace view conversion_dispatches_public as
select
  checkout_id,
  order_nsu,
  event_id,
  value,
  currency,
  ga4_status,
  ga4_error,
  ads_status,
  ads_error,
  meta_status,
  meta_error,
  attempts,
  created_at,
  updated_at,
  paid_at,
  gclid,
  gbraid,
  wbraid,
  fbp,
  fbc,
  ga_client_id,
  event_source_url,
  (email_sha256 is not null) as has_email,
  (phone_sha256 is not null) as has_phone
from conversion_dispatches;

grant select on conversion_dispatches_public to anon;
