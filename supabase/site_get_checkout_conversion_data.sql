-- ============================================================================
-- get_checkout_conversion_data — RPC a ser criada no Supabase do SITE
-- (ozquoloetuzynnyzkado), não no Supabase do dashboard.
--
-- Usada pela Edge Function `dispatch-conversions` (repo do dashboard) pra
-- resolver email/telefone do comprador via auth.users — a REST API do
-- PostgREST não expõe o schema `auth` diretamente, então precisa de uma
-- function SECURITY DEFINER (mesmo padrão já usado em get_bureau_costs_daily).
--
-- P1.3: adicionadas as colunas de atribuição (gclid/gbraid/wbraid/fbclid/fbp/fbc/
-- ga_client_id/utm_*/landing_page/event_source_url/client_user_agent).
-- Pré-requisito: rodar antes site_add_attribution_columns.sql (mesmo projeto).
--
-- Deploy: rodar este SQL no SQL Editor do projeto ozquoloetuzynnyzkado (site).
-- Chamada via REST: POST {BUREAU_SUPABASE_URL}/rest/v1/rpc/get_checkout_conversion_data
--   headers: apikey/Authorization = BUREAU_SUPABASE_KEY (mesmo secret já usado
--   pelo sync-bureau-daily — é o MESMO projeto, reaproveita o secret existente).
-- ============================================================================

-- P1.3: o return type mudou (colunas de atribuição novas) — Postgres não
-- permite mudar retorno com create or replace, precisa dropar antes.
-- Seguro: só a Edge Function dispatch-conversions consome esta RPC; entre o
-- drop e o create (mesma execução) não há downtime perceptível.
drop function if exists public.get_checkout_conversion_data(uuid);

create or replace function public.get_checkout_conversion_data(p_checkout_id uuid)
returns table (
  checkout_id       uuid,
  order_nsu         character varying,
  placa             character varying,
  plano             character varying,
  valor             numeric,
  transaction_id    character varying,
  paid_at           timestamptz,
  user_id           uuid,
  email             character varying,
  phone             text,
  gclid             text,
  gbraid            text,
  wbraid            text,
  fbclid            text,
  fbp               text,
  fbc               text,
  ga_client_id      text,
  utm_source        text,
  utm_medium        text,
  utm_campaign      text,
  utm_term          text,
  utm_content       text,
  landing_page      text,
  event_source_url  text,
  client_user_agent text
)
language sql
security definer
set search_path = public, auth
as $$
  select
    c.id            as checkout_id,
    c.order_nsu,
    c.placa,
    c.plano,
    c.valor,
    c.transaction_id,
    c.paid_at,
    c.user_id,
    u.email,
    u.phone,
    c.gclid,
    c.gbraid,
    c.wbraid,
    c.fbclid,
    c.fbp,
    c.fbc,
    c.ga_client_id,
    c.utm_source,
    c.utm_medium,
    c.utm_campaign,
    c.utm_term,
    c.utm_content,
    c.landing_page,
    c.event_source_url,
    c.client_user_agent
  from public.checkouts c
  left join auth.users u on u.id = c.user_id
  where c.id = p_checkout_id;
$$;

-- Permite que a anon key (usada pelo secret BUREAU_SUPABASE_KEY) chame a function.
-- A função em si roda com privilégio do dono (security definer), então isso NÃO
-- expõe auth.users publicamente — só o retorno desta function específica,
-- filtrado por checkout_id.
grant execute on function public.get_checkout_conversion_data(uuid) to anon;
