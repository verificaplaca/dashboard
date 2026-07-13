-- ============================================================================
-- get_checkout_conversion_data — RPC a ser criada no Supabase do SITE
-- (ozquoloetuzynnyzkado), não no Supabase do dashboard.
--
-- Usada pela Edge Function `dispatch-conversions` (repo do dashboard) pra
-- resolver email/telefone do comprador via auth.users — a REST API do
-- PostgREST não expõe o schema `auth` diretamente, então precisa de uma
-- function SECURITY DEFINER (mesmo padrão já usado em get_bureau_costs_daily).
--
-- Deploy: rodar este SQL no SQL Editor do projeto ozquoloetuzynnyzkado (site).
-- Chamada via REST: POST {BUREAU_SUPABASE_URL}/rest/v1/rpc/get_checkout_conversion_data
--   headers: apikey/Authorization = BUREAU_SUPABASE_KEY (mesmo secret já usado
--   pelo sync-bureau-daily — é o MESMO projeto, reaproveita o secret existente).
-- ============================================================================

create or replace function public.get_checkout_conversion_data(p_checkout_id uuid)
returns table (
  checkout_id     uuid,
  order_nsu       character varying,
  placa           character varying,
  plano           character varying,
  valor           numeric,
  transaction_id  character varying,
  paid_at         timestamptz,
  user_id         uuid,
  email           character varying,
  phone           text
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
    u.phone
  from public.checkouts c
  left join auth.users u on u.id = c.user_id
  where c.id = p_checkout_id;
$$;

-- Permite que a anon key (usada pelo secret BUREAU_SUPABASE_KEY) chame a function.
-- A função em si roda com privilégio do dono (security definer), então isso NÃO
-- expõe auth.users publicamente — só o retorno desta function específica,
-- filtrado por checkout_id.
grant execute on function public.get_checkout_conversion_data(uuid) to anon;
