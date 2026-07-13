-- ============================================================================
-- list_paid_checkout_ids — RPC a ser criada no Supabase do SITE
-- (ozquoloetuzynnyzkado), não no Supabase do dashboard.
--
-- Usada pela Edge Function `reconcile-conversion-dispatches` (repo do
-- dashboard) pra listar os checkouts pagos nas últimas N horas e comparar
-- com o que já está em conversion_dispatches — cobre falha/timeout do
-- Database Webhook (P1.5).
--
-- Deploy: rodar este SQL no SQL Editor do projeto ozquoloetuzynnyzkado (site).
-- Chamada via REST: POST {BUREAU_SUPABASE_URL}/rest/v1/rpc/list_paid_checkout_ids
--   headers: apikey/Authorization = BUREAU_SUPABASE_KEY (mesmo secret já usado
--   por get_checkout_conversion_data — é o MESMO projeto).
-- ============================================================================

create or replace function public.list_paid_checkout_ids(p_since timestamptz)
returns table (checkout_id uuid)
language sql security definer set search_path = public
as $$ select id from public.checkouts where paid_at is not null and paid_at >= p_since; $$;

grant execute on function public.list_paid_checkout_ids(timestamptz) to anon;
