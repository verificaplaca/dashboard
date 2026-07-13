-- ============================================================================
-- P1.2 — Colunas de atribuição em public.checkouts
-- Rodar no SQL Editor do projeto Supabase do SITE (ozquoloetuzynnyzkado),
-- NÃO no projeto do dashboard (ftmgmfdqdqxboiktxcoj).
--
-- Pré-requisito para o P1.1 (patch do site) e para o P1.3 (RPC atualizada)
-- funcionarem — as colunas precisam existir antes do insert/select tentar
-- gravar/ler nelas.
-- ============================================================================

alter table public.checkouts
  add column if not exists gclid text,
  add column if not exists gbraid text,
  add column if not exists wbraid text,
  add column if not exists fbclid text,
  add column if not exists fbp text,
  add column if not exists fbc text,
  add column if not exists ga_client_id text,
  add column if not exists utm_source text,
  add column if not exists utm_medium text,
  add column if not exists utm_campaign text,
  add column if not exists utm_term text,
  add column if not exists utm_content text,
  add column if not exists landing_page text,
  add column if not exists event_source_url text,
  add column if not exists client_user_agent text;
