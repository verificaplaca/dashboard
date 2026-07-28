-- 015_phone_hash_split.sql — separa o hash de telefone por formato de canal
-- Projeto do DASHBOARD (ftmgmfdqdqxboiktxcoj). Rodar no SQL Editor desse projeto.
--
-- PROBLEMA
-- A coluna `phone_sha256` (migration 011) guardava SEMPRE o hash do telefone em
-- formato de dígitos ('5511999998888'), que é o que o Meta CAPI exige. Mas o
-- Google Ads exige E.164 ('+5511999998888'). O `retry-conversion-dispatches` lê
-- essa coluna e entrega o mesmo valor pros dois canais
-- (conversion-dispatch.ts:244), então todo reenvio mandava pro Google um hash
-- que nunca casa — o telefone simplesmente não conta como identificador.
--
-- O fluxo fresco (dispatch-conversions/reconcile) NÃO tinha esse bug: cada
-- dispatcher normalizava o telefone cru vindo da RPC do site do seu jeito.
--
-- SOLUÇÃO
-- Duas colunas, uma por formato. O site já produz as duas variantes em
-- _shared/conversion-hash.ts, e a RPC get_checkout_conversion_data devolve o
-- telefone em E.164 com '+' (site_get_checkout_conversion_data.sql:73-81).

alter table public.conversion_dispatches
  add column if not exists phone_sha256_e164   text,   -- sha256('+55...')  → Google Ads
  add column if not exists phone_sha256_digits text;   -- sha256('55...')   → Meta CAPI

-- Backfill do que dá: o formato antigo é o de dígitos.
update public.conversion_dispatches
set phone_sha256_digits = phone_sha256
where phone_sha256 is not null and phone_sha256_digits is null;

-- ⚠️ phone_sha256_e164 NÃO é recuperável a partir do hash antigo (hash é
-- irreversível e o telefone cru foi dropado na migration 011). Só passa a
-- existir nas linhas gravadas DEPOIS do deploy das Edge Functions. Linhas
-- antigas continuam casando pelo email, que está presente em 100% delas
-- (medido em 28/07/2026: 3439 linhas, 0 sem email_sha256).

-- A coluna `phone_sha256` fica no lugar como legado: ainda é lida pelo Meta
-- (formato compatível) nas linhas anteriores a esta migration. Não dropar
-- antes de todas as linhas com phone_sha256 saírem da janela de retry.

-- ── View pública ────────────────────────────────────────────────────────────
-- OBRIGATÓRIO: `has_phone` da view lia só `phone_sha256`, que as Edge Functions
-- deixam de escrever. Sem este replace, o Tracking Gateway do dashboard
-- (tracking-gateway.html:940 e :1004) mostraria "telefone não enviado" em todas
-- as conversões novas. create or replace só permite ADICIONAR colunas no fim —
-- a lista abaixo repete a ordem da migration 008 e só troca a expressão final.
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
  (phone_sha256_e164 is not null
    or phone_sha256_digits is not null
    or phone_sha256 is not null) as has_phone
from conversion_dispatches;

grant select on conversion_dispatches_public to anon;
