-- 011_pii_hash.sql — PII hash at rest em conversion_dispatches (P2.3)
-- Projeto do DASHBOARD (ftmgmfdqdqxboiktxcoj). Rodar no SQL Editor desse projeto.
--
-- conversion_dispatches guardava email/phone em claro. Passa a guardar só o
-- SHA-256, mesma normalização usada pelos dispatchers (conversion-dispatch.ts):
--   email: lower(trim(email))            (igual à função sha256Hex do shared)
--   phone: apenas dígitos (regexp_replace \D), igual à normalização do dispatchMeta
-- A view pública (008_conversion_dispatches_public_view.sql) não referencia
-- email/phone — confirmado antes de dropar as colunas, drop é seguro.

create extension if not exists pgcrypto;

alter table public.conversion_dispatches
  add column if not exists email_sha256 text,
  add column if not exists phone_sha256 text;

-- Backfill email (lower+trim, igual sha256Hex no TS).
update public.conversion_dispatches
set email_sha256 = encode(digest(lower(trim(email)), 'sha256'), 'hex')
where email is not null and trim(email) <> '' and email_sha256 is null;

-- Backfill phone (só dígitos antes do hash, igual dispatchMeta no TS).
update public.conversion_dispatches
set phone_sha256 = encode(digest(regexp_replace(phone, '\D', '', 'g'), 'sha256'), 'hex')
where phone is not null
  and regexp_replace(phone, '\D', '', 'g') <> ''
  and phone_sha256 is null;

alter table public.conversion_dispatches drop column if exists email;
alter table public.conversion_dispatches drop column if exists phone;
