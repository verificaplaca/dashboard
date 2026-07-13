-- P1.4.f — colunas de atribuição em conversion_dispatches (projeto DASHBOARD,
-- ftmgmfdqdqxboiktxcoj). Persistir aqui pra o retry não depender de nova
-- chamada de RPC no site.
alter table public.conversion_dispatches
  add column if not exists gclid text,
  add column if not exists gbraid text,
  add column if not exists wbraid text,
  add column if not exists fbp text,
  add column if not exists fbc text,
  add column if not exists ga_client_id text,
  add column if not exists event_source_url text,
  add column if not exists client_user_agent text;
