-- Estornos: pré-requisito para syncRefundsToSupabase()
--
-- A view `refunds_daily` já existe e agrega da tabela `orders`.
-- Este script apenas garante o unique constraint necessário para o upsert
-- via ?on_conflict=provider_order_id funcionar corretamente.
--
-- Execute uma única vez no SQL Editor do Supabase.

alter table public.orders
  add constraint if not exists orders_provider_order_id_key
  unique (provider_order_id);
