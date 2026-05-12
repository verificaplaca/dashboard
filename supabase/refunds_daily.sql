-- Tabela: refunds_daily
-- Populada por syncRefundsToSupabase() no pagarme-unified.js (Google Apps Script)
-- Lida pelo dashboard.html via supaGet('/rest/v1/refunds_daily')

create table if not exists public.refunds_daily (
  date          date           not null primary key,
  refund_count  integer        not null default 0,
  refund_value  numeric(12, 2) not null default 0
);

-- Libera leitura anônima (dashboard usa anon key)
alter table public.refunds_daily enable row level security;

create policy "anon_read" on public.refunds_daily
  for select using (true);

-- Permite upsert via service_role / anon key com Prefer: resolution=merge-duplicates
-- (o Google Apps Script usa a anon key; se precisar de escrita, use a service_role key)
create policy "anon_upsert" on public.refunds_daily
  for insert with check (true);

create policy "anon_update" on public.refunds_daily
  for update using (true);
