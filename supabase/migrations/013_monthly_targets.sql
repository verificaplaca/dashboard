-- 013_monthly_targets.sql
-- Metas mensais do sistema (módulo "Metas" da sidebar do dashboard.html).
-- Rodar no Supabase PRINCIPAL (ftmgmfdqdqxboiktxcoj), SQL Editor.
--
-- Regra de leitura no front (targetFor): usa a linha do mês da data; se não
-- existir, a linha mais recente ANTERIOR (meta vale até ser alterada); se a
-- tabela estiver vazia/inacessível, constantes hardcoded como último fallback.
--
-- Escrita protegida por Supabase Auth (email+senha):
--   1. Authentication → Sign In / Up → desabilitar "Allow new users to sign up"
--   2. Authentication → Users → Add user → criar o usuário do Daniel manualmente
-- Leitura continua anon (dashboard público).

create table monthly_targets (
  month              date primary key,        -- sempre dia 01 (ex.: 2026-07-01)
  target_cpa         numeric,                 -- R$ meta CAC
  target_upsell_pct  numeric,                 -- % meta upsell
  monthly_budget     numeric,                 -- R$ orçamento Google Ads
  revenue_target     numeric,                 -- R$ meta receita
  profit_target      numeric,                 -- R$ meta lucro bruto
  net_profit_target  numeric,                 -- R$ meta lucro líquido
  updated_at         timestamptz default now()
);

alter table monthly_targets enable row level security;

create policy "anon_read"   on monthly_targets for select to anon          using (true);
create policy "auth_read"   on monthly_targets for select to authenticated using (true);
create policy "auth_write"  on monthly_targets for insert to authenticated with check (true);
create policy "auth_update" on monthly_targets for update to authenticated using (true) with check (true);

-- Seed jul/2026 — valores em vigor no dashboard em 15/07/2026
-- (target_cpa = 9, NÃO 10; upsell = 32 confirmado com Daniel em 15/07/2026)
insert into monthly_targets
  (month, target_cpa, target_upsell_pct, monthly_budget, revenue_target, profit_target, net_profit_target)
values
  ('2026-07-01', 9, 32, 89000, 163000, 36000, 30000);
