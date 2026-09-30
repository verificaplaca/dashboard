-- 016_monthly_targets_tax.sql
-- Imposto (%) sobre a receita, editável no módulo "Metas" do dashboard.html.
-- Rodar no Supabase PRINCIPAL (ftmgmfdqdqxboiktxcoj), SQL Editor.
--
-- Semântica igual às outras metas (targetFor): herança POR CAMPO — vale o
-- tax_pct não-nulo mais recente <= mês; sem nenhum valor → 8% (TAX_PCT_DEFAULT
-- no dashboard e fallback no resumo-diario). Por isso não há seed: todos os
-- meses ficam em 8% até alguém cadastrar outro valor.
-- RLS: as policies da 013 já cobrem a coluna nova.

alter table monthly_targets
  add column if not exists tax_pct numeric
  check (tax_pct is null or (tax_pct >= 0 and tax_pct <= 100));
