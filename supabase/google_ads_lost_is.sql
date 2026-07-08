-- ─────────────────────────────────────────────────────────────────────────────
-- google_ads_lost_is.sql
-- Adiciona métricas de Impression Share (Search) em google_ads_campaign_daily,
-- para a seção "Lost Impression Share" do dashboard.
--
-- Script reexecutável (ADD COLUMN IF NOT EXISTS). Não altera dados existentes —
-- as colunas novas ficam NULL até o próximo sync-google-ads rodar e preenchê-las.
--
-- Execute no Supabase: SQL Editor → Run (depois de existir a tabela
-- google_ads_campaign_daily, criada originalmente fora deste repo).
-- ─────────────────────────────────────────────────────────────────────────────

ALTER TABLE google_ads_campaign_daily
  ADD COLUMN IF NOT EXISTS search_impression_share NUMERIC(6,4),  -- 0..1 (ex: 0.13 = 13%)
  ADD COLUMN IF NOT EXISTS search_budget_lost_is   NUMERIC(6,4),  -- 0..1, perda por orçamento
  ADD COLUMN IF NOT EXISTS search_rank_lost_is      NUMERIC(6,4); -- 0..1, perda por rank/qualidade

COMMENT ON COLUMN google_ads_campaign_daily.search_impression_share IS
  'metrics.search_impression_share da Google Ads API — só populado para campanhas Search; NULL para Shopping/PMax/Display.';
COMMENT ON COLUMN google_ads_campaign_daily.search_budget_lost_is IS
  'metrics.search_budget_lost_impression_share — fração de impressões elegíveis perdidas por orçamento insuficiente.';
COMMENT ON COLUMN google_ads_campaign_daily.search_rank_lost_is IS
  'metrics.search_rank_lost_impression_share — fração de impressões elegíveis perdidas por rank/qualidade do anúncio.';
