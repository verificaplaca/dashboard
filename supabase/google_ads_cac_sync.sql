-- ─────────────────────────────────────────────────────────────────────────────
-- google_ads_cac_sync.sql
-- Migration complementar ao google_ads_cac.sql — prepara as tabelas
-- google_ads_search_terms_daily e google_ads_keywords_daily para receber dados
-- reais via Edge Function sync-google-ads-cac (upsert idempotente).
--
-- Script reexecutável: limpa duplicatas antes de criar os UNIQUE INDEX, então
-- pode rodar em banco que já tem dados mock OU em banco novo.
--
-- Execute no Supabase: SQL Editor → Run (depois de google_ads_cac.sql)
-- ─────────────────────────────────────────────────────────────────────────────

-- 1. campaign_id/ad_group_id precisam ser NOT NULL DEFAULT '' (não NULL), porque
--    UNIQUE INDEX não bloqueia duplicatas quando alguma coluna da chave é NULL.
DO $$
BEGIN
  ALTER TABLE google_ads_search_terms_daily ALTER COLUMN campaign_id SET DEFAULT '';
  ALTER TABLE google_ads_search_terms_daily ALTER COLUMN ad_group_id SET DEFAULT '';
  UPDATE google_ads_search_terms_daily SET campaign_id = '' WHERE campaign_id IS NULL;
  UPDATE google_ads_search_terms_daily SET ad_group_id = '' WHERE ad_group_id IS NULL;
  ALTER TABLE google_ads_search_terms_daily ALTER COLUMN campaign_id SET NOT NULL;
  ALTER TABLE google_ads_search_terms_daily ALTER COLUMN ad_group_id SET NOT NULL;

  ALTER TABLE google_ads_keywords_daily ALTER COLUMN campaign_id SET DEFAULT '';
  ALTER TABLE google_ads_keywords_daily ALTER COLUMN ad_group_id SET DEFAULT '';
  UPDATE google_ads_keywords_daily SET campaign_id = '' WHERE campaign_id IS NULL;
  UPDATE google_ads_keywords_daily SET ad_group_id = '' WHERE ad_group_id IS NULL;
  ALTER TABLE google_ads_keywords_daily ALTER COLUMN campaign_id SET NOT NULL;
  ALTER TABLE google_ads_keywords_daily ALTER COLUMN ad_group_id SET NOT NULL;
END $$;

-- 2. Colunas de rastreio do sync (mesma convenção de google_ads_campaign_daily:
--    raw_json para auditoria, ingested_at para saber quando foi sincronizado).
ALTER TABLE google_ads_search_terms_daily ADD COLUMN IF NOT EXISTS raw_json JSONB;
ALTER TABLE google_ads_search_terms_daily ADD COLUMN IF NOT EXISTS ingested_at TIMESTAMPTZ;

ALTER TABLE google_ads_keywords_daily ADD COLUMN IF NOT EXISTS raw_json JSONB;
ALTER TABLE google_ads_keywords_daily ADD COLUMN IF NOT EXISTS ingested_at TIMESTAMPTZ;

-- 3. Limpeza de duplicatas ANTES do unique index (mesmo raciocínio do
--    google_ads_cac.sql para google_ads_optimization_reviews). Mantém a linha
--    mais recente (maior id) por chave composta e remove o resto.
WITH ranked_st AS (
  SELECT
    id,
    ROW_NUMBER() OVER (
      PARTITION BY date, search_term, campaign_id, ad_group_id
      ORDER BY id DESC
    ) AS rn
  FROM google_ads_search_terms_daily
)
DELETE FROM google_ads_search_terms_daily t
USING ranked_st
WHERE t.id = ranked_st.id
  AND ranked_st.rn > 1;

WITH ranked_kw AS (
  SELECT
    id,
    ROW_NUMBER() OVER (
      PARTITION BY date, keyword, campaign_id, ad_group_id
      ORDER BY id DESC
    ) AS rn
  FROM google_ads_keywords_daily
)
DELETE FROM google_ads_keywords_daily k
USING ranked_kw
WHERE k.id = ranked_kw.id
  AND ranked_kw.rn > 1;

-- 4. UNIQUE INDEX — chave usada pelo upsert da Edge Function:
--    onConflict: 'date,search_term,campaign_id,ad_group_id'
--    onConflict: 'date,keyword,campaign_id,ad_group_id'
CREATE UNIQUE INDEX IF NOT EXISTS idx_search_terms_unique
  ON google_ads_search_terms_daily (date, search_term, campaign_id, ad_group_id);

CREATE UNIQUE INDEX IF NOT EXISTS idx_keywords_unique
  ON google_ads_keywords_daily (date, keyword, campaign_id, ad_group_id);

-- 5. Remove o seed MOCK agora que o sync real vai popular as tabelas.
--    Sem WHERE adicional: a V1 só tinha dados de CURRENT_DATE inseridos pelo
--    seed do google_ads_cac.sql, então isso é seguro mesmo rodando mais de uma vez.
DELETE FROM google_ads_search_terms_daily WHERE raw_json IS NULL AND ingested_at IS NULL;
DELETE FROM google_ads_keywords_daily WHERE raw_json IS NULL AND ingested_at IS NULL;
