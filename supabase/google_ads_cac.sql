-- ─────────────────────────────────────────────────────────────────────────────
-- google_ads_cac.sql
-- Módulo "Google Ads CAC" — análise de search terms e keywords focada em CAC.
-- V1: somente leitura/recomendação. NADA aqui altera o Google Ads automaticamente.
--
-- Script reexecutável (pode rodar em banco novo OU em banco que já recebeu uma
-- versão anterior deste arquivo). Usa CREATE ... IF NOT EXISTS, DROP POLICY IF
-- EXISTS antes de recriar policies, e blocos DO $$ ... $$ para alterações de
-- schema condicionais (constraints, defaults, backfill de dados).
--
-- Execute no Supabase: SQL Editor → Run
-- ─────────────────────────────────────────────────────────────────────────────

-- 1. Lista configurável de termos bloqueados (substitui lista hardcoded na view)
CREATE TABLE IF NOT EXISTS google_ads_blocked_terms (
  id          BIGSERIAL PRIMARY KEY,
  term        TEXT        NOT NULL,
  match_type  TEXT        NOT NULL DEFAULT 'contains', -- 'contains' | 'exact' | 'word'
  category    TEXT,                                     -- ex: 'orgao_publico', 'concorrente', 'preco'
  active      BOOLEAN     NOT NULL DEFAULT true,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Garante que match_type aceita 'word' mesmo em bancos que já tinham a tabela
-- criada com a constraint antiga (só 'contains'/'exact', ou sem constraint nenhuma).
DO $$
BEGIN
  IF EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'chk_blocked_terms_match_type'
  ) THEN
    ALTER TABLE google_ads_blocked_terms DROP CONSTRAINT chk_blocked_terms_match_type;
  END IF;
  ALTER TABLE google_ads_blocked_terms
    ADD CONSTRAINT chk_blocked_terms_match_type CHECK (match_type IN ('contains', 'exact', 'word'));
END $$;

CREATE UNIQUE INDEX IF NOT EXISTS idx_blocked_terms_term
  ON google_ads_blocked_terms (lower(term));

ALTER TABLE google_ads_blocked_terms ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "allow_anon_read" ON google_ads_blocked_terms;
CREATE POLICY "allow_anon_read" ON google_ads_blocked_terms
  FOR SELECT TO anon USING (true);

-- Seed inicial da blocklist (idempotente)
INSERT INTO google_ads_blocked_terms (term, match_type, category) VALUES
  ('detran',       'contains', 'orgao_publico'),
  ('sinesp',       'contains', 'orgao_publico'),
  ('denatran',     'contains', 'orgao_publico'),
  ('olho no carro','contains', 'concorrente'),
  ('gratis',       'contains', 'preco'),
  ('grátis',       'contains', 'preco'),
  ('de graça',     'contains', 'preco'),
  ('graça',        'contains', 'preco'),
  ('mtix',         'contains', 'concorrente'),
  ('up',           'word',     'concorrente'),
  ('gratuita',     'contains', 'preco'),
  ('despachante',  'contains', 'concorrente'),
  ('gringo',       'contains', 'irrelevante'),
  ('buscasim',     'contains', 'concorrente'),
  ('sos',          'contains', 'concorrente'),
  ('checkauto',    'contains', 'concorrente')
ON CONFLICT (lower(term)) DO NOTHING;

-- Corrige match_type de 'up' em bancos onde o seed já rodou com 'contains'
-- (evita falso positivo tipo "upgrade", "grupo" sendo tratado como bloqueado).
UPDATE google_ads_blocked_terms SET match_type = 'word', updated_at = NOW()
  WHERE lower(term) = 'up' AND match_type <> 'word';


-- 2. Revisões manuais (ação tomada internamente sobre uma recomendação)
CREATE TABLE IF NOT EXISTS google_ads_optimization_reviews (
  id              BIGSERIAL PRIMARY KEY,
  entity_type     TEXT        NOT NULL,  -- 'search_term' | 'keyword'
  entity_key      TEXT        NOT NULL,  -- texto do termo/keyword (chave de matching)
  campaign_id     TEXT,
  campaign_name   TEXT,
  ad_group_id     TEXT,
  ad_group_name   TEXT,
  recommendation  TEXT        NOT NULL,  -- snapshot da recomendação no momento da revisão
  action_taken    TEXT        NOT NULL,  -- 'accepted' | 'ignored' | 'reviewed'
  notes           TEXT,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── Backfill defensivo para bancos que já rodaram uma versão anterior ──────
-- campaign_id/ad_group_id precisam ser NOT NULL DEFAULT '' (não NULL), porque
-- UNIQUE/CREATE UNIQUE INDEX não bloqueia duplicatas quando alguma coluna da
-- chave é NULL (NULL nunca é igual a NULL em comparação de unicidade).
DO $$
BEGIN
  -- 1) garante DEFAULT '' nas duas colunas (idempotente — só altera se necessário)
  ALTER TABLE google_ads_optimization_reviews ALTER COLUMN campaign_id SET DEFAULT '';
  ALTER TABLE google_ads_optimization_reviews ALTER COLUMN ad_group_id SET DEFAULT '';

  -- 2) substitui NULL existente por '' antes de aplicar NOT NULL
  UPDATE google_ads_optimization_reviews SET campaign_id = '' WHERE campaign_id IS NULL;
  UPDATE google_ads_optimization_reviews SET ad_group_id = '' WHERE ad_group_id IS NULL;

  -- 3) só agora aplica NOT NULL (a essa altura não existe mais NULL na coluna)
  ALTER TABLE google_ads_optimization_reviews ALTER COLUMN campaign_id SET NOT NULL;
  ALTER TABLE google_ads_optimization_reviews ALTER COLUMN ad_group_id SET NOT NULL;
END $$;

-- Garante os CHECK constraints de entity_type e action_taken, recriando se já existirem
-- com definição antiga/divergente.
DO $$
BEGIN
  IF EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'chk_reviews_entity_type') THEN
    ALTER TABLE google_ads_optimization_reviews DROP CONSTRAINT chk_reviews_entity_type;
  END IF;
  ALTER TABLE google_ads_optimization_reviews
    ADD CONSTRAINT chk_reviews_entity_type CHECK (entity_type IN ('search_term', 'keyword'));

  IF EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'chk_reviews_action_taken') THEN
    ALTER TABLE google_ads_optimization_reviews DROP CONSTRAINT chk_reviews_action_taken;
  END IF;
  ALTER TABLE google_ads_optimization_reviews
    ADD CONSTRAINT chk_reviews_action_taken CHECK (action_taken IN ('accepted', 'ignored', 'reviewed'));
END $$;

-- ── Limpeza de duplicatas ANTES do unique index ────────────────────────────
-- Se o módulo já rodou em V1 sem upsert, pode haver múltiplas linhas para a
-- mesma chave (entity_type, entity_key, campaign_id, ad_group_id). Mantém só
-- a mais recente (maior updated_at; em empate, maior id) e remove o resto.
-- Bloco seguro: roda sempre, não falha se não houver duplicatas (DELETE de 0 linhas).
WITH ranked AS (
  SELECT
    id,
    ROW_NUMBER() OVER (
      PARTITION BY entity_type, entity_key, campaign_id, ad_group_id
      ORDER BY updated_at DESC, id DESC
    ) AS rn
  FROM google_ads_optimization_reviews
)
DELETE FROM google_ads_optimization_reviews r
USING ranked
WHERE r.id = ranked.id
  AND ranked.rn > 1;

-- Status atual = 1 linha por chave composta (entity_type, entity_key, campaign_id,
-- ad_group_id). O front faz upsert via PostgREST:
--   POST .../google_ads_optimization_reviews?on_conflict=entity_type,entity_key,campaign_id,ad_group_id
--   Prefer: resolution=merge-duplicates
-- Isso só funciona se existir exatamente este unique index/constraint nessas 4
-- colunas — por isso a limpeza de duplicatas acima roda antes desta linha.
CREATE UNIQUE INDEX IF NOT EXISTS idx_reviews_unique_entity
  ON google_ads_optimization_reviews (entity_type, entity_key, campaign_id, ad_group_id);

CREATE INDEX IF NOT EXISTS idx_reviews_entity
  ON google_ads_optimization_reviews (entity_type, entity_key);

ALTER TABLE google_ads_optimization_reviews ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "allow_anon_read" ON google_ads_optimization_reviews;
CREATE POLICY "allow_anon_read" ON google_ads_optimization_reviews
  FOR SELECT TO anon USING (true);

-- V1 mantém o padrão já existente no projeto (anon key no front com insert/update
-- liberados). Isso permite que qualquer pessoa com a anon key grave/altere revisões
-- nesta tabela — não há autenticação de usuário por trás. Se isso não for aceitável,
-- restrinja estas policies (ex: exigir role autenticada) antes de ir para produção.
DROP POLICY IF EXISTS "allow_anon_insert" ON google_ads_optimization_reviews;
CREATE POLICY "allow_anon_insert" ON google_ads_optimization_reviews
  FOR INSERT TO anon WITH CHECK (true);

DROP POLICY IF EXISTS "allow_anon_update" ON google_ads_optimization_reviews;
CREATE POLICY "allow_anon_update" ON google_ads_optimization_reviews
  FOR UPDATE TO anon USING (true) WITH CHECK (true);


-- 3. Dados brutos de Search Terms por dia (camada de integração — preenchida por
--    um sync futuro do Google Ads; estrutura pronta para receber dados reais).
CREATE TABLE IF NOT EXISTS google_ads_search_terms_daily (
  id              BIGSERIAL PRIMARY KEY,
  date            DATE        NOT NULL,
  search_term     TEXT        NOT NULL,
  campaign_id     TEXT,
  campaign_name   TEXT        NOT NULL,
  ad_group_id     TEXT,
  ad_group_name   TEXT        NOT NULL,
  clicks          INTEGER     NOT NULL DEFAULT 0,
  impressions     INTEGER     NOT NULL DEFAULT 0,
  cost_micros     BIGINT      NOT NULL DEFAULT 0,   -- custo em micros (1e6 = R$1)
  purchases       NUMERIC(10,2) NOT NULL DEFAULT 0,
  is_existing_keyword BOOLEAN NOT NULL DEFAULT false, -- já existe como keyword ativa?
  created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_search_terms_date ON google_ads_search_terms_daily (date);
CREATE INDEX IF NOT EXISTS idx_search_terms_campaign ON google_ads_search_terms_daily (campaign_name);

ALTER TABLE google_ads_search_terms_daily ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "allow_anon_read" ON google_ads_search_terms_daily;
CREATE POLICY "allow_anon_read" ON google_ads_search_terms_daily
  FOR SELECT TO anon USING (true);


-- 4. Dados brutos de Keywords por dia (mesma lógica de preparação para sync futuro)
CREATE TABLE IF NOT EXISTS google_ads_keywords_daily (
  id                BIGSERIAL PRIMARY KEY,
  date              DATE        NOT NULL,
  keyword           TEXT        NOT NULL,
  match_type        TEXT        NOT NULL DEFAULT 'broad', -- 'broad' | 'phrase' | 'exact'
  campaign_id       TEXT,
  campaign_name     TEXT        NOT NULL,
  ad_group_id       TEXT,
  ad_group_name     TEXT        NOT NULL,
  status_google_ads TEXT        NOT NULL DEFAULT 'ENABLED', -- status real na conta
  clicks            INTEGER     NOT NULL DEFAULT 0,
  impressions       INTEGER     NOT NULL DEFAULT 0,
  cost_micros       BIGINT      NOT NULL DEFAULT 0,
  purchases         NUMERIC(10,2) NOT NULL DEFAULT 0,
  created_at        TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_keywords_date ON google_ads_keywords_daily (date);
CREATE INDEX IF NOT EXISTS idx_keywords_campaign ON google_ads_keywords_daily (campaign_name);

ALTER TABLE google_ads_keywords_daily ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "allow_anon_read" ON google_ads_keywords_daily;
CREATE POLICY "allow_anon_read" ON google_ads_keywords_daily
  FOR SELECT TO anon USING (true);


-- 5. Seed de dados MOCK para testar a UI antes de existir sync real.
--    Remover/ignorar quando o sync real estiver no ar.
INSERT INTO google_ads_search_terms_daily
  (date, search_term, campaign_name, ad_group_name, clicks, impressions, cost_micros, purchases, is_existing_keyword)
VALUES
  (CURRENT_DATE, 'consultar placa veiculo',      'RENAVAM',    'Genérico',     420, 3100, 62000000, 14, false),
  (CURRENT_DATE, 'historico veiculo cpf',        'RISCO_DOC',  'Histórico',    180, 1400, 21000000,  6, false),
  (CURRENT_DATE, 'consulta detran gratis',       'RENAVAM',    'Genérico',      95,  900,  9500000,  0, false),
  (CURRENT_DATE, 'verificar multas placa',       'MULTAS',     'Multas',       260, 2000, 24000000,  3, false),
  (CURRENT_DATE, 'roubo furto consulta veiculo', 'ROUBO_FURTO','Roubo Furto',   60,  500,  4800000,  3, false),
  (CURRENT_DATE, 'sinesp cidadao consulta',      'RENAVAM',    'Genérico',      40,  600,  3000000,  0, false),
  (CURRENT_DATE, 'placa veiculo dados completos','RISCO_DOC',  'Histórico',     30,  220,  2600000,  1, false)
ON CONFLICT DO NOTHING;

INSERT INTO google_ads_keywords_daily
  (date, keyword, match_type, campaign_name, ad_group_name, status_google_ads, clicks, impressions, cost_micros, purchases)
VALUES
  (CURRENT_DATE, 'consultar placa',        'phrase', 'RENAVAM',    'Genérico',  'ENABLED', 510, 4200, 71000000, 18),
  (CURRENT_DATE, 'historico do veiculo',   'phrase', 'RISCO_DOC',  'Histórico', 'ENABLED', 140, 1100, 19000000,  4),
  (CURRENT_DATE, 'consulta multas',        'broad',  'MULTAS',     'Multas',    'ENABLED', 300, 2600, 38000000,  2),
  (CURRENT_DATE, 'roubo e furto veiculo',  'exact',  'ROUBO_FURTO','Roubo Furto','ENABLED',  70,  580,  6700000,  5)
ON CONFLICT DO NOTHING;
