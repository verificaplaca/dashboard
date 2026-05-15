-- ─────────────────────────────────────────────────────────────────────────────
-- ads_balance.sql
-- Histórico de saldo Google Ads + controle de alertas Telegram
--
-- Execute no Supabase: SQL Editor → Run
-- ─────────────────────────────────────────────────────────────────────────────

-- 1. Histórico de snapshots de saldo (uma linha por execução do cron)
CREATE TABLE IF NOT EXISTS ads_balance_history (
  id             BIGSERIAL PRIMARY KEY,
  checked_at     TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  balance_brl    NUMERIC(12, 2) NOT NULL,          -- saldo em R$
  alert_level    INTEGER                            -- threshold atingido neste check: 1500, 1000, 500 ou NULL
);

-- Índice para leitura rápida do registro mais recente
CREATE INDEX IF NOT EXISTS idx_ads_balance_checked_at
  ON ads_balance_history (checked_at DESC);

-- 2. View para o dashboard — saldo atual + últimas 30 leituras
CREATE OR REPLACE VIEW ads_balance_latest AS
SELECT
  id,
  checked_at,
  balance_brl,
  alert_level,
  CASE
    WHEN balance_brl <= 500  THEN 'crítico'
    WHEN balance_brl <= 1000 THEN 'atenção'
    WHEN balance_brl <= 1500 THEN 'aviso'
    ELSE                          'ok'
  END AS status
FROM ads_balance_history
ORDER BY checked_at DESC
LIMIT 30;

-- 3. RLS: a Edge Function usa service_role, então não precisa de policy,
--    mas habilitamos RLS para segurança (service_role bypassa automaticamente)
ALTER TABLE ads_balance_history ENABLE ROW LEVEL SECURITY;
