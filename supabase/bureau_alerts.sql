-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_alerts.sql
-- Controle de alertas de gasto do bureau (Assertiva)
--
-- Execute no Supabase: SQL Editor → Run
-- ─────────────────────────────────────────────────────────────────────────────

-- Registra cada alerta enviado para evitar spam de notificações repetidas.
-- A lógica é: só envia o mesmo tipo de alerta uma vez por dia.
CREATE TABLE IF NOT EXISTS bureau_spend_alerts (
  id           BIGSERIAL PRIMARY KEY,
  alert_date   DATE        NOT NULL DEFAULT CURRENT_DATE,
  alert_type   TEXT        NOT NULL,  -- 'daily_spike' | 'monthly_spike' | 'monthly_ok'
  details      JSONB,                 -- dados do momento do alerta (para histórico)
  sent_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (alert_date, alert_type)     -- garante no máximo 1 alerta por tipo por dia
);

CREATE INDEX IF NOT EXISTS idx_bureau_alerts_date
  ON bureau_spend_alerts (alert_date DESC);

ALTER TABLE bureau_spend_alerts ENABLE ROW LEVEL SECURITY;

CREATE POLICY "allow_anon_read"
  ON bureau_spend_alerts
  FOR SELECT TO anon
  USING (true);
