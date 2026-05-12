-- View: refunds_daily
--
-- Agrega estornos a partir da tabela `orders`.
-- Na Pagar.me v5, pedidos estornados ficam com status = 'canceled' e
-- paid_at preenchido (foram pagos antes de serem cancelados/estornados).
-- Usa updated_at como data do estorno (quando o status mudou para canceled).
--
-- Pré-requisito aplicado anteriormente:
--   ALTER TABLE orders ADD CONSTRAINT orders_provider_order_id_key UNIQUE (provider_order_id);

CREATE OR REPLACE VIEW public.refunds_daily AS
SELECT
  date((updated_at AT TIME ZONE 'America/Sao_Paulo')) AS date,
  count(*)                                             AS refund_count,
  round(sum(amount) / 100.0, 2)                        AS refund_value
FROM orders
WHERE lower(status) = 'canceled'
  AND paid_at IS NOT NULL
  AND amount > 0
GROUP BY date((updated_at AT TIME ZONE 'America/Sao_Paulo'))
ORDER BY date((updated_at AT TIME ZONE 'America/Sao_Paulo'));
