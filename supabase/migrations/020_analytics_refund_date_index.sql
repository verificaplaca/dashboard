-- The parent dashboard reads refunds alongside revenue/upsell. Preserve the
-- refunds_daily predicate and formula, while avoiding another wide orders scan.
-- Apply alone, outside a transaction (CREATE INDEX CONCURRENTLY).
CREATE INDEX CONCURRENTLY IF NOT EXISTS ix_orders_analytics_refund_date
ON public.orders ((date(updated_at AT TIME ZONE 'America/Sao_Paulo')))
INCLUDE (amount, updated_at)
WHERE lower(status) = 'canceled' AND paid_at IS NOT NULL AND amount > 0;
