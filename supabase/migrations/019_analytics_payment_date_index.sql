-- Preserve all view definitions, payment statuses, dates and revenue formulas.
-- PostgREST filters the computed Sao Paulo payment date, so separate indexes
-- on paid_at/created_at cannot support these scans. Cover the narrow analytics
-- columns to avoid repeatedly scanning the wide orders payload.
-- Apply this file alone (outside a transaction); never db push legacy history.
CREATE INDEX CONCURRENTLY IF NOT EXISTS ix_orders_analytics_payment_date
ON public.orders (((COALESCE(paid_at, created_at) AT TIME ZONE 'America/Sao_Paulo')::date))
INCLUDE (status, amount, paid_at, created_at, provider, provider_order_id);
